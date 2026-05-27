"""Local web server.

Pure stdlib http.server with:
- File uploads via multipart/form-data
- Server-sent events (SSE) for progress
- Local file system only — nothing is sent over the network
"""

from __future__ import annotations

import argparse
import io
import json
import mimetypes
import shutil
import sys
import tempfile
import threading
import time
import traceback
import uuid
import webbrowser
from collections import deque
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from .pipeline import MigrationOptions, run
from . import _multipart


WEB_ROOT = Path(__file__).resolve().parent.parent.parent / "web"
WORK_ROOT = Path(tempfile.gettempdir()) / "excel-migrator-jobs"
WORK_ROOT.mkdir(parents=True, exist_ok=True)


# -------------------- job tracking --------------------

class Job:
    def __init__(self, jid: str, work_dir: Path) -> None:
        self.id = jid
        self.work_dir = work_dir
        self.events: deque[dict] = deque()
        self.lock = threading.Lock()
        self.done = threading.Event()
        self.subscribers: list[threading.Event] = []
        self.result_paths: dict[str, str] = {}

    def emit(self, payload: dict) -> None:
        with self.lock:
            self.events.append(payload)
            for s in self.subscribers:
                s.set()

    def add_subscriber(self) -> threading.Event:
        ev = threading.Event()
        with self.lock:
            self.subscribers.append(ev)
            if self.events:
                ev.set()
        return ev


JOBS: dict[str, Job] = {}
JOBS_LOCK = threading.Lock()


class _LimitedReader:
    """Reads at most `length` bytes from the underlying stream."""

    def __init__(self, stream, length: int) -> None:
        self.stream = stream
        self.remaining = length

    def read(self, n: int = -1) -> bytes:
        if self.remaining <= 0:
            return b""
        if n < 0 or n > self.remaining:
            n = self.remaining
        data = self.stream.read(n)
        self.remaining -= len(data)
        return data


def _new_job() -> Job:
    jid = uuid.uuid4().hex[:12]
    work = WORK_ROOT / jid
    work.mkdir(parents=True, exist_ok=True)
    job = Job(jid, work)
    with JOBS_LOCK:
        JOBS[jid] = job
    return job


def _get_job(jid: str) -> Job | None:
    with JOBS_LOCK:
        return JOBS.get(jid)


# -------------------- worker --------------------

def _run_job(job: Job, opts: MigrationOptions) -> None:
    def progress(stage: str, current: int, total: int) -> None:
        job.emit({"type": "progress", "stage": stage, "current": current, "total": total})

    try:
        job.emit({"type": "log", "message": "开始处理..."})
        result = run(opts, progress=progress)
        job.result_paths = {
            "output": str(result.output),
            "excel_report": str(result.excel_report),
        }
        if result.markdown_report:
            job.result_paths["markdown_report"] = str(result.markdown_report)
        job.emit({
            "type": "done",
            "filled": len(result.cell_actions),
            "skipped": len(result.skipped_cells),
            "images": len(result.image_actions),
            "elapsed": round(result.elapsed_seconds, 1),
            "downloads": [
                {"name": "新版输出 Excel", "key": "output", "filename": opts.output.name},
                {"name": "迁移报告 (Excel)", "key": "excel_report", "filename": opts.excel_report.name},
            ] + ([
                {"name": "迁移摘要 (Markdown)", "key": "markdown_report", "filename": opts.markdown_report.name}
            ] if opts.markdown_report else []),
        })
    except Exception:
        tb = traceback.format_exc()
        job.emit({"type": "error", "message": tb})
    finally:
        job.done.set()


# -------------------- request handler --------------------

class Handler(BaseHTTPRequestHandler):
    server_version = "ExcelMigrator/0.1"

    def log_message(self, fmt: str, *args) -> None:  # noqa: A003 - quieter logging
        sys.stderr.write("[server] " + fmt % args + "\n")

    # ---- helpers ----

    def _send_json(self, status: HTTPStatus, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: Path, download_name: str | None = None) -> None:
        if not path.exists() or not path.is_file():
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "file not found"})
            return
        ctype, _ = mimetypes.guess_type(str(path))
        ctype = ctype or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(path.stat().st_size))
        if download_name:
            from urllib.parse import quote
            self.send_header(
                "Content-Disposition",
                f"attachment; filename*=UTF-8''{quote(download_name)}",
            )
        self.end_headers()
        with open(path, "rb") as f:
            shutil.copyfileobj(f, self.wfile)

    # ---- routing ----

    def do_GET(self) -> None:  # noqa: N802
        url = urlparse(self.path)
        path = url.path

        if path == "/" or path == "/index.html":
            self._send_static("index.html")
            return
        if path.startswith("/web/"):
            self._send_static(path[len("/web/"):])
            return
        if path == "/api/job/events":
            self._api_events(parse_qs(url.query))
            return
        if path == "/api/job/download":
            self._api_download(parse_qs(url.query))
            return

        self._send_json(HTTPStatus.NOT_FOUND, {"error": "not found"})

    def do_POST(self) -> None:  # noqa: N802
        url = urlparse(self.path)
        if url.path == "/api/migrate":
            self._api_migrate()
            return
        self._send_json(HTTPStatus.NOT_FOUND, {"error": "not found"})

    # ---- static ----

    def _send_static(self, rel: str) -> None:
        target = (WEB_ROOT / rel).resolve()
        try:
            target.relative_to(WEB_ROOT.resolve())
        except ValueError:
            self._send_json(HTTPStatus.FORBIDDEN, {"error": "forbidden"})
            return
        if target.is_dir():
            target = target / "index.html"
        if not target.exists():
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "not found"})
            return
        self._send_file(target)

    # ---- API ----

    def _api_migrate(self) -> None:
        ctype = self.headers.get("Content-Type", "")
        if not ctype.startswith("multipart/form-data"):
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": "expected multipart/form-data"})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0

        job = _new_job()
        upload_dir = job.work_dir / "upload"
        try:
            limited = _LimitedReader(self.rfile, content_length) if content_length else self.rfile
            form = _multipart.parse(limited, ctype, upload_dir)
        except Exception as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": f"failed to parse form: {exc}"})
            return

        profile = (form.fields.get("profile") or "generic").strip()
        overwrite = form.fields.get("overwrite") == "1"
        no_images = form.fields.get("no_images") == "1"
        keep_template_images = form.fields.get("keep_template_images") == "1"
        emit_markdown = form.fields.get("markdown") == "1"

        # Advanced settings
        def _float(key: str, default: float) -> float:
            try:
                return float(form.fields.get(key, str(default)))
            except (ValueError, TypeError):
                return default

        context_threshold = _float("context_threshold", 0.55)
        fuzzy_threshold = _float("fuzzy_threshold", 0.70)
        image_margin = _float("image_margin", 0.92)
        cross_sheet = form.fields.get("cross_sheet") == "1"
        filter_status = form.fields.get("filter_status", "1") == "1"
        keep_instructional = form.fields.get("keep_instructional", "1") == "1"

        def resolve_input(field_name: str) -> Path | None:
            ff = form.files.get(field_name)
            if ff is not None and ff.size > 0:
                return ff.path
            return None

        source = resolve_input("source")
        template = resolve_input("template")

        if not source or not template:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": "缺少旧版文件或新模板"})
            return

        out_dir = job.work_dir / "out"
        out_dir.mkdir(exist_ok=True)
        output = out_dir / (source.stem + "_migrated.xlsx")
        excel_report = out_dir / (source.stem + "_迁移报告.xlsx")
        markdown_report = out_dir / (source.stem + "_迁移摘要.md") if emit_markdown else None

        opts = MigrationOptions(
            source=source,
            template=template,
            output=output,
            excel_report=excel_report,
            markdown_report=markdown_report,
            profile=profile if profile in ("generic", "esf") else "generic",
            overwrite=overwrite,
            include_images=not no_images,
            keep_template_images=keep_template_images,
            context_threshold=context_threshold,
            fuzzy_threshold=fuzzy_threshold,
            image_margin=image_margin,
            cross_sheet=cross_sheet,
            filter_status=filter_status,
            keep_instructional=keep_instructional,
        )

        threading.Thread(target=_run_job, args=(job, opts), daemon=True).start()
        self._send_json(HTTPStatus.OK, {"job_id": job.id})

    def _api_events(self, params: dict) -> None:
        jid = (params.get("job_id") or [""])[0]
        job = _get_job(jid)
        if not job:
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "unknown job"})
            return
        ev = job.add_subscriber()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("X-Accel-Buffering", "no")
        self.end_headers()
        cursor = 0
        try:
            while True:
                ev.wait(timeout=15)
                ev.clear()
                with job.lock:
                    pending = list(job.events)[cursor:]
                    cursor = len(job.events)
                for payload in pending:
                    data = json.dumps(payload, ensure_ascii=False)
                    try:
                        self.wfile.write(f"data: {data}\n\n".encode("utf-8"))
                        self.wfile.flush()
                    except (BrokenPipeError, ConnectionResetError):
                        return
                if job.done.is_set():
                    # send a heartbeat/closing event then exit
                    try:
                        self.wfile.write(b": end\n\n")
                        self.wfile.flush()
                    except Exception:
                        pass
                    return
                # heartbeat to keep proxies happy
                try:
                    self.wfile.write(b": ping\n\n")
                    self.wfile.flush()
                except (BrokenPipeError, ConnectionResetError):
                    return
        except Exception:
            return

    def _api_download(self, params: dict) -> None:
        jid = (params.get("job_id") or [""])[0]
        key = (params.get("key") or [""])[0]
        job = _get_job(jid)
        if not job or key not in job.result_paths:
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "not found"})
            return
        path = Path(job.result_paths[key])
        self._send_file(path, download_name=path.name)


# -------------------- entry point --------------------

def serve(host: str = "127.0.0.1", port: int = 8765, open_browser: bool = True) -> None:
    server = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"\n  Excel Migrator 已启动：{url}\n  按 Ctrl+C 关闭\n", flush=True)
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  正在关闭...", flush=True)
        server.shutdown()


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="excel-migrator-server")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=8765)
    p.add_argument("--no-browser", action="store_true")
    args = p.parse_args(argv)
    serve(host=args.host, port=args.port, open_browser=not args.no_browser)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
