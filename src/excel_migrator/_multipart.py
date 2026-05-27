"""Minimal multipart/form-data parser for Python 3.13+.

`cgi` was removed in 3.13. This parser streams the body and writes file fields
to disk so multi-MB uploads do not blow up memory.
"""

from __future__ import annotations

import io
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import IO


_BOUNDARY_RE = re.compile(rb'boundary="?([^";]+)"?', re.IGNORECASE)
_NAME_RE = re.compile(rb'name="([^"]*)"', re.IGNORECASE)
_FILENAME_RE = re.compile(rb'filename="([^"]*)"', re.IGNORECASE)


@dataclass
class FormFile:
    name: str
    filename: str
    path: Path
    size: int


@dataclass
class FormData:
    fields: dict[str, str]
    files: dict[str, FormFile]


class _Scanner:
    """Streaming scanner with a small accumulator and lookahead."""

    def __init__(self, stream: IO[bytes], chunk_size: int = 64 * 1024) -> None:
        self.stream = stream
        self.chunk_size = chunk_size
        self.buf = bytearray()

    def _ensure(self, n: int) -> None:
        while len(self.buf) < n:
            chunk = self.stream.read(self.chunk_size)
            if not chunk:
                return
            self.buf.extend(chunk)

    def read_exact(self, n: int) -> bytes:
        self._ensure(n)
        if len(self.buf) < n:
            data = bytes(self.buf)
            self.buf.clear()
            return data
        data = bytes(self.buf[:n])
        del self.buf[:n]
        return data

    def read_until(self, marker: bytes, sink: IO[bytes] | None) -> None:
        """Read from stream until `marker` is found. Bytes before marker go to sink.

        After this call the marker has been consumed; the next read picks up
        immediately after the marker.
        """
        keep = len(marker) - 1  # we may need to keep this many bytes between reads
        while True:
            idx = self.buf.find(marker)
            if idx >= 0:
                if sink is not None and idx > 0:
                    sink.write(bytes(self.buf[:idx]))
                del self.buf[: idx + len(marker)]
                return
            # No match yet. Flush most of the buffer (keep last `keep` bytes for
            # the next iteration in case the marker straddles a boundary).
            if sink is not None and len(self.buf) > keep:
                sink.write(bytes(self.buf[:-keep]))
                del self.buf[:-keep]
            elif sink is None and len(self.buf) > keep:
                del self.buf[:-keep]
            chunk = self.stream.read(self.chunk_size)
            if not chunk:
                # marker never appeared
                if sink is not None and self.buf:
                    sink.write(bytes(self.buf))
                self.buf.clear()
                raise EOFError("multipart marker not found")
            self.buf.extend(chunk)


def parse(rfile: IO[bytes], content_type: str, work_dir: Path) -> FormData:
    work_dir.mkdir(parents=True, exist_ok=True)
    m = _BOUNDARY_RE.search(content_type.encode("latin-1"))
    if not m:
        raise ValueError("missing boundary in Content-Type")
    boundary = m.group(1)
    delim = b"--" + boundary

    scanner = _Scanner(rfile)

    # Skip preamble up to first boundary
    scanner.read_until(delim, None)
    # Then either CRLF (more parts) or "--" (terminal)
    suffix = scanner.read_exact(2)
    if suffix == b"--":
        return FormData(fields={}, files={})

    fields: dict[str, str] = {}
    files: dict[str, FormFile] = {}

    while True:
        head_buf = io.BytesIO()
        scanner.read_until(b"\r\n\r\n", head_buf)
        head_bytes = head_buf.getvalue()

        disposition = b""
        for line in head_bytes.split(b"\r\n"):
            if line.lower().startswith(b"content-disposition:"):
                disposition = line
                break

        name_m = _NAME_RE.search(disposition)
        file_m = _FILENAME_RE.search(disposition)
        body_marker = b"\r\n" + delim

        if name_m and file_m:
            name = name_m.group(1).decode("utf-8", "replace")
            filename = file_m.group(1).decode("utf-8", "replace")
            target = work_dir / _safe_name(filename)
            with open(target, "wb") as sink:
                scanner.read_until(body_marker, sink)
            files[name] = FormFile(name=name, filename=filename, path=target, size=target.stat().st_size)
        elif name_m:
            name = name_m.group(1).decode("utf-8", "replace")
            buf = io.BytesIO()
            scanner.read_until(body_marker, buf)
            fields[name] = buf.getvalue().decode("utf-8", "replace")
        else:
            scanner.read_until(body_marker, None)

        suffix = scanner.read_exact(2)
        if suffix == b"--":
            break
        # CRLF -> next part
    return FormData(fields=fields, files=files)


def _safe_name(name: str) -> str:
    base = os.path.basename(name) or "upload.bin"
    base = base.replace("\x00", "_")
    return base
