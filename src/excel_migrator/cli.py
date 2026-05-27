"""CLI entry point."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .pipeline import MigrationOptions, run


def _progress(stage: str, current: int, total: int) -> None:
    print(f"[{current}/{total}] {stage}", flush=True)


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="excel-migrator",
        description="把旧 Excel 内容迁移到新模板。默认通用模式，可加 --profile esf 启用 FSC ESF 专用规则。",
    )
    p.add_argument("source", help="旧版 Excel 文件路径")
    p.add_argument("template", help="新版模板 Excel 路径")
    p.add_argument("-o", "--output", default=None, help="输出 xlsx 路径（默认：旧文件同目录的 _migrated.xlsx）")
    p.add_argument("-r", "--report", default=None, help="Excel 报告路径（默认：输出同目录的 _报告.xlsx）")
    p.add_argument("-m", "--markdown", default=None, help="可选的 Markdown 摘要路径")
    p.add_argument("--profile", choices=("generic", "esf"), default="generic", help="迁移配置")
    p.add_argument("--overwrite", action="store_true", help="覆盖新版模板中已有的非占位内容")
    p.add_argument("--no-images", action="store_true", help="不迁移图片")
    p.add_argument("--keep-template-images", action="store_true", help="保留新模板中原有的图片")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    source = Path(args.source).expanduser().resolve()
    template = Path(args.template).expanduser().resolve()
    if not source.exists():
        print(f"找不到旧版文件：{source}", file=sys.stderr)
        return 2
    if not template.exists():
        print(f"找不到新版模板：{template}", file=sys.stderr)
        return 2

    output = Path(args.output).expanduser().resolve() if args.output else source.with_name(source.stem + "_migrated.xlsx")
    report = Path(args.report).expanduser().resolve() if args.report else output.with_name(output.stem + "_报告.xlsx")
    markdown = Path(args.markdown).expanduser().resolve() if args.markdown else None

    opts = MigrationOptions(
        source=source,
        template=template,
        output=output,
        excel_report=report,
        markdown_report=markdown,
        profile=args.profile,
        overwrite=args.overwrite,
        include_images=not args.no_images,
        keep_template_images=args.keep_template_images,
    )
    result = run(opts, progress=_progress)
    print()
    print(f"完成：已填充 {len(result.cell_actions)} 个单元格，迁移图片 {len(result.image_actions)} 张")
    print(f"输出：{result.output}")
    print(f"报告：{result.excel_report}")
    if result.markdown_report:
        print(f"摘要：{result.markdown_report}")
    print(f"耗时：{result.elapsed_seconds:.1f} 秒")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
