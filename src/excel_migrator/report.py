"""Build human-friendly migration reports.

Two outputs:
- Excel report (.xlsx): multi-sheet, KPI-first, color-coded, links to source.
- Markdown summary (.md): concise overview suitable for quick review or Git.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .core import CellAction, SkippedCell, safe_preview
from .images import ImageAction, SkippedImage


@dataclass
class ReportInputs:
    source_path: Path
    template_path: Path
    output_path: Path
    profile: str
    overwrite: bool
    elapsed_seconds: float
    cell_actions: list[CellAction]
    skipped_cells: list[SkippedCell]
    image_actions: list[ImageAction]
    skipped_images: list[SkippedImage]
    formula_fixes: list[str]
    strict_patched_cells: int
    remaining_extra_empty_rows: int


# ---- styling helpers --------------------------------------------------------

_THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

FILL_HEADER = PatternFill("solid", fgColor="2F5496")
FILL_KPI = PatternFill("solid", fgColor="DDEBF7")
FILL_GOOD = PatternFill("solid", fgColor="E2EFDA")
FILL_WARN = PatternFill("solid", fgColor="FFF2CC")
FILL_BAD = PatternFill("solid", fgColor="FCE4D6")
FILL_INFO = PatternFill("solid", fgColor="F2F2F2")

FONT_TITLE = Font(name="Arial", size=18, bold=True, color="1F3864")
FONT_SECTION = Font(name="Arial", size=12, bold=True, color="1F3864")
FONT_HEADER = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FONT_KPI_LABEL = Font(name="Arial", size=11, bold=True, color="1F3864")
FONT_KPI_VALUE = Font(name="Arial", size=20, bold=True, color="1F3864")
FONT_NOTE = Font(name="Arial", size=10, italic=True, color="595959")


def _set_widths(ws: Any, widths: list[float]) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _apply_header(ws: Any, row: int, headers: list[str]) -> None:
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


# ---- report generation -----------------------------------------------------

def write_excel_report(report_path: Path, info: ReportInputs) -> None:
    wb = Workbook()
    _build_summary_sheet(wb.active, info)
    _build_skipped_sheet(wb.create_sheet("需人工确认"), info)
    _build_distribution_sheet(wb.create_sheet("分布与方法"), info)
    _build_image_sheet(wb.create_sheet("图片迁移"), info)
    _build_detail_sheet(wb.create_sheet("单元格明细"), info)
    if info.formula_fixes or info.skipped_images:
        _build_fixes_sheet(wb.create_sheet("修复与异常"), info)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(report_path)


def _build_summary_sheet(ws: Any, info: ReportInputs) -> None:
    ws.title = "总览"
    _set_widths(ws, [3, 22, 22, 22, 22, 22])

    ws["B2"] = "Excel 迁移报告"
    ws["B2"].font = FONT_TITLE
    ws.merge_cells("B2:F2")

    ws["B3"] = (
        f"配置：{'FSC ESF 专用' if info.profile == 'esf' else '通用模式'}  "
        f"｜  覆盖已有内容：{'是' if info.overwrite else '否'}  "
        f"｜  生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  "
        f"｜  耗时：{info.elapsed_seconds:.1f}s"
    )
    ws["B3"].font = FONT_NOTE
    ws.merge_cells("B3:F3")

    # File info block
    ws["B5"] = "旧版来源"
    ws["C5"] = str(info.source_path)
    ws["B6"] = "新版模板"
    ws["C6"] = str(info.template_path)
    ws["B7"] = "输出文件"
    ws["C7"] = str(info.output_path)
    for r in (5, 6, 7):
        ws.cell(row=r, column=2).font = FONT_KPI_LABEL
        ws.cell(row=r, column=2).fill = FILL_INFO
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="right", vertical="center")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)

    # KPI cards
    filled = len(info.cell_actions)
    skipped = len(info.skipped_cells)
    images = len(info.image_actions)
    skipped_imgs = len(info.skipped_images)
    fixes = len(info.formula_fixes)

    total_input = filled + skipped if (filled + skipped) > 0 else 1
    coverage = filled / total_input * 100

    cards = [
        ("已自动填充", filled, "单元格", FILL_GOOD),
        ("覆盖率", f"{coverage:.0f}%", f"共 {total_input} 个目标输入项", FILL_GOOD if coverage >= 80 else FILL_WARN),
        ("需人工确认", skipped, "见「需人工确认」sheet", FILL_WARN if skipped else FILL_GOOD),
        ("迁移图片", images, "另外 {} 张未迁移".format(skipped_imgs) if skipped_imgs else "全部完成", FILL_GOOD if not skipped_imgs else FILL_WARN),
        ("公式修复", fixes, "模板已修补", FILL_INFO),
    ]
    start_row = 9
    for i, (label, value, sub, fill) in enumerate(cards):
        col = 2 + i
        ws.cell(row=start_row, column=col, value=label).font = FONT_KPI_LABEL
        ws.cell(row=start_row, column=col).fill = fill
        ws.cell(row=start_row, column=col).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=start_row, column=col).border = BORDER

        v = ws.cell(row=start_row + 1, column=col, value=value)
        v.font = FONT_KPI_VALUE
        v.fill = fill
        v.alignment = Alignment(horizontal="center", vertical="center")
        v.border = BORDER
        ws.row_dimensions[start_row + 1].height = 36

        s = ws.cell(row=start_row + 2, column=col, value=sub)
        s.font = FONT_NOTE
        s.fill = fill
        s.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        s.border = BORDER
    ws.row_dimensions[start_row].height = 22
    ws.row_dimensions[start_row + 2].height = 28

    # "How to read this report" guidance
    ws["B14"] = "如何阅读这份报告"
    ws["B14"].font = FONT_SECTION
    ws.merge_cells("B14:F14")
    notes = [
        "1. 先看「总览」KPI 数字，判断本次迁移是否成功。",
        "2. 重点查看「需人工确认」sheet，按 sheet 名快速定位需手动核对的字段。",
        "3. 「分布与方法」展示按 sheet 的填充密度，验证各页迁移完成度。",
        "4. 「单元格明细」用于审计，仅在需要追溯具体来源时翻阅。",
        "5. 输出 Excel 已用新版模板的格式、合并区与公式作为骨架，仅写入了可安全确认的内容。",
    ]
    for i, line in enumerate(notes):
        cell = ws.cell(row=15 + i, column=2, value=line)
        cell.alignment = Alignment(vertical="center")
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=6)


def _build_skipped_sheet(ws: Any, info: ReportInputs) -> None:
    _set_widths(ws, [22, 12, 22, 50])
    ws.cell(row=1, column=1, value="未自动填充 / 需人工确认").font = FONT_SECTION
    ws.merge_cells("A1:D1")
    ws.cell(row=2, column=1, value=f"共 {len(info.skipped_cells)} 项。建议在新版输出文件中按下方坐标核对。").font = FONT_NOTE
    ws.merge_cells("A2:D2")

    _apply_header(ws, 4, ["Sheet", "坐标", "原因", "提示"])

    sheet_count = Counter(s.sheet for s in info.skipped_cells)
    row = 5
    for s in info.skipped_cells:
        ws.cell(row=row, column=1, value=s.sheet).border = BORDER
        ws.cell(row=row, column=2, value=s.coord).border = BORDER
        rcell = ws.cell(row=row, column=3, value=s.reason)
        rcell.border = BORDER
        if s.reason.startswith("存在多处"):
            rcell.fill = FILL_BAD
        elif s.reason == "目标已有内容":
            rcell.fill = FILL_INFO
        else:
            rcell.fill = FILL_WARN
        ws.cell(row=row, column=4, value="请打开输出文件并定位到该坐标手动确认。").border = BORDER
        row += 1
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:D{max(row - 1, 4)}"


def _build_distribution_sheet(ws: Any, info: ReportInputs) -> None:
    _set_widths(ws, [28, 14, 14, 14])
    ws.cell(row=1, column=1, value="按 Sheet 的填充分布").font = FONT_SECTION
    ws.merge_cells("A1:D1")
    _apply_header(ws, 3, ["Sheet", "已填充", "需人工确认", "覆盖率"])
    by_sheet_filled = Counter(a.sheet for a in info.cell_actions)
    by_sheet_skipped = Counter(s.sheet for s in info.skipped_cells)
    sheets = sorted(set(by_sheet_filled) | set(by_sheet_skipped))
    row = 4
    for sheet in sheets:
        f = by_sheet_filled.get(sheet, 0)
        s = by_sheet_skipped.get(sheet, 0)
        total = f + s if (f + s) else 1
        ws.cell(row=row, column=1, value=sheet).border = BORDER
        ws.cell(row=row, column=2, value=f).border = BORDER
        ws.cell(row=row, column=3, value=s).border = BORDER
        cov = ws.cell(row=row, column=4, value=f"{f/total*100:.0f}%")
        cov.border = BORDER
        cov.fill = FILL_GOOD if f/total >= 0.8 else (FILL_WARN if f/total >= 0.5 else FILL_BAD)
        row += 1

    # method distribution
    method_start = row + 2
    ws.cell(row=method_start - 1, column=1, value="匹配方法分布").font = FONT_SECTION
    _apply_header(ws, method_start, ["方法", "数量", "说明", ""])
    explanations = {
        "same-coordinate": "旧/新坐标完全一致且上下文相符（最强）",
        "left_section": "依靠左侧标签 + 段落标题命中",
        "rowcol": "依靠左侧标签 + 段落 + 表头联合命中",
        "table": "依靠段落 + 表头 + 行偏移联合命中",
        "left": "依靠左侧标签命中",
        "grounding-part": "FSC ESF 接地表 表 1（按部件）",
        "grounding-carbon-row": "FSC ESF 接地表 表 2（按行序）",
        "grounding-enclosure-row": "FSC ESF 接地表 表 3（按行序）",
    }
    by_method = Counter(a.method for a in info.cell_actions)
    r = method_start + 1
    for method, count in by_method.most_common():
        ws.cell(row=r, column=1, value=method).border = BORDER
        ws.cell(row=r, column=2, value=count).border = BORDER
        ws.cell(row=r, column=3, value=explanations.get(method, method)).border = BORDER
        ws.cell(row=r, column=4, value="").border = BORDER
        r += 1


def _build_image_sheet(ws: Any, info: ReportInputs) -> None:
    _set_widths(ws, [28, 14, 14, 18, 18])
    ws.cell(row=1, column=1, value="图片迁移").font = FONT_SECTION
    ws.merge_cells("A1:E1")
    _apply_header(ws, 3, ["Sheet", "源锚点", "目标锚点", "方法", "尺寸"])
    row = 4
    for a in info.image_actions:
        ws.cell(row=row, column=1, value=a.sheet).border = BORDER
        ws.cell(row=row, column=2, value=a.source_anchor).border = BORDER
        ws.cell(row=row, column=3, value=a.target_anchor).border = BORDER
        ws.cell(row=row, column=4, value=a.method).border = BORDER
        ws.cell(row=row, column=5, value=a.size).border = BORDER
        row += 1
    ws.freeze_panes = "A4"
    if row > 4:
        ws.auto_filter.ref = f"A3:E{row - 1}"


def _build_detail_sheet(ws: Any, info: ReportInputs) -> None:
    _set_widths(ws, [28, 14, 14, 22, 60])
    ws.cell(row=1, column=1, value="单元格迁移明细 (审计用)").font = FONT_SECTION
    ws.merge_cells("A1:E1")
    ws.cell(row=2, column=1, value="此 sheet 仅在需要追溯具体来源时使用。可用筛选/排序定位。").font = FONT_NOTE
    ws.merge_cells("A2:E2")

    _apply_header(ws, 4, ["Sheet", "源坐标", "目标坐标", "方法", "值预览"])
    row = 5
    for a in info.cell_actions:
        ws.cell(row=row, column=1, value=a.sheet).border = BORDER
        ws.cell(row=row, column=2, value=a.source).border = BORDER
        ws.cell(row=row, column=3, value=a.target).border = BORDER
        ws.cell(row=row, column=4, value=a.method).border = BORDER
        ws.cell(row=row, column=5, value=safe_preview(a.value, 200)).border = BORDER
        row += 1
    ws.freeze_panes = "A5"
    if row > 5:
        ws.auto_filter.ref = f"A4:E{row - 1}"


def _build_fixes_sheet(ws: Any, info: ReportInputs) -> None:
    _set_widths(ws, [60, 60])
    ws.cell(row=1, column=1, value="模板公式修复").font = FONT_SECTION
    ws.merge_cells("A1:B1")
    _apply_header(ws, 3, ["修复说明", ""])
    row = 4
    for fix in info.formula_fixes:
        ws.cell(row=row, column=1, value=fix).border = BORDER
        ws.cell(row=row, column=2, value="").border = BORDER
        row += 1
    if not info.formula_fixes:
        ws.cell(row=row, column=1, value="无").border = BORDER
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="未迁移图片").font = FONT_SECTION
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 1
    _apply_header(ws, row, ["Sheet!锚点", "原因"])
    row += 1
    for skip in info.skipped_images:
        ws.cell(row=row, column=1, value=f"{skip.sheet}!{skip.anchor}").border = BORDER
        ws.cell(row=row, column=2, value=skip.reason).border = BORDER
        row += 1


# ---- markdown summary ------------------------------------------------------

def write_markdown_summary(md_path: Path, info: ReportInputs) -> None:
    filled = len(info.cell_actions)
    skipped = len(info.skipped_cells)
    total = filled + skipped if (filled + skipped) else 1
    coverage = filled / total * 100

    by_sheet_filled = Counter(a.sheet for a in info.cell_actions)
    by_sheet_skipped = Counter(s.sheet for s in info.skipped_cells)
    sheets = sorted(set(by_sheet_filled) | set(by_sheet_skipped))

    lines: list[str] = []
    lines.append("# Excel 迁移摘要")
    lines.append("")
    lines.append(f"- 配置：{'FSC ESF 专用' if info.profile == 'esf' else '通用模式'}")
    lines.append(f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"- 耗时：{info.elapsed_seconds:.1f}s")
    lines.append(f"- 旧版来源：`{info.source_path}`")
    lines.append(f"- 新版模板：`{info.template_path}`")
    lines.append(f"- 输出文件：`{info.output_path}`")
    lines.append("")
    lines.append("## 关键指标")
    lines.append("")
    lines.append(f"- 已自动填充：**{filled}**")
    lines.append(f"- 需人工确认：**{skipped}**")
    lines.append(f"- 覆盖率：**{coverage:.0f}%**")
    lines.append(f"- 已迁移图片：{len(info.image_actions)}（未迁移 {len(info.skipped_images)}）")
    lines.append(f"- 模板公式修复：{len(info.formula_fixes)}")
    lines.append("")
    lines.append("## 按 Sheet 分布")
    lines.append("")
    lines.append("| Sheet | 已填充 | 需人工确认 | 覆盖率 |")
    lines.append("|---|---|---|---|")
    for sheet in sheets:
        f = by_sheet_filled.get(sheet, 0)
        s = by_sheet_skipped.get(sheet, 0)
        t = f + s if (f + s) else 1
        lines.append(f"| {sheet} | {f} | {s} | {f/t*100:.0f}% |")
    lines.append("")
    if info.skipped_cells:
        lines.append("## 需人工确认（前 50 条）")
        lines.append("")
        for s in info.skipped_cells[:50]:
            lines.append(f"- `{s.sheet}!{s.coord}` — {s.reason}")
        if len(info.skipped_cells) > 50:
            lines.append(f"- ...等共 {len(info.skipped_cells)} 条，详见 Excel 报告。")
    md_path.parent.mkdir(parents=True, exist_ok=True)
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
