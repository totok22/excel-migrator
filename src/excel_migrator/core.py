"""Core migration logic: cell-level matching with cached worksheet context."""

from __future__ import annotations

import copy
import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Iterable

from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


INPUT_FILL = "FFFFCC99"

PLACEHOLDER_PATTERNS = (
    "fill upon request",
    "需要时填写",
    "部件 parts",
    "if not conductive",
    "n/a if not conductive",
    "如果不是“导电”材料",
    "如果不是”导电”材料",
)

# Instructional text: preserved in template output but treated as "overwritable"
# (i.e. user data can replace them, but they are NOT cleared during template prep)
INSTRUCTIONAL_PATTERNS = (
    "请将图片控制在此单元格范围内",
    "place the picture in this cell",
    "replace texts in this cell",
    "在单元格内填写true",
)

RULE_WORDS = (
    "规则参考",
    "rulesref",
    "ruleref",
    "rulesreference",
)

_NORM_STRIP = re.compile(r"[，。；：、,;:/\\()\[\]{}<>\"'“”‘’\-_\s]+")
_NORM_NUMERIC = re.compile(r"(ev|t|cn|cnonly)?[\d\.]+")
_NORM_WS = re.compile(r"\s+")

# ---- Runtime configuration (set by pipeline before migration) ----

class _MigrationConfig:
    context_threshold: float = 0.55
    fuzzy_threshold: float = 0.70
    image_margin: float = 0.92
    cross_sheet: bool = False
    filter_status: bool = True
    keep_instructional: bool = True

_config = _MigrationConfig()


def set_migration_config(**kwargs) -> None:
    for k, v in kwargs.items():
        setattr(_config, k, v)


# ------------------------- value helpers -------------------------

def is_formula(value: Any) -> bool:
    return (isinstance(value, str) and value.startswith("=")) or type(value).__name__ == "ArrayFormula"


def is_input_cell(cell: Any) -> bool:
    if cell is None or isinstance(cell, MergedCell):
        return False
    fill = cell.fill
    return fill.fill_type == "solid" and fill.fgColor.type == "rgb" and fill.fgColor.rgb == INPUT_FILL


def is_instructional(value: Any) -> bool:
    """Check if value is instructional template text (image placement hints etc.)."""
    if not isinstance(value, str):
        return False
    text = value.strip().lower()
    if not text:
        return False
    return any(pat in text for pat in INSTRUCTIONAL_PATTERNS)


def is_placeholder(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    text = value.strip().lower()
    if not text:
        return True
    if any(pat in text for pat in PLACEHOLDER_PATTERNS):
        return True
    # Instructional text is also considered a placeholder (can be overwritten)
    if any(pat in text for pat in INSTRUCTIONAL_PATTERNS):
        return True
    return False


# Values that are typically formula-generated status indicators and should NOT be migrated.
# Note: TRUE/FALSE are user-entered confirmations in ESF forms and ARE copyable.
_STATUS_VALUES = frozenset({
    "warning", "incomplete", "ok",
    "recommended", "oktext", "incompletetext",
    "n/a",
})


def is_status_value(value: Any) -> bool:
    """Check if a value looks like a formula-generated status indicator."""
    if not isinstance(value, str):
        return False
    text = value.strip().lower()
    return text in _STATUS_VALUES


def copyable_value(cell: Any) -> bool:
    return (
        cell is not None
        and not isinstance(cell, MergedCell)
        and cell.value is not None
        and not is_formula(cell.value)
        and not is_placeholder(cell.value)
    )


def norm_text(value: Any) -> str:
    if value is None or is_formula(value):
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return ""
    text = str(value).strip().lower().replace("`", "")
    text = _NORM_WS.sub("", text)
    text = text.replace("universit", "universi")
    text = _NORM_STRIP.sub("", text)
    if not text:
        return ""
    if any(word in text for word in RULE_WORDS):
        return ""
    if _NORM_NUMERIC.fullmatch(text):
        return ""
    return text


def safe_preview(value: Any, limit: int = 80) -> str:
    text = repr(value)
    return text if len(text) <= limit else text[: limit - 3] + "..."


# ------------------------- data classes -------------------------

@dataclass
class CellAction:
    sheet: str
    source: str
    target: str
    method: str
    value: Any


@dataclass
class SkippedCell:
    sheet: str
    coord: str
    reason: str


# ------------------------- worksheet cache -------------------------

class SheetCache:
    """
    Per-worksheet caches:
        - merged_top: {(row, col) -> (row, col) of the merged-range top-left, or self}
        - static_text_cache: {(row, col) -> normalized static text}
        - left_label_cache: {(row, col) -> normalized left label}
        - header_cache: {(row, col) -> (header_text, header_row)}
        - section_cache: {(row, col, header_row) -> section_title}
        - field_keys_cache: {(row, col) -> tuple of keys}
        - row_static_cache: {row -> [(col, text), ...]}
    """

    __slots__ = (
        "ws",
        "title",
        "max_row",
        "max_column",
        "_cells",
        "_merged_top",
        "_static_cache",
        "_left_label_cache",
        "_header_cache",
        "_section_cache",
        "_field_keys_cache",
        "_row_static_cache",
    )

    def __init__(self, ws: Any) -> None:
        self.ws = ws
        self.title = ws.title
        self.max_row = ws.max_row
        self.max_column = ws.max_column
        self._cells = ws._cells

        self._merged_top: dict[tuple[int, int], tuple[int, int]] = {}
        for merged in ws.merged_cells.ranges:
            top = (merged.min_row, merged.min_col)
            for r in range(merged.min_row, merged.max_row + 1):
                for c in range(merged.min_col, merged.max_col + 1):
                    if (r, c) != top:
                        self._merged_top[(r, c)] = top

        self._static_cache: dict[tuple[int, int], str] = {}
        self._left_label_cache: dict[tuple[int, int], str] = {}
        self._header_cache: dict[tuple[int, int], tuple[str, int | None]] = {}
        self._section_cache: dict[tuple[int, int, int | None], str] = {}
        self._field_keys_cache: dict[tuple[int, int], tuple[tuple[Any, ...], ...]] = {}
        self._row_static_cache: dict[int, list[tuple[int, str]]] = {}

    # --- raw access ---

    def cell(self, row: int, col: int) -> Any:
        if row < 1 or col < 1:
            return None
        return self._cells.get((row, col))

    def existing_cells(self) -> list[Any]:
        return [cell for _, cell in sorted(self._cells.items())]

    def merged_top_left(self, row: int, col: int) -> Any:
        if row < 1 or col < 1 or row > self.max_row or col > self.max_column:
            return None
        top = self._merged_top.get((row, col))
        if top is not None:
            return self._cells.get(top)
        return self._cells.get((row, col))

    # --- cached lookups ---

    def static_text(self, row: int, col: int) -> str:
        key = (row, col)
        cached = self._static_cache.get(key)
        if cached is not None:
            return cached
        cell = self.merged_top_left(row, col)
        if cell is None or isinstance(cell, MergedCell):
            self._static_cache[key] = ""
            return ""
        if is_input_cell(cell):
            self._static_cache[key] = ""
            return ""
        value = cell.value
        if value is None or is_formula(value):
            self._static_cache[key] = ""
            return ""
        if isinstance(value, (int, float, bool)):
            self._static_cache[key] = ""
            return ""
        text = norm_text(value)
        if len(text) > 140:
            text = ""
        self._static_cache[key] = text
        return text

    def left_label(self, row: int, col: int) -> str:
        key = (row, col)
        cached = self._left_label_cache.get(key)
        if cached is not None:
            return cached
        for cc in range(col - 1, 0, -1):
            text = self.static_text(row, cc)
            if text:
                self._left_label_cache[key] = text
                return text
        self._left_label_cache[key] = ""
        return ""

    def nearest_header(self, row: int, col: int, max_up: int = 25) -> tuple[str, int | None]:
        key = (row, col)
        cached = self._header_cache.get(key)
        if cached is not None:
            return cached
        stop = max(1, row - max_up)
        for rr in range(row - 1, stop - 1, -1):
            text = self.static_text(rr, col)
            if text:
                self._header_cache[key] = (text, rr)
                return text, rr
        self._header_cache[key] = ("", None)
        return "", None

    def row_static_texts(self, row: int, start_col: int, end_col: int) -> list[tuple[int, str]]:
        # cache by row only; trim afterwards
        cached = self._row_static_cache.get(row)
        if cached is None:
            out: list[tuple[int, str]] = []
            seen: set[str] = set()
            for col in range(1, self.max_column + 1):
                cell = self.merged_top_left(row, col)
                if cell is None or cell.coordinate in seen:
                    continue
                seen.add(cell.coordinate)
                text = self.static_text(row, col)
                if text:
                    out.append((col, text))
            self._row_static_cache[row] = out
            cached = out
        if start_col <= 1 and end_col >= self.max_column:
            return cached
        return [(c, t) for c, t in cached if start_col <= c <= end_col]

    def section_title(self, row: int, col: int, header_row: int | None) -> str:
        key = (row, col, header_row)
        cached = self._section_cache.get(key)
        if cached is not None:
            return cached
        if header_row is None:
            scan_start = max(1, col - 5)
            scan_end = min(self.max_column, col + 5)
            top = row - 1
        else:
            scan_start = 1
            scan_end = self.max_column
            top = header_row - 1
        title = ""
        for rr in range(top, max(0, top - 12), -1):
            texts = self.row_static_texts(rr, scan_start, scan_end)
            filtered = [(c, t) for c, t in texts if len(t) > 2 and not any(w in t for w in RULE_WORDS)]
            if not filtered:
                continue
            title = sorted(filtered, key=lambda item: item[0])[0][1]
            break
        self._section_cache[key] = title
        return title

    def table_context(self, row: int, col: int) -> tuple[str, str, int | None, int | None]:
        header, header_row = self.nearest_header(row, col)
        if not header or header_row is None:
            return "", "", None, None
        section = self.section_title(row, col, header_row)
        return section, header, header_row, row - header_row

    def field_keys(self, row: int, col: int) -> tuple[tuple[Any, ...], ...]:
        key = (row, col)
        cached = self._field_keys_cache.get(key)
        if cached is not None:
            return cached
        label = self.left_label(row, col)
        section, header, _, offset = self.table_context(row, col)
        keys: list[tuple[Any, ...]] = []
        if label and section:
            keys.append(("left_section", self.title, section, label))
        if label and section and header:
            keys.append(("rowcol", self.title, section, label, header))
        if section and header and offset is not None:
            keys.append(("table", self.title, section, header, offset))
        if label:
            keys.append(("left", self.title, label))
        result = tuple(keys)
        self._field_keys_cache[key] = result
        return result

    def context_tokens(self, row: int, col: int) -> set[str]:
        tokens: set[str] = set()
        label = self.left_label(row, col)
        if label:
            tokens.add("L:" + label)
        section, header, _, offset = self.table_context(row, col)
        if section:
            tokens.add("S:" + section)
        if header:
            tokens.add("H:" + header)
        if offset is not None:
            tokens.add("O:" + str(offset))
        rmin = max(1, row - 3)
        rmax = min(self.max_row, row + 3)
        cmin = max(1, col - 4)
        cmax = min(self.max_column, col + 4)
        for rr in range(rmin, rmax + 1):
            for cc in range(cmin, cmax + 1):
                text = self.static_text(rr, cc)
                if text:
                    tokens.add(text)
        return tokens


# ------------------------- migration logic -------------------------

def _can_write(cell: Any, overwrite: bool) -> bool:
    if cell is None:
        return False
    if isinstance(cell, MergedCell) or is_formula(cell.value):
        return False
    if overwrite:
        return True
    # Allow writing over None, placeholders, and instructional text
    return cell.value is None or is_placeholder(cell.value) or is_instructional(cell.value)


def _copy_cell_value(source: Any, target: Any) -> None:
    target.value = source.value
    if source.hyperlink:
        target._hyperlink = copy.copy(source.hyperlink)
    if source.comment:
        target.comment = copy.copy(source.comment)


def _fuzzy_text_match(a: str, b: str) -> bool:
    """Check if two normalized texts are similar enough (substring or high overlap)."""
    if not a or not b:
        return False
    if a == b:
        return True
    # One contains the other
    if a in b or b in a:
        return True
    # Character-level Jaccard for short texts
    if len(a) < 4 or len(b) < 4:
        return False
    sa, sb = set(a), set(b)
    score = len(sa & sb) / len(sa | sb)
    return score >= _config.fuzzy_threshold


def _context_match(src: SheetCache, src_cell: Any, dst: SheetCache, dst_cell: Any) -> bool:
    s_label = src.left_label(src_cell.row, src_cell.column)
    t_label = dst.left_label(dst_cell.row, dst_cell.column)
    if s_label and t_label:
        if s_label == t_label:
            return True
        # Fuzzy label match for cases like "补充说明" vs "补充说明 Additional Comments"
        if _fuzzy_text_match(s_label, t_label):
            return True
    s_section, s_header, _, s_offset = src.table_context(src_cell.row, src_cell.column)
    t_section, t_header, _, t_offset = dst.table_context(dst_cell.row, dst_cell.column)
    if s_section and t_section and s_header and t_header:
        if s_section == t_section and s_header == t_header and s_offset == t_offset:
            return True
        # Fuzzy section+header match
        if _fuzzy_text_match(s_section, t_section) and _fuzzy_text_match(s_header, t_header) and s_offset == t_offset:
            return True
    s_tokens = src.context_tokens(src_cell.row, src_cell.column)
    t_tokens = dst.context_tokens(dst_cell.row, dst_cell.column)
    if not s_tokens or not t_tokens:
        return False
    score = len(s_tokens & t_tokens) / len(s_tokens | t_tokens)
    return score >= _config.context_threshold


def clear_target_placeholders(wb: Any) -> int:
    """Clear placeholder text from input cells, but PRESERVE instructional text.

    Instructional text (e.g. image placement hints) stays in the template so
    that cells that don't get filled retain their guidance text.
    """
    cleared = 0
    for ws in wb.worksheets:
        for cell in [c for _, c in sorted(ws._cells.items())]:
            if isinstance(cell, MergedCell):
                continue
            if is_input_cell(cell) and not is_formula(cell.value) and is_placeholder(cell.value):
                # Keep instructional text - only clear truly empty placeholders
                if _config.keep_instructional and is_instructional(cell.value):
                    continue
                cell.value = None
                cell._hyperlink = None
                cell.comment = None
                cleared += 1
    return cleared


def build_source_index(
    source_caches: dict[str, SheetCache],
    target_caches: dict[str, SheetCache],
    skip_sheets: set[str],
) -> dict[tuple[Any, ...], list[tuple[SheetCache, Any]]]:
    index: dict[tuple[Any, ...], list[tuple[SheetCache, Any]]] = defaultdict(list)
    for title, src in source_caches.items():
        if title not in target_caches or title in skip_sheets:
            continue
        for cell in src.existing_cells():
            if is_input_cell(cell) and copyable_value(cell):
                for k in src.field_keys(cell.row, cell.column):
                    index[k].append((src, cell))
    return index


def migrate_generic_cells(
    source_caches: dict[str, SheetCache],
    target_caches: dict[str, SheetCache],
    overwrite: bool,
    skip_sheets: set[str],
    progress: callable | None = None,
) -> tuple[list[CellAction], list[SkippedCell]]:
    actions: list[CellAction] = []
    skipped: list[SkippedCell] = []
    filled: set[tuple[str, str]] = set()
    used_sources: set[tuple[str, str]] = set()  # (sheet, coord) - each source cell used at most once
    index = build_source_index(source_caches, target_caches, skip_sheets)

    common_titles = [t for t in target_caches if t in source_caches and t not in skip_sheets]
    total = len(common_titles)

    # Pass 1: same-coordinate matches (highest confidence, consume sources first)
    for i, title in enumerate(common_titles, 1):
        if progress:
            progress(f"精确匹配 {title}", i, total)
        target = target_caches[title]
        source = source_caches[title]
        for tcell in target.existing_cells():
            if not is_input_cell(tcell) or (title, tcell.coordinate) in filled:
                continue
            if not _can_write(tcell, overwrite):
                continue
            scell = source._cells.get((tcell.row, tcell.column))
            if scell is not None and is_input_cell(scell) and copyable_value(scell):
                if _config.filter_status and is_status_value(scell.value):
                    continue
                if _context_match(source, scell, target, tcell):
                    _copy_cell_value(scell, tcell)
                    actions.append(CellAction(title, scell.coordinate, tcell.coordinate, "same-coordinate", scell.value))
                    filled.add((title, tcell.coordinate))
                    used_sources.add((title, scell.coordinate))

    # Pass 2: index-based and fuzzy matches for remaining unfilled cells
    for i, title in enumerate(common_titles, 1):
        if progress:
            progress(f"智能匹配 {title}", i, total)
        target = target_caches[title]
        source = source_caches[title]
        for tcell in target.existing_cells():
            if not is_input_cell(tcell) or (title, tcell.coordinate) in filled:
                continue
            if not _can_write(tcell, overwrite):
                skipped.append(SkippedCell(title, tcell.coordinate, "目标已有内容"))
                continue
            match = None
            method = ""
            ambiguous = False
            for k in target.field_keys(tcell.row, tcell.column):
                candidates = [c for _, c in index.get(k, [])
                              if copyable_value(c)
                              and not (_config.filter_status and is_status_value(c.value))
                              and (title, c.coordinate) not in used_sources]
                if len(candidates) == 1:
                    match = candidates[0]
                    method = k[0]
                    break
                if len(candidates) > 1:
                    unique_by_value = {repr(c.value): c for c in candidates}
                    if len(unique_by_value) == 1:
                        match = next(iter(unique_by_value.values()))
                        method = k[0] + "-same-value"
                        break
                    ambiguous = True

            # Fuzzy fallback: try matching by scanning source cells with similar context
            if match is None and not ambiguous:
                t_label = target.left_label(tcell.row, tcell.column)
                t_section, t_header, _, t_offset = target.table_context(tcell.row, tcell.column)
                if t_label or t_section:
                    for scell in source.existing_cells():
                        if not is_input_cell(scell) or not copyable_value(scell):
                            continue
                        if (title, scell.coordinate) in used_sources:
                            continue
                        if _config.filter_status and is_status_value(scell.value):
                            continue
                        s_label = source.left_label(scell.row, scell.column)
                        if t_label and s_label and _fuzzy_text_match(t_label, s_label):
                            s_section, s_header, _, s_offset = source.table_context(scell.row, scell.column)
                            # Must share section context or be in same relative position
                            if t_section and s_section and _fuzzy_text_match(t_section, s_section):
                                if t_offset == s_offset or (t_header and s_header and _fuzzy_text_match(t_header, s_header)):
                                    match = scell
                                    method = "fuzzy-context"
                                    break

            if match is not None:
                _copy_cell_value(match, tcell)
                actions.append(CellAction(title, match.coordinate, tcell.coordinate, method, match.value))
                filled.add((title, tcell.coordinate))
                used_sources.add((title, match.coordinate))
            elif tcell.value is None or is_placeholder(tcell.value):
                reason = "存在多处可能值，需人工确认" if ambiguous else "未找到安全匹配"
                skipped.append(SkippedCell(title, tcell.coordinate, reason))

    return actions, skipped


def set_recalc_on_open(wb: Any) -> None:
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
