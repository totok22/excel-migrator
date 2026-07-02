"""Strict template write-back.

Use the new template's worksheet XML as the structural base, only patching cell
values from the staging workbook. Keeps formatting/merged cells/validations
intact and lets us trim empty rows that openpyxl introduced.
"""

from __future__ import annotations

import copy
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {
    "main": SHEET_NS,
    "rel": OFFICE_REL_NS,
    "pkgrel": PKG_REL_NS,
}

# Register ALL namespace prefixes that Excel uses so that ET.tostring preserves
# the correct prefix names. Without this, ET renames them to ns0/ns1/ns2 which
# breaks mc:Ignorable references and causes Excel to report corruption.
ET.register_namespace("", SHEET_NS)
ET.register_namespace("r", OFFICE_REL_NS)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
ET.register_namespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision6")
ET.register_namespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision10")
ET.register_namespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
ET.register_namespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
ET.register_namespace("xm", "http://schemas.microsoft.com/office/excel/2006/main")


# Lower compresslevel makes a noticeable difference for large image-heavy xlsx.
ZIP_LEVEL = 1

# Mapping of namespace prefixes commonly used in Excel OOXML worksheets.
# ET.tostring only emits xmlns declarations for prefixes actually used in the tree,
# but mc:Ignorable may reference prefixes whose elements were stripped. We must
# ensure those declarations are present or Excel reports the file as corrupt.
_EXCEL_NS_URIS: dict[str, str] = {
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision10",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
}


import re as _re

_IGNORABLE_RE = _re.compile(rb'Ignorable="([^"]+)"')
_ROOT_TAG_RE = _re.compile(rb'(<worksheet\b[^>]*)(>)')


def _ensure_ignorable_ns_declared(xml_bytes: bytes) -> bytes:
    """Inject missing namespace declarations for prefixes referenced in mc:Ignorable.

    Python's ElementTree only serializes xmlns declarations for prefixes that are
    actively used in element/attribute names within the tree. However, Excel's
    mc:Ignorable attribute may list prefixes (like x14ac, xr2, xr3) whose elements
    were removed during processing. If those xmlns declarations are missing, Excel
    considers the file corrupt.
    """
    m = _IGNORABLE_RE.search(xml_bytes)
    if not m:
        return xml_bytes

    ignorable_prefixes = m.group(1).decode("utf-8").split()
    missing_decls: list[str] = []
    for prefix in ignorable_prefixes:
        # Check if xmlns:prefix= is already declared
        decl_pattern = f'xmlns:{prefix}='.encode("utf-8")
        if decl_pattern not in xml_bytes:
            uri = _EXCEL_NS_URIS.get(prefix)
            if uri:
                missing_decls.append(f' xmlns:{prefix}="{uri}"')

    if not missing_decls:
        return xml_bytes

    # Insert missing declarations into the root <worksheet ...> tag
    def _inject(match: _re.Match) -> bytes:
        return match.group(1) + "".join(missing_decls).encode("utf-8") + match.group(2)

    return _ROOT_TAG_RE.sub(_inject, xml_bytes, count=1)


def _norm_target(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target.lstrip("/")


def workbook_sheet_paths(xlsx_path: Path) -> dict[str, str]:
    with ZipFile(xlsx_path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_paths = {
        rel.attrib["Id"]: _norm_target(rel.attrib["Target"])
        for rel in rels.findall("pkgrel:Relationship", NS)
    }
    sheet_paths: dict[str, str] = {}
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        rel_id = sheet.attrib[f"{{{OFFICE_REL_NS}}}id"]
        sheet_paths[sheet.attrib["name"]] = rel_paths[rel_id]
    return sheet_paths


def empty_row_numbers(xlsx_path: Path, sheet_path: str) -> set[int]:
    with ZipFile(xlsx_path) as zf:
        root = ET.fromstring(zf.read(sheet_path))
    rows: set[int] = set()
    for row in root.findall("main:sheetData/main:row", NS):
        if not row.findall("main:c", NS):
            rows.add(int(row.attrib["r"]))
    return rows


_CELL_REF_RE = re.compile(r"([A-Z]+)([0-9]+)")


def _col_to_index(col: str) -> int:
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _cell_sort_key(ref: str) -> tuple[int, int]:
    match = _CELL_REF_RE.fullmatch(ref)
    if not match:
        return (0, 0)
    col, row = match.groups()
    return (int(row), _col_to_index(col))


def _is_empty_style_cell(cell: ET.Element) -> bool:
    return not _value_children(cell)


def _remap_layout_style_attrs(root: ET.Element, style_remap: dict[int, int]) -> None:
    if not style_remap:
        return
    for col in root.findall("main:cols/main:col", NS):
        style = col.attrib.get("style")
        if style is not None:
            col.attrib["style"] = str(style_remap.get(int(style), int(style)))
    for row in root.findall("main:sheetData/main:row", NS):
        style = row.attrib.get("s")
        if style is not None:
            row.attrib["s"] = str(style_remap.get(int(style), int(style)))


def _restore_empty_cells_xml(
    reference_xml: bytes,
    workbook_xml: bytes,
    layout_xml: bytes | None = None,
    style_remap: dict[int, int] | None = None,
) -> tuple[bytes, int]:
    reference_root = ET.fromstring(reference_xml)
    workbook_root = ET.fromstring(workbook_xml)
    layout_root = ET.fromstring(layout_xml) if layout_xml is not None else reference_root
    if style_remap:
        _remap_layout_style_attrs(layout_root, style_remap)

    ref_cols = layout_root.find("main:cols", NS)
    workbook_cols = workbook_root.find("main:cols", NS)
    if workbook_cols is not None:
        workbook_root.remove(workbook_cols)
    if ref_cols is not None:
        sheet_data_pos = next(
            (idx for idx, child in enumerate(list(workbook_root)) if child.tag == f"{{{SHEET_NS}}}sheetData"),
            len(list(workbook_root)),
        )
        workbook_root.insert(sheet_data_pos, copy.deepcopy(ref_cols))

    sheet_data = workbook_root.find("main:sheetData", NS)
    reference_sheet_data = reference_root.find("main:sheetData", NS)
    if sheet_data is None or reference_sheet_data is None:
        return workbook_xml, 0

    layout_sheet_data = layout_root.find("main:sheetData", NS)
    layout_rows_by_number = {
        row.attrib.get("r"): row
        for row in (layout_sheet_data.findall("main:row", NS) if layout_sheet_data is not None else [])
        if row.attrib.get("r")
    }
    for row in sheet_data.findall("main:row", NS):
        row_num = row.attrib.get("r")
        ref_row = layout_rows_by_number.get(row_num)
        if ref_row is not None:
            row.attrib.clear()
            row.attrib.update(ref_row.attrib)

    existing = {
        cell.attrib.get("r")
        for cell in sheet_data.findall("main:row/main:c", NS)
        if cell.attrib.get("r")
    }

    rows_by_number = {
        row.attrib.get("r"): row
        for row in sheet_data.findall("main:row", NS)
        if row.attrib.get("r")
    }

    restored = 0
    for ref_row in reference_sheet_data.findall("main:row", NS):
        row_num = ref_row.attrib.get("r")
        if not row_num:
            continue
        row = rows_by_number.get(row_num)
        if row is None:
            row = copy.deepcopy(layout_rows_by_number.get(row_num, ref_row))
            for cell in list(row):
                row.remove(cell)
            for cell in ref_row.findall("main:c", NS):
                row.append(copy.deepcopy(cell))
            for cell in list(row):
                if not _is_empty_style_cell(cell):
                    row.remove(cell)
            if not list(row):
                continue
            sheet_data.append(row)
            rows_by_number[row_num] = row
            restored += len(row.findall("main:c", NS))
            continue

        added = False
        for ref_cell in ref_row.findall("main:c", NS):
            ref = ref_cell.attrib.get("r")
            if not ref or ref in existing or not _is_empty_style_cell(ref_cell):
                continue
            row.append(copy.deepcopy(ref_cell))
            existing.add(ref)
            restored += 1
            added = True
        if added:
            cells = sorted(row.findall("main:c", NS), key=lambda c: _cell_sort_key(c.attrib.get("r", "")))
            for cell in list(row):
                row.remove(cell)
            for cell in cells:
                row.append(cell)

    if not restored:
        return workbook_xml, 0

    rows = sorted(sheet_data.findall("main:row", NS), key=lambda r: int(r.attrib.get("r", "0")))
    for row in list(sheet_data):
        sheet_data.remove(row)
    for row in rows:
        sheet_data.append(row)

    return _ensure_ignorable_ns_declared(
        ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)
    ), restored


def restore_empty_cells_from_reference(
    reference_path: Path,
    workbook_path: Path,
    layout_path: Path | None = None,
) -> int:
    """Restore empty style-only cells dropped by workbook normalization."""
    reference_sheets = workbook_sheet_paths(reference_path)
    workbook_sheets = workbook_sheet_paths(workbook_path)
    layout_sheets = workbook_sheet_paths(layout_path) if layout_path else reference_sheets
    sheet_paths = {
        workbook_sheets[name]: reference_sheets[name]
        for name in workbook_sheets
        if name in reference_sheets
    }
    layout_sheet_paths = {
        workbook_sheets[name]: layout_sheets[name]
        for name in workbook_sheets
        if name in layout_sheets
    }

    style_remap: dict[int, int] = {}

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmpf:
        temp_path = Path(tmpf.name)
    restored = 0
    try:
        with ZipFile(reference_path) as ref_zip, ZipFile(workbook_path) as src, ZipFile(
            temp_path, "w", ZIP_DEFLATED, compresslevel=ZIP_LEVEL
        ) as dst:
            layout_zip_ctx = ZipFile(layout_path) if layout_path else None
            try:
                for info in src.infolist():
                    data = src.read(info.filename)
                    ref_sheet_path = sheet_paths.get(info.filename)
                    if ref_sheet_path:
                        layout_xml = (
                            layout_zip_ctx.read(layout_sheet_paths[info.filename])
                            if layout_zip_ctx is not None and info.filename in layout_sheet_paths
                            else None
                        )
                        data, count = _restore_empty_cells_xml(
                            ref_zip.read(ref_sheet_path),
                            data,
                            layout_xml=layout_xml,
                            style_remap=style_remap,
                        )
                        restored += count
                    dst.writestr(info, data)
            finally:
                if layout_zip_ctx is not None:
                    layout_zip_ctx.close()
        shutil.move(str(temp_path), workbook_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return restored


def extra_empty_rows_not_in_template(template_path: Path, workbook_path: Path) -> dict[str, set[int]]:
    template_sheets = workbook_sheet_paths(template_path)
    output_sheets = workbook_sheet_paths(workbook_path)
    extra: dict[str, set[int]] = {}
    for sheet_name, output_sheet_path in output_sheets.items():
        template_sheet_path = template_sheets.get(sheet_name)
        if template_sheet_path is None:
            continue
        t_empty = empty_row_numbers(template_path, template_sheet_path)
        o_empty = empty_row_numbers(workbook_path, output_sheet_path)
        diff = o_empty - t_empty
        if diff:
            extra[output_sheet_path] = diff
    return extra


def prune_rows_not_in_template(template_path: Path, workbook_path: Path) -> int:
    rows_to_remove = extra_empty_rows_not_in_template(template_path, workbook_path)
    if not rows_to_remove:
        return 0
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmpf:
        temp_path = Path(tmpf.name)
    removed = 0
    try:
        with ZipFile(workbook_path) as src, ZipFile(temp_path, "w", ZIP_DEFLATED, compresslevel=ZIP_LEVEL) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                rm = rows_to_remove.get(info.filename)
                if rm:
                    root = ET.fromstring(data)
                    sheet_data = root.find("main:sheetData", NS)
                    if sheet_data is not None:
                        for row in list(sheet_data):
                            r = row.attrib.get("r")
                            if r and int(r) in rm and not row.findall("main:c", NS):
                                sheet_data.remove(row)
                                removed += 1
                        data = _ensure_ignorable_ns_declared(
                            ET.tostring(root, encoding="utf-8", xml_declaration=True)
                        )
                dst.writestr(info, data)
        shutil.move(str(temp_path), workbook_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return removed


def _value_children(cell: ET.Element) -> list[ET.Element]:
    return [
        ch
        for ch in list(cell)
        if ch.tag in {f"{{{SHEET_NS}}}f", f"{{{SHEET_NS}}}v", f"{{{SHEET_NS}}}is"}
    ]


def _patch_cell(template_cell: ET.Element, staging_cell: ET.Element | None) -> bool:
    """Patch a template cell with values from the staging cell.

    If staging_cell is None or has no meaningful content (no formula, value, or
    inline string), the template cell is left unchanged. This preserves template
    formulas for cells that openpyxl wrote as empty during the save cycle.
    """
    if staging_cell is None:
        return False
    staging_children = _value_children(staging_cell)
    staging_t = staging_cell.attrib.get("t", "")
    # If staging has no value content and just declares a type (like t='n' for
    # empty cells), skip patching to preserve the template's original content.
    if not staging_children and staging_t in ("", "n"):
        return False

    # Array formulas are structural template elements (status indicators, computed
    # fields). openpyxl preserves the formula text but drops the cached <v> value.
    # Excel requires array formula cells to either have a valid cached value or
    # be left exactly as the template defines them. Skip patching these cells
    # so the template's original structure (including any cached value) is preserved.
    staging_f = staging_cell.find(f"{{{SHEET_NS}}}f")
    if staging_f is not None and staging_f.attrib.get("t") == "array":
        return False

    before = ET.tostring(template_cell)
    for ch in _value_children(template_cell):
        template_cell.remove(ch)
    template_cell.attrib.pop("t", None)
    if "t" in staging_cell.attrib:
        template_cell.attrib["t"] = staging_cell.attrib["t"]
    for ch in staging_children:
        template_cell.append(copy.deepcopy(ch))
    return ET.tostring(template_cell) != before


def _remove_child(root: ET.Element, tag_name: str) -> None:
    tag = f"{{{SHEET_NS}}}{tag_name}"
    for ch in list(root):
        if ch.tag == tag:
            root.remove(ch)


def _drop_printer_settings_ref(root: ET.Element) -> None:
    """Remove printerSettings relationships that are not preserved in staging.

    openpyxl does not write the legacy printerSettings binary parts. If we keep
    a template pageSetup r:id while writing the staging package relationships,
    Excel sees a dangling relationship and refuses to open the workbook.
    """
    page_setup = root.find("main:pageSetup", NS)
    if page_setup is not None:
        page_setup.attrib.pop(f"{{{OFFICE_REL_NS}}}id", None)


def _insert_child(root: ET.Element, child: ET.Element, before_tags: tuple[str, ...]) -> None:
    before = {f"{{{SHEET_NS}}}{tag}" for tag in before_tags}
    for idx, existing in enumerate(list(root)):
        if existing.tag in before:
            root.insert(idx, child)
            return
    root.append(child)


def _strict_worksheet_xml(template_xml: bytes, staging_xml: bytes, shared_strings: list[str] | None = None) -> tuple[bytes, int]:
    template_root = ET.fromstring(template_xml)
    staging_root = ET.fromstring(staging_xml)

    staging_cells: dict[str, ET.Element] = {}
    for cell in staging_root.findall("main:sheetData/main:row/main:c", NS):
        ref = cell.attrib.get("r")
        if ref:
            staging_cells[ref] = cell

    patched = 0
    for tcell in template_root.findall("main:sheetData/main:row/main:c", NS):
        ref = tcell.attrib.get("r")
        if ref and _patch_cell(tcell, staging_cells.get(ref)):
            patched += 1

    # Convert any remaining t='s' (shared string) cells to inlineStr.
    # The output uses staging's package which has no sharedStrings.xml,
    # so shared string indices from the template would be dangling references.
    if shared_strings is not None:
        for tcell in template_root.findall("main:sheetData/main:row/main:c", NS):
            if tcell.attrib.get("t") == "s":
                v_elem = tcell.find(f"{{{SHEET_NS}}}v")
                if v_elem is not None and v_elem.text is not None:
                    try:
                        idx = int(v_elem.text)
                        text = shared_strings[idx] if idx < len(shared_strings) else ""
                    except (ValueError, IndexError):
                        text = ""
                    # Convert to inlineStr
                    tcell.remove(v_elem)
                    tcell.attrib["t"] = "inlineStr"
                    is_elem = ET.SubElement(tcell, f"{{{SHEET_NS}}}is")
                    t_elem = ET.SubElement(is_elem, f"{{{SHEET_NS}}}t")
                    t_elem.text = text
                    patched += 1
                else:
                    # No value - just remove the 't' attribute
                    tcell.attrib.pop("t", None)

    for tag in ("hyperlinks", "drawing", "legacyDrawing"):
        _remove_child(template_root, tag)

    _drop_printer_settings_ref(template_root)

    s_hyper = staging_root.find("main:hyperlinks", NS)
    if s_hyper is not None:
        _insert_child(
            template_root,
            copy.deepcopy(s_hyper),
            ("pageMargins", "pageSetup", "headerFooter", "drawing", "legacyDrawing"),
        )
    s_draw = staging_root.find("main:drawing", NS)
    if s_draw is not None:
        _insert_child(template_root, copy.deepcopy(s_draw), ("legacyDrawing",))
    s_leg = staging_root.find("main:legacyDrawing", NS)
    if s_leg is not None:
        _insert_child(template_root, copy.deepcopy(s_leg), ())

    return _ensure_ignorable_ns_declared(
        ET.tostring(template_root, encoding="utf-8", xml_declaration=True)
    ), patched


_BARE_AVLST_RE = _re.compile(rb"<avLst\s*/>")


def _fix_drawing_xml(data: bytes) -> bytes:
    """Fix openpyxl drawing bug: <avLst /> should be <a:avLst /> inside <a:prstGeom>.

    openpyxl writes the avLst element without the drawingml namespace prefix,
    placing it in the spreadsheetDrawing default namespace instead. Excel
    considers this invalid and flags the drawing as corrupt.
    """
    if b"<avLst" not in data:
        return data
    return _BARE_AVLST_RE.sub(b"<a:avLst />", data)


def _ensure_metadata_content_type(content_types_xml: bytes) -> bytes:
    root = ET.fromstring(content_types_xml)
    override_tag = "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
    for node in root.findall(override_tag):
        if node.attrib.get("PartName") == "/xl/metadata.xml":
            return content_types_xml
    ET.SubElement(
        root,
        override_tag,
        {
            "PartName": "/xl/metadata.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_metadata_relationship(workbook_rels_xml: bytes) -> bytes:
    root = ET.fromstring(workbook_rels_xml)
    rel_tag = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
    for node in root.findall(rel_tag):
        if node.attrib.get("Type") == rel_type and node.attrib.get("Target") == "metadata.xml":
            return workbook_rels_xml

    used_ids: set[int] = set()
    for node in root.findall(rel_tag):
        rid = node.attrib.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            used_ids.add(int(rid[3:]))
    next_id = 1
    while next_id in used_ids:
        next_id += 1

    ET.SubElement(
        root,
        rel_tag,
        {"Id": f"rId{next_id}", "Type": rel_type, "Target": "metadata.xml"},
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _load_shared_strings(zf: ZipFile) -> list[str] | None:
    """Load the shared strings table from an xlsx ZIP file."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return None
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall(f"{{{SHEET_NS}}}si"):
        # <si> can contain <t>text</t> or <r><t>text</t></r> (rich text)
        t_elem = si.find(f"{{{SHEET_NS}}}t")
        if t_elem is not None:
            strings.append(t_elem.text or "")
        else:
            # Rich text: concatenate all <r><t> elements
            parts = []
            for r in si.findall(f"{{{SHEET_NS}}}r"):
                rt = r.find(f"{{{SHEET_NS}}}t")
                if rt is not None:
                    parts.append(rt.text or "")
            strings.append("".join(parts))
    return strings


def _dedup_style_children(parent: ET.Element) -> tuple[dict[int, int], int]:
    children = list(parent)
    seen: dict[bytes, int] = {}
    remap: dict[int, int] = {}
    keep: list[ET.Element] = []

    for old_idx, child in enumerate(children):
        key = ET.tostring(child)
        if key in seen:
            remap[old_idx] = seen[key]
        else:
            new_idx = len(keep)
            seen[key] = new_idx
            remap[old_idx] = new_idx
            keep.append(child)

    if len(keep) == len(children):
        return remap, 0

    parent.clear()
    parent.attrib["count"] = str(len(keep))
    for child in keep:
        parent.append(child)
    return remap, len(children) - len(keep)


def _remap_xf_attr(root: ET.Element, attr: str, remap: dict[int, int]) -> None:
    if not remap:
        return
    for path in ("main:cellXfs/main:xf", "main:cellStyleXfs/main:xf"):
        for xf in root.findall(path, NS):
            value = xf.attrib.get(attr)
            if value is None:
                continue
            old_idx = int(value)
            xf.attrib[attr] = str(remap.get(old_idx, old_idx))


def _dedup_styles(styles_xml: bytes) -> tuple[bytes, dict[int, int], dict[int, int]]:
    """Deduplicate style tables and return cell/differential style remaps.

    WPS-authored templates often contain duplicate cellXfs entries. Excel
    considers this repairable and deduplicates on open, triggering a repair
    prompt. Fonts can be duplicated too; remapping fontId before deduplicating
    cellXfs catches style-equivalent entries that are not byte-identical at
    first pass. Differential styles are deduplicated with a separate dxfId remap.

    Returns (new_styles_xml, cell_style_remap, dxf_remap).
    """
    root = ET.fromstring(styles_xml)

    fonts = root.find(f"{{{SHEET_NS}}}fonts")
    if fonts is not None:
        font_remap, _ = _dedup_style_children(fonts)
        _remap_xf_attr(root, "fontId", font_remap)

    dxf_remap: dict[int, int] = {}
    dxfs = root.find(f"{{{SHEET_NS}}}dxfs")
    if dxfs is not None:
        dxf_remap, _ = _dedup_style_children(dxfs)

    cellXfs = root.find(f"{{{SHEET_NS}}}cellXfs")
    if cellXfs is None:
        return ET.tostring(root, encoding="utf-8", xml_declaration=True), {}, dxf_remap

    style_remap, _ = _dedup_style_children(cellXfs)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), style_remap, dxf_remap


def _remap_style_indices(worksheet_xml: bytes, style_remap: dict[int, int], dxf_remap: dict[int, int]) -> bytes:
    """Remap worksheet style indices after style table deduplication."""
    if not style_remap and not dxf_remap:
        return worksheet_xml

    root = ET.fromstring(worksheet_xml)
    if style_remap:
        for cell in root.findall(f"{{{SHEET_NS}}}sheetData/{{{SHEET_NS}}}row/{{{SHEET_NS}}}c"):
            s = cell.attrib.get("s")
            if s is not None:
                old_idx = int(s)
                cell.attrib["s"] = str(style_remap.get(old_idx, old_idx))

    if dxf_remap:
        for rule in root.findall(f".//{{{SHEET_NS}}}cfRule"):
            dxf_id = rule.attrib.get("dxfId")
            if dxf_id is not None:
                old_idx = int(dxf_id)
                rule.attrib["dxfId"] = str(dxf_remap.get(old_idx, old_idx))

    return _ensure_ignorable_ns_declared(
        ET.tostring(root, encoding="utf-8", xml_declaration=True)
    )


def build_strict_output(template_path: Path, staging_path: Path, output_path: Path) -> tuple[int, list[str]]:
    template_sheets = workbook_sheet_paths(template_path)
    staging_sheets = workbook_sheet_paths(staging_path)
    sheet_paths = {
        staging_sheets[name]: template_sheets[name]
        for name in staging_sheets
        if name in template_sheets
    }

    patched = 0
    name_fixes: list[str] = []
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(template_path) as template_zip, ZipFile(staging_path) as staging_zip, ZipFile(
        output_path, "w", ZIP_DEFLATED, compresslevel=ZIP_LEVEL
    ) as out_zip:
        # Load template's shared strings for converting t='s' cells
        template_shared_strings = _load_shared_strings(template_zip)
        dxf_remap: dict[int, int] = {}
        styles_data: bytes | None = None
        if "xl/styles.xml" in template_zip.namelist():
            styles_data = template_zip.read("xl/styles.xml")
        has_template_metadata = "xl/metadata.xml" in template_zip.namelist()

        for info in staging_zip.infolist():
            data = staging_zip.read(info.filename)
            template_sheet_path = sheet_paths.get(info.filename)
            if template_sheet_path:
                data, sheet_patched = _strict_worksheet_xml(
                    template_zip.read(template_sheet_path), data, template_shared_strings
                )
                data = _remap_style_indices(data, {}, dxf_remap)
                patched += sheet_patched
            elif info.filename == "xl/styles.xml":
                # Use template's styles.xml (worksheet XML uses template style indices)
                if styles_data is not None:
                    data = styles_data
            elif info.filename == "xl/theme/theme1.xml":
                if "xl/theme/theme1.xml" in template_zip.namelist():
                    data = template_zip.read("xl/theme/theme1.xml")
            elif info.filename == "xl/workbook.xml":
                data, name_fixes = _repair_workbook_defined_names(data)
            elif has_template_metadata and info.filename == "[Content_Types].xml":
                data = _ensure_metadata_content_type(data)
            elif has_template_metadata and info.filename == "xl/_rels/workbook.xml.rels":
                data = _ensure_metadata_relationship(data)
            elif info.filename.startswith("xl/drawings/") and info.filename.endswith(".xml"):
                data = _fix_drawing_xml(data)
            out_zip.writestr(info, data)
        if has_template_metadata and "xl/metadata.xml" not in staging_zip.namelist():
            out_zip.writestr("xl/metadata.xml", template_zip.read("xl/metadata.xml"))
    return patched, name_fixes


def _repair_workbook_defined_names(workbook_xml: bytes) -> tuple[bytes, list[str]]:
    """Fix global defined names that resolve to #REF! when a sheet-local name
    of the same identifier is valid. Operates directly on workbook.xml so it
    handles duplicate names that high-level libraries collapse.
    """
    fixes: list[str] = []
    root = ET.fromstring(workbook_xml)
    names_el = root.find("main:definedNames", NS)
    if names_el is None:
        return workbook_xml, fixes

    by_name: dict[str, list[ET.Element]] = {}
    for n in names_el.findall("main:definedName", NS):
        nm = n.attrib.get("name", "")
        by_name.setdefault(nm, []).append(n)

    changed = False
    for nm, group in by_name.items():
        global_node = next((e for e in group if "localSheetId" not in e.attrib), None)
        if global_node is None:
            continue
        text = (global_node.text or "").strip()
        if "#REF!" not in text:
            continue
        replacement = None
        for e in group:
            if "localSheetId" not in e.attrib:
                continue
            t = (e.text or "").strip()
            if t and "#REF!" not in t:
                replacement = t
                break
        if replacement:
            global_node.text = replacement
            fixes.append(f"defined name `{nm}`: #REF! -> {replacement}")
            changed = True

    if not changed:
        return workbook_xml, fixes
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), fixes
