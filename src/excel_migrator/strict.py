"""Strict template write-back.

Use the new template's worksheet XML as the structural base, only patching cell
values from the staging workbook. Keeps formatting/merged cells/validations
intact and lets us trim empty rows that openpyxl introduced.
"""

from __future__ import annotations

import copy
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


def _strict_worksheet_xml(template_xml: bytes, staging_xml: bytes) -> tuple[bytes, int]:
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
        for info in staging_zip.infolist():
            data = staging_zip.read(info.filename)
            template_sheet_path = sheet_paths.get(info.filename)
            if template_sheet_path:
                data, sheet_patched = _strict_worksheet_xml(template_zip.read(template_sheet_path), data)
                patched += sheet_patched
            elif info.filename == "xl/workbook.xml":
                data, name_fixes = _repair_workbook_defined_names(data)
            elif info.filename.startswith("xl/drawings/") and info.filename.endswith(".xml"):
                data = _fix_drawing_xml(data)
            out_zip.writestr(info, data)
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
