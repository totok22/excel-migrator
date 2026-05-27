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

ET.register_namespace("", SHEET_NS)
ET.register_namespace("r", OFFICE_REL_NS)


# Lower compresslevel makes a noticeable difference for large image-heavy xlsx.
ZIP_LEVEL = 1


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
                        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
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
    before = ET.tostring(template_cell)
    for ch in _value_children(template_cell):
        template_cell.remove(ch)
    template_cell.attrib.pop("t", None)
    if staging_cell is not None:
        if "t" in staging_cell.attrib:
            template_cell.attrib["t"] = staging_cell.attrib["t"]
        for ch in _value_children(staging_cell):
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

    return ET.tostring(template_root, encoding="utf-8", xml_declaration=True), patched


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
