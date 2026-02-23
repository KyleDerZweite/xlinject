from __future__ import annotations

import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _normalize_target(target: str) -> str:
    stripped = target.strip()
    if stripped.startswith("/"):
        stripped = stripped[1:]
    if stripped.startswith("xl/"):
        return stripped
    return f"xl/{stripped}"


def map_sheet_name_to_part(archive: ZipFile, sheet_name: str) -> str:
    workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
    workbook_rels_xml = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    relationship_targets: dict[str, str] = {}
    rel_tag = f"{{{NS_REL}}}Relationship"
    for relationship in workbook_rels_xml.findall(rel_tag):
        rel_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if rel_id and target:
            relationship_targets[rel_id] = _normalize_target(target)

    sheet_tag = f"{{{NS_MAIN}}}sheet"
    sheets_tag = f"{{{NS_MAIN}}}sheets"

    sheets_parent = workbook_xml.find(sheets_tag)
    if sheets_parent is None:
        raise RuntimeError("Workbook has no <sheets> node")

    rid_key = f"{{{NS_DOC_REL}}}id"
    for sheet in sheets_parent.findall(sheet_tag):
        name = sheet.attrib.get("name")
        if name != sheet_name:
            continue
        rel_id = sheet.attrib.get(rid_key)
        if rel_id is None:
            break
        if rel_id not in relationship_targets:
            break
        return relationship_targets[rel_id]

    raise ValueError(f"Sheet not found or unresolved in workbook relationships: {sheet_name}")
