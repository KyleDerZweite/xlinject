from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import zipfile
import xml.etree.ElementTree as ET

from xlinject.cellrefs import (
    build_cell_reference,
    column_name_to_index,
    parse_single_column_range,
    split_cell_reference,
)
from xlinject.workbook_map import NS_MAIN, map_sheet_name_to_part


@dataclass(frozen=True)
class ReplaceReport:
    replaced_count: int
    consumed_values: int
    untouched_sentinel_count: int
    output_file: Path


X_MAIN = {"x": NS_MAIN}


def _find_cell(root: ET.Element, cell_ref: str) -> ET.Element | None:
    return root.find(f".//x:c[@r='{cell_ref}']", X_MAIN)


def _extract_guard_signature(root: ET.Element, cell_ref: str) -> tuple[dict[str, str], dict[str, str] | None, str | None, str | None] | None:
    cell = _find_cell(root, cell_ref)
    if cell is None:
        return None
    formula = cell.find("x:f", X_MAIN)
    value = cell.find("x:v", X_MAIN)

    return (
        dict(cell.attrib),
        dict(formula.attrib) if formula is not None else None,
        formula.text if formula is not None else None,
        value.text if value is not None else None,
    )


def _iter_row_cells(sheet_data: ET.Element, row_number: int) -> tuple[ET.Element | None, list[ET.Element]]:
    row_tag = f"{{{NS_MAIN}}}row"
    cell_tag = f"{{{NS_MAIN}}}c"

    row_element: ET.Element | None = None
    cell_elements: list[ET.Element] = []

    for row in sheet_data.findall(row_tag):
        row_r = row.attrib.get("r")
        if row_r is None:
            continue
        try:
            current_row = int(row_r)
        except ValueError:
            continue
        if current_row == row_number:
            row_element = row
            cell_elements = list(row.findall(cell_tag))
            break

    return row_element, cell_elements


def _insert_row_sorted(sheet_data: ET.Element, row_number: int) -> ET.Element:
    row_tag = f"{{{NS_MAIN}}}row"
    new_row = ET.Element(row_tag, {"r": str(row_number)})

    inserted = False
    existing_rows = list(sheet_data.findall(row_tag))
    for idx, row in enumerate(existing_rows):
        row_r = row.attrib.get("r")
        if row_r is None:
            continue
        try:
            current_row = int(row_r)
        except ValueError:
            continue
        if current_row > row_number:
            sheet_data.insert(idx, new_row)
            inserted = True
            break

    if not inserted:
        sheet_data.append(new_row)

    return new_row


def _insert_cell_sorted(row: ET.Element, target_ref: str) -> ET.Element:
    cell_tag = f"{{{NS_MAIN}}}c"
    target_column, _ = split_cell_reference(target_ref)
    target_index = column_name_to_index(target_column)

    new_cell = ET.Element(cell_tag, {"r": target_ref})

    inserted = False
    existing_cells = list(row.findall(cell_tag))
    for idx, cell in enumerate(existing_cells):
        cell_ref = cell.attrib.get("r")
        if cell_ref is None:
            continue
        current_column, _ = split_cell_reference(cell_ref)
        current_index = column_name_to_index(current_column)
        if current_index > target_index:
            row.insert(idx, new_cell)
            inserted = True
            break

    if not inserted:
        row.append(new_cell)

    return new_cell


def _get_or_create_cell(root: ET.Element, column_name: str, row_number: int) -> ET.Element:
    sheet_data = root.find("x:sheetData", X_MAIN)
    if sheet_data is None:
        raise RuntimeError("Worksheet XML is missing <sheetData>")

    cell_ref = build_cell_reference(column_name, row_number)

    row_element, _ = _iter_row_cells(sheet_data, row_number)
    if row_element is None:
        row_element = _insert_row_sorted(sheet_data, row_number)

    existing_cell = None
    for cell in row_element.findall(f"{{{NS_MAIN}}}c"):
        if cell.attrib.get("r") == cell_ref:
            existing_cell = cell
            break

    if existing_cell is not None:
        return existing_cell

    return _insert_cell_sorted(row_element, cell_ref)


def _is_numeric_sentinel(cell: ET.Element, sentinel: float) -> bool:
    if cell.find("x:f", X_MAIN) is not None:
        return False

    cell_type = cell.attrib.get("t")
    if cell_type not in (None, "n"):
        return False

    value_node = cell.find("x:v", X_MAIN)
    if value_node is None or value_node.text is None:
        return False

    try:
        numeric_value = float(value_node.text)
    except ValueError:
        return False

    return abs(numeric_value - sentinel) < 1e-12


def _set_numeric_value(cell: ET.Element, value: float) -> None:
    cell.attrib.pop("t", None)

    is_tag = f"{{{NS_MAIN}}}is"
    for node in list(cell.findall(is_tag)):
        cell.remove(node)

    value_tag = f"{{{NS_MAIN}}}v"
    value_node = cell.find(value_tag)
    if value_node is None:
        value_node = ET.SubElement(cell, value_tag)

    value_node.text = format(value, ".15g")


def _count_remaining_sentinel(root: ET.Element, column_name: str, start_row: int, end_row: int, sentinel: float) -> int:
    remaining = 0
    for row in range(start_row, end_row + 1):
        ref = build_cell_reference(column_name, row)
        cell = _find_cell(root, ref)
        if cell is None:
            continue
        if _is_numeric_sentinel(cell, sentinel):
            remaining += 1
    return remaining


def replace_sentinel_in_column_range(
    input_path: Path,
    output_path: Path,
    *,
    sheet_name: str,
    range_ref: str,
    values: Iterable[float],
    sentinel: float = -1.0,
    guard_cells: Iterable[str] = (),
) -> ReplaceReport:
    column_name, start_row, end_row = parse_single_column_range(range_ref)
    pending_values = list(values)
    value_index = 0

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(input_path, "r") as in_archive:
        sheet_part = map_sheet_name_to_part(in_archive, sheet_name)
        original_sheet_xml = in_archive.read(sheet_part)

        root = ET.fromstring(original_sheet_xml)

        guard_refs = [ref.strip().upper() for ref in guard_cells if ref.strip()]
        guard_before = {
            ref: _extract_guard_signature(root, ref)
            for ref in guard_refs
        }

        replaced_count = 0
        for row in range(start_row, end_row + 1):
            if value_index >= len(pending_values):
                break

            cell = _get_or_create_cell(root, column_name, row)
            if not _is_numeric_sentinel(cell, sentinel):
                continue

            _set_numeric_value(cell, float(pending_values[value_index]))
            value_index += 1
            replaced_count += 1

        guard_after = {
            ref: _extract_guard_signature(root, ref)
            for ref in guard_refs
        }

        changed_guard_refs = [
            ref for ref in guard_refs
            if guard_before.get(ref) != guard_after.get(ref)
        ]
        if changed_guard_refs:
            changed = ", ".join(changed_guard_refs)
            raise RuntimeError(
                "Guard cell signatures changed after write operation: "
                f"{changed}"
            )

        updated_sheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(output_path, "w") as out_archive:
            for item in in_archive.infolist():
                data = updated_sheet_xml if item.filename == sheet_part else in_archive.read(item.filename)
                out_archive.writestr(item, data)

    remaining_sentinel = 0
    with zipfile.ZipFile(output_path, "r") as out_archive:
        post_root = ET.fromstring(out_archive.read(map_sheet_name_to_part(out_archive, sheet_name)))
        remaining_sentinel = _count_remaining_sentinel(
            post_root,
            column_name,
            start_row,
            end_row,
            sentinel,
        )

    return ReplaceReport(
        replaced_count=replaced_count,
        consumed_values=value_index,
        untouched_sentinel_count=remaining_sentinel,
        output_file=output_path,
    )
