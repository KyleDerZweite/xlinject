from __future__ import annotations

from dataclasses import dataclass
import io
from pathlib import Path
from typing import Iterable, Mapping, cast
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


@dataclass(frozen=True)
class WriteReport:
    written_count: int
    skipped_nan_count: int
    output_file: Path


X_MAIN = {"x": NS_MAIN}
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
IGNORABLE_NAMESPACE_FALLBACKS: dict[str, str] = {
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
}


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


def _is_nan(value: object) -> bool:
    try:
        return bool(value != value)
    except Exception:
        return False


def _write_archive_with_sheet_update(
    in_archive: zipfile.ZipFile,
    output_path: Path,
    sheet_part: str,
    updated_sheet_xml: bytes,
) -> None:
    with zipfile.ZipFile(output_path, "w") as out_archive:
        for item in in_archive.infolist():
            data = (
                updated_sheet_xml
                if item.filename == sheet_part
                else in_archive.read(item.filename)
            )
            out_archive.writestr(item, data)


def _collect_namespace_prefixes(xml_bytes: bytes) -> dict[str, str]:
    prefix_map: dict[str, str] = {}
    for event, data in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        del event
        if not isinstance(data, tuple) or len(data) != 2:
            continue

        raw_prefix, raw_uri = data
        if not isinstance(raw_uri, str):
            continue

        prefix = raw_prefix if isinstance(raw_prefix, str) else ""
        if prefix not in prefix_map:
            prefix_map[prefix] = raw_uri
    return prefix_map


def _serialize_with_ignorable_namespace_preservation(
    root: ET.Element,
    original_sheet_xml: bytes,
) -> bytes:
    serialized = cast(bytes, ET.tostring(root, encoding="utf-8", xml_declaration=True))

    ignorable_text = None
    ignorable_key = f"{{{NS_MC}}}Ignorable"
    if ignorable_key in root.attrib:
        ignorable_text = root.attrib.get(ignorable_key)

    if not ignorable_text:
        return serialized

    original_prefixes = _collect_namespace_prefixes(original_sheet_xml)
    required_prefixes = [p for p in str(ignorable_text).split() if p]

    text = serialized.decode("utf-8")
    root_start = text.find("<", text.find("?>") + 2 if "?>" in text else 0)
    if root_start < 0:
        return serialized

    root_end = text.find(">", root_start)
    if root_end < 0:
        return serialized

    start_tag = text[root_start:root_end]
    additions: list[str] = []
    for prefix in required_prefixes:
        if prefix == "":
            continue
        if f"xmlns:{prefix}=" in start_tag:
            continue
        uri = original_prefixes.get(prefix) or IGNORABLE_NAMESPACE_FALLBACKS.get(prefix)
        if uri:
            additions.append(f' xmlns:{prefix}="{uri}"')

    if not additions:
        return serialized

    patched_start_tag = start_tag + "".join(additions)
    patched_text = text[:root_start] + patched_start_tag + text[root_end:]
    return patched_text.encode("utf-8")


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

        updated_sheet_xml = _serialize_with_ignorable_namespace_preservation(
            root,
            original_sheet_xml,
        )

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


def write_numeric_cells(
    input_path: Path,
    output_path: Path,
    *,
    sheet_name: str,
    cell_values: Mapping[str, float],
    guard_cells: Iterable[str] = (),
    allow_formula_overwrite: bool = False,
) -> WriteReport:
    normalized_items = [(str(ref).strip().upper(), value) for ref, value in cell_values.items()]
    skipped_nan_count = 0

    output_path.parent.mkdir(parents=True, exist_ok=True)

    temp_output = output_path
    writing_in_place = input_path.resolve() == output_path.resolve()
    if writing_in_place:
        temp_output = output_path.with_suffix(f"{output_path.suffix}.tmp")

    with zipfile.ZipFile(input_path, "r") as in_archive:
        sheet_part = map_sheet_name_to_part(in_archive, sheet_name)
        original_sheet_xml = in_archive.read(sheet_part)
        root = ET.fromstring(original_sheet_xml)
        sheet_data = root.find("x:sheetData", X_MAIN)
        if sheet_data is None:
            raise RuntimeError("Worksheet XML is missing <sheetData>")

        row_tag = f"{{{NS_MAIN}}}row"
        cell_tag = f"{{{NS_MAIN}}}c"

        row_cache: dict[int, ET.Element] = {}
        cell_cache: dict[str, ET.Element] = {}

        for row in sheet_data.findall(row_tag):
            row_r = row.attrib.get("r")
            if row_r is None:
                continue
            try:
                row_number = int(row_r)
            except ValueError:
                continue
            row_cache[row_number] = row

            for cell in row.findall(cell_tag):
                ref = cell.attrib.get("r")
                if ref:
                    cell_cache[ref.strip().upper()] = cell

        def _get_or_create_cell_fast(cell_ref: str) -> ET.Element:
            existing = cell_cache.get(cell_ref)
            if existing is not None:
                return existing

            column_name, row_number = split_cell_reference(cell_ref)
            row_element = row_cache.get(row_number)
            if row_element is None:
                row_element = _insert_row_sorted(sheet_data, row_number)
                row_cache[row_number] = row_element

            new_cell = _insert_cell_sorted(row_element, build_cell_reference(column_name, row_number))
            cell_cache[cell_ref] = new_cell
            return new_cell

        guard_refs = [ref.strip().upper() for ref in guard_cells if ref.strip()]
        guard_before = {
            ref: _extract_guard_signature(root, ref)
            for ref in guard_refs
        }

        written_count = 0
        for cell_ref, raw_value in normalized_items:
            if _is_nan(raw_value):
                skipped_nan_count += 1
                continue

            cell = _get_or_create_cell_fast(cell_ref)

            if not allow_formula_overwrite and cell.find("x:f", X_MAIN) is not None:
                raise RuntimeError(
                    f"Refusing to overwrite formula cell: {cell_ref}. "
                    "Set allow_formula_overwrite=True to override."
                )

            _set_numeric_value(cell, float(raw_value))
            written_count += 1

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

        updated_sheet_xml = _serialize_with_ignorable_namespace_preservation(
            root,
            original_sheet_xml,
        )
        _write_archive_with_sheet_update(in_archive, temp_output, sheet_part, updated_sheet_xml)

    if writing_in_place:
        temp_output.replace(output_path)

    return WriteReport(
        written_count=written_count,
        skipped_nan_count=skipped_nan_count,
        output_file=output_path,
    )
