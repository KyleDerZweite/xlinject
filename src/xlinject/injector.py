from __future__ import annotations

from dataclasses import dataclass
import io
from pathlib import Path
from typing import Any, Iterable, Mapping, TypeAlias
import zipfile

from lxml import etree as ET

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


@dataclass(frozen=True)
class ValidationRule:
    sqref: str
    validation_type: str | None
    operator: str | None
    formula1: str | None
    allow_blank: bool
    error_message: str | None


X_MAIN = {"x": NS_MAIN}
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_SPACE = "http://www.w3.org/XML/1998/namespace"
IGNORABLE_NAMESPACE_FALLBACKS: dict[str, str] = {
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
}

XmlElement: TypeAlias = Any


def _parse_xml(xml_bytes: bytes) -> ET._Element:
    parser = ET.XMLParser(remove_blank_text=False, recover=False)
    return ET.fromstring(xml_bytes, parser=parser)


def _find_cell(root: XmlElement, cell_ref: str) -> XmlElement | None:
    return root.find(f".//x:c[@r='{cell_ref}']", X_MAIN)


def _extract_guard_signature(root: XmlElement, cell_ref: str) -> tuple[dict[str, str], dict[str, str] | None, str | None, str | None] | None:
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


def _iter_row_cells(sheet_data: XmlElement, row_number: int) -> tuple[XmlElement | None, list[XmlElement]]:
    row_tag = f"{{{NS_MAIN}}}row"
    cell_tag = f"{{{NS_MAIN}}}c"

    row_element: XmlElement | None = None
    cell_elements: list[XmlElement] = []

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


def _insert_row_sorted(sheet_data: XmlElement, row_number: int) -> XmlElement:
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


def _insert_cell_sorted(row: XmlElement, target_ref: str) -> XmlElement:
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


def _get_or_create_cell(root: XmlElement, column_name: str, row_number: int) -> XmlElement:
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


def _is_numeric_sentinel(cell: XmlElement, sentinel: float) -> bool:
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


def _set_numeric_value(cell: XmlElement, value: float) -> None:
    cell.attrib.pop("t", None)

    is_tag = f"{{{NS_MAIN}}}is"
    for node in list(cell.findall(is_tag)):
        cell.remove(node)

    value_tag = f"{{{NS_MAIN}}}v"
    value_node = cell.find(value_tag)
    if value_node is None:
        value_node = ET.SubElement(cell, value_tag)

    value_node.text = format(value, ".15g")


def _remove_child_nodes(cell: XmlElement, local_name: str) -> None:
    tag = f"{{{NS_MAIN}}}{local_name}"
    for node in list(cell.findall(tag)):
        cell.remove(node)


def _set_inline_string_value(cell: XmlElement, value: str) -> None:
    cell.attrib["t"] = "inlineStr"
    _remove_child_nodes(cell, "v")
    _remove_child_nodes(cell, "is")

    is_node = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
    text_node = ET.SubElement(is_node, f"{{{NS_MAIN}}}t")
    if value != value.strip() or "\n" in value or "  " in value:
        text_node.attrib[f"{{{XML_SPACE}}}space"] = "preserve"
    text_node.text = value


def _coerce_number(value: object) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if _is_nan(value) else float(value)
    return None


def _first_formula_text(data_validation: XmlElement) -> str | None:
    for node in data_validation.iter():
        if node.tag.endswith("}formula1") and node.text is not None:
            return str(node.text)
    return None


def _column_index_to_name(column_index: int) -> str:
    chars: list[str] = []
    current = column_index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        chars.append(chr(ord("A") + remainder))
    return "".join(reversed(chars))


def _expand_range_token(range_token: str) -> list[str]:
    token = range_token.strip().upper()
    if not token:
        return []
    if ":" not in token:
        return [token]

    start_ref, end_ref = token.split(":", maxsplit=1)
    start_column, start_row = split_cell_reference(start_ref)
    end_column, end_row = split_cell_reference(end_ref)
    start_column_index = column_name_to_index(start_column)
    end_column_index = column_name_to_index(end_column)

    refs: list[str] = []
    for row_number in range(start_row, end_row + 1):
        for column_index in range(start_column_index, end_column_index + 1):
            column_name = _column_index_to_name(column_index)
            refs.append(build_cell_reference(column_name, row_number))
    return refs


def _parse_sqref(sqref: str) -> list[str]:
    refs: list[str] = []
    for token in sqref.split():
        refs.extend(_expand_range_token(token))
    return refs


def _parse_list_options(formula: str | None) -> list[str]:
    if formula is None:
        return []
    text = formula.strip()
    if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
        text = text[1:-1]
    return [item.strip() for item in text.split(",")]


def _is_empty_value(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value == ""
    return False


def _validate_rule(cell_ref: str, value: object, rule: ValidationRule) -> None:
    if _is_empty_value(value):
        if rule.allow_blank:
            return
        raise ValueError(f"Cell {cell_ref} does not allow blank values")

    if rule.validation_type is None:
        return

    if rule.validation_type == "list":
        options = _parse_list_options(rule.formula1)
        if str(value) not in options:
            raise ValueError(f"Cell {cell_ref} expects one of {options!r}, got {value!r}")
        return

    if rule.validation_type == "textLength":
        expected_length = int(float(rule.formula1 or "0"))
        if len(str(value)) != expected_length:
            raise ValueError(
                f"Cell {cell_ref} expects text length {expected_length}, got {len(str(value))}"
            )
        return

    numeric_value = _coerce_number(value)
    if numeric_value is None:
        raise ValueError(f"Cell {cell_ref} expects a numeric value, got {value!r}")

    if rule.validation_type == "whole" and not numeric_value.is_integer():
        raise ValueError(f"Cell {cell_ref} expects a whole number, got {numeric_value!r}")

    threshold = float(rule.formula1 or "0")
    operator = rule.operator or "equal"
    comparisons = {
        "equal": numeric_value == threshold,
        "greaterThan": numeric_value > threshold,
        "greaterThanOrEqual": numeric_value >= threshold,
        "lessThan": numeric_value < threshold,
        "lessThanOrEqual": numeric_value <= threshold,
    }
    if operator not in comparisons:
        return
    if not comparisons[operator]:
        raise ValueError(
            "Cell "
            f"{cell_ref} violates {rule.validation_type} validation "
            f"{operator} {threshold}: {numeric_value!r}"
        )


def extract_validation_rules(
    input_path: Path,
    *,
    sheet_name: str,
) -> dict[str, list[ValidationRule]]:
    with zipfile.ZipFile(input_path, "r") as in_archive:
        sheet_part = map_sheet_name_to_part(in_archive, sheet_name)
        root = _parse_xml(in_archive.read(sheet_part))

    rules_by_cell: dict[str, list[ValidationRule]] = {}
    data_validations = root.find("x:dataValidations", X_MAIN)
    if data_validations is None:
        return rules_by_cell

    for data_validation in data_validations.findall("x:dataValidation", X_MAIN):
        rule = ValidationRule(
            sqref=data_validation.attrib.get("sqref", ""),
            validation_type=data_validation.attrib.get("type"),
            operator=data_validation.attrib.get("operator"),
            formula1=_first_formula_text(data_validation),
            allow_blank=data_validation.attrib.get("allowBlank") == "1",
            error_message=data_validation.attrib.get("error"),
        )
        for cell_ref in _parse_sqref(rule.sqref):
            rules_by_cell.setdefault(cell_ref, []).append(rule)

    return rules_by_cell


def validate_cell_values(
    input_path: Path,
    *,
    sheet_name: str,
    cell_values: Mapping[str, object],
) -> None:
    rules_by_cell = extract_validation_rules(input_path, sheet_name=sheet_name)
    for raw_ref, value in cell_values.items():
        cell_ref = str(raw_ref).strip().upper()
        if not cell_ref or _is_empty_value(value) or _is_nan(value):
            continue
        for rule in rules_by_cell.get(cell_ref, []):
            _validate_rule(cell_ref, value, rule)


def _build_cell_cache(
    sheet_data: XmlElement,
) -> tuple[dict[int, XmlElement], dict[str, XmlElement]]:
    row_tag = f"{{{NS_MAIN}}}row"
    cell_tag = f"{{{NS_MAIN}}}c"
    row_cache: dict[int, XmlElement] = {}
    cell_cache: dict[str, XmlElement] = {}

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

    return row_cache, cell_cache


def _get_or_create_cell_fast(
    sheet_data: XmlElement,
    row_cache: dict[int, XmlElement],
    cell_cache: dict[str, XmlElement],
    cell_ref: str,
) -> XmlElement:
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
    root: XmlElement,
    original_sheet_xml: bytes,
) -> bytes:
    serialized = ET.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )

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


def _count_remaining_sentinel(root: XmlElement, column_name: str, start_row: int, end_row: int, sentinel: float) -> int:
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

        root = _parse_xml(original_sheet_xml)

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
        post_root = _parse_xml(out_archive.read(map_sheet_name_to_part(out_archive, sheet_name)))
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
        root = _parse_xml(original_sheet_xml)
        sheet_data = root.find("x:sheetData", X_MAIN)
        if sheet_data is None:
            raise RuntimeError("Worksheet XML is missing <sheetData>")

        row_tag = f"{{{NS_MAIN}}}row"
        cell_tag = f"{{{NS_MAIN}}}c"

        row_cache: dict[int, XmlElement] = {}
        cell_cache: dict[str, XmlElement] = {}

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

        def _get_or_create_cell_fast(cell_ref: str) -> XmlElement:
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


def write_cells(
    input_path: Path,
    output_path: Path,
    *,
    sheet_name: str,
    cell_values: Mapping[str, object],
    guard_cells: Iterable[str] = (),
    allow_formula_overwrite: bool = False,
    validate_sheet_rules: bool = False,
) -> WriteReport:
    normalized_items: list[tuple[str, object]] = []
    skipped_nan_count = 0

    for raw_ref, raw_value in cell_values.items():
        cell_ref = str(raw_ref).strip().upper()
        if not cell_ref or _is_empty_value(raw_value):
            continue
        if _is_nan(raw_value):
            skipped_nan_count += 1
            continue
        normalized_items.append((cell_ref, raw_value))

    if validate_sheet_rules:
        validate_cell_values(
            input_path,
            sheet_name=sheet_name,
            cell_values={ref: value for ref, value in normalized_items},
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    temp_output = output_path
    writing_in_place = input_path.resolve() == output_path.resolve()
    if writing_in_place:
        temp_output = output_path.with_suffix(f"{output_path.suffix}.tmp")

    with zipfile.ZipFile(input_path, "r") as in_archive:
        sheet_part = map_sheet_name_to_part(in_archive, sheet_name)
        original_sheet_xml = in_archive.read(sheet_part)
        root = _parse_xml(original_sheet_xml)
        sheet_data = root.find("x:sheetData", X_MAIN)
        if sheet_data is None:
            raise RuntimeError("Worksheet XML is missing <sheetData>")

        row_cache, cell_cache = _build_cell_cache(sheet_data)

        guard_refs = [ref.strip().upper() for ref in guard_cells if ref.strip()]
        guard_before = {
            ref: _extract_guard_signature(root, ref)
            for ref in guard_refs
        }

        written_count = 0
        for cell_ref, raw_value in normalized_items:
            cell = _get_or_create_cell_fast(sheet_data, row_cache, cell_cache, cell_ref)

            if not allow_formula_overwrite and cell.find("x:f", X_MAIN) is not None:
                raise RuntimeError(
                    f"Refusing to overwrite formula cell: {cell_ref}. "
                    "Set allow_formula_overwrite=True to override."
                )

            numeric_value = _coerce_number(raw_value)
            if numeric_value is not None:
                _set_numeric_value(cell, numeric_value)
            else:
                _set_inline_string_value(cell, str(raw_value))
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
