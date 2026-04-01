from __future__ import annotations

from datetime import datetime
from math import isnan
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping, TypeAlias
import io
import zipfile

from lxml import etree as ET

from xlinject.cellrefs import build_cell_reference
from xlinject.injector import WriteReport, write_cells, write_numeric_cells
from xlinject.workbook_map import NS_MAIN, map_sheet_name_to_part

NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"x": NS_MAIN}

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


def to_excel_serial(value: object) -> float | None:
    """Convert datetime-like values to Excel serial date number (1900 date system)."""
    if value is None:
        return None

    dt: datetime
    if isinstance(value, datetime):
        dt = value
    elif hasattr(value, "to_pydatetime"):
        try:
            dt = value.to_pydatetime()  # pandas.Timestamp compatibility without hard dependency
        except Exception:
            return None
    else:
        return None

    if dt.tzinfo is not None:
        try:
            dt = dt.replace(tzinfo=None)
        except Exception:
            return None

    epoch = datetime(1899, 12, 30)
    return (dt - epoch).total_seconds() / 86400.0


def normalize_numeric_value(value: object) -> float | None:
    """Normalize value into float; skip empty/None/NaN values."""
    if value is None:
        return None

    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return None
        value = text.replace(",", ".")

    if isinstance(value, (int, float)):
        numeric = float(value)
    else:
        try:
            numeric = float(str(value))
        except Exception:
            return None

    if isnan(numeric):
        return None

    return numeric


def build_column_cell_map(
    column_index: int,
    values: Iterable[object],
    start_row: int,
    *,
    value_transform: Callable[[object], float | None] | None = None,
) -> dict[str, float]:
    """Build an A1->value mapping for one column from raw iterable values.

    By default values are normalized via `normalize_numeric_value`.
    A custom `value_transform` can be provided (for example `to_excel_serial`).
    """
    if column_index <= 0:
        raise ValueError("column_index must be positive")
    if start_row <= 0:
        raise ValueError("start_row must be positive")

    # Local conversion to avoid exposing index conversion internals here
    chars: list[str] = []
    current = column_index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        chars.append(chr(ord("A") + remainder))
    col_name = "".join(reversed(chars))

    transform = value_transform or normalize_numeric_value

    mapped: dict[str, float] = {}
    for i, raw_val in enumerate(values):
        converted = transform(raw_val)
        if converted is None:
            continue
        mapped[build_cell_reference(col_name, start_row + i)] = float(converted)

    return mapped


def merge_cell_maps(*maps: Mapping[str, float]) -> dict[str, float]:
    """Merge multiple A1->value mappings, where later maps overwrite earlier maps."""
    merged: dict[str, float] = {}
    for cell_map in maps:
        merged.update({str(k).strip().upper(): float(v) for k, v in cell_map.items()})
    return merged


def _normalize_numeric_cell_values(cell_values: Mapping[str, object]) -> dict[str, float]:
    normalized: dict[str, float] = {}
    for raw_ref, raw_val in cell_values.items():
        ref = str(raw_ref).strip().upper()
        if not ref:
            continue

        numeric = normalize_numeric_value(raw_val)
        if numeric is None:
            continue

        normalized[ref] = numeric
    return normalized


def remove_calc_chain(xlsx_path: str | Path) -> None:
    """Remove xl/calcChain.xml to trigger safe recalculation on next Excel open."""
    path = Path(xlsx_path)
    temp_path = path.with_suffix(f"{path.suffix}.tmp")

    with zipfile.ZipFile(path, "r") as zip_in:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for item in zip_in.infolist():
                if item.filename != "xl/calcChain.xml":
                    zip_out.writestr(item, zip_in.read(item.filename))

    temp_path.replace(path)


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


def _serialize_with_ignorable_namespace_preservation(root: XmlElement, original_xml: bytes) -> bytes:
    serialized = ET.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )

    ignorable_text = root.attrib.get(f"{{{NS_MC}}}Ignorable")
    if not ignorable_text:
        return serialized

    original_prefixes = _collect_namespace_prefixes(original_xml)
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
        if f"xmlns:{prefix}=" in start_tag:
            continue
        uri = original_prefixes.get(prefix) or IGNORABLE_NAMESPACE_FALLBACKS.get(prefix)
        if uri:
            additions.append(f' xmlns:{prefix}="{uri}"')

    if not additions:
        return serialized

    patched_start_tag = start_tag + "".join(additions)
    patched = text[:root_start] + patched_start_tag + text[root_end:]
    return patched.encode("utf-8")


def _remove_formula_cached_values(sheet_xml_bytes: bytes) -> bytes:
    root = _parse_xml(sheet_xml_bytes)
    formula_cells = root.findall(".//x:c[x:f]", NS)
    for cell in formula_cells:
        value_node = cell.find("x:v", NS)
        if value_node is not None:
            cell.remove(value_node)
    return _serialize_with_ignorable_namespace_preservation(root, sheet_xml_bytes)


def apply_recalc_policy(
    xlsx_path: str | Path,
    *,
    remove_calc_chain_file: bool = True,
    set_full_calc_on_load: bool = True,
    clear_formula_cached_values: bool = False,
    clear_formula_cache_sheets: Iterable[str] = (),
) -> None:
    """Apply workbook recalculation policy after cell injection.

    - Optionally removes `xl/calcChain.xml`
    - Optionally sets workbook `calcPr` flags for full recalculation on load
    - Optionally clears cached `<v>` values for formula cells on selected sheets
    """
    path = Path(xlsx_path)
    temp_path = path.with_suffix(f"{path.suffix}.tmp")

    with zipfile.ZipFile(path, "r") as zip_in:
        workbook_xml = zip_in.read("xl/workbook.xml")
        workbook_root = _parse_xml(workbook_xml)

        if set_full_calc_on_load:
            calc_pr = workbook_root.find("x:calcPr", NS)
            if calc_pr is None:
                calc_pr = ET.SubElement(workbook_root, f"{{{NS_MAIN}}}calcPr")
            calc_pr.set("fullCalcOnLoad", "1")
            calc_pr.set("forceFullCalc", "1")
            calc_pr.set("calcOnSave", "1")
            calc_pr.set("calcCompleted", "0")

        updated_workbook_xml = _serialize_with_ignorable_namespace_preservation(
            workbook_root,
            workbook_xml,
        )

        sheet_parts_to_clear: set[str] = set()
        if clear_formula_cached_values:
            for sheet_name in clear_formula_cache_sheets:
                try:
                    sheet_part = map_sheet_name_to_part(zip_in, str(sheet_name))
                    sheet_parts_to_clear.add(sheet_part)
                except Exception:
                    continue

        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for item in zip_in.infolist():
                if remove_calc_chain_file and item.filename == "xl/calcChain.xml":
                    continue

                if item.filename == "xl/workbook.xml":
                    zip_out.writestr(item, updated_workbook_xml)
                    continue

                if item.filename in sheet_parts_to_clear:
                    original_sheet_xml = zip_in.read(item.filename)
                    updated_sheet_xml = _remove_formula_cached_values(original_sheet_xml)
                    zip_out.writestr(item, updated_sheet_xml)
                    continue

                zip_out.writestr(item, zip_in.read(item.filename))

    temp_path.replace(path)


def inject_cells(
    input_path: str | Path,
    output_path: str | Path,
    *,
    sheet_name: str,
    cell_values: Mapping[str, object],
    guard_cells: list[str] | tuple[str, ...] = (),
    allow_formula_overwrite: bool = False,
    remove_calc_chain_after_write: bool = True,
    set_full_calc_on_load: bool = True,
    clear_formula_cached_values: bool = False,
) -> WriteReport:
    """High-level dict-based injection API for easy project integration.

    Users can pass a plain A1->value mapping and paths as strings or Path objects.
    Values may be int/float/str; empty values and NaN are skipped.
    """
    normalized = _normalize_numeric_cell_values(cell_values)

    report = write_numeric_cells(
        Path(input_path),
        Path(output_path),
        sheet_name=sheet_name,
        cell_values=normalized,
        guard_cells=guard_cells,
        allow_formula_overwrite=allow_formula_overwrite,
    )

    if remove_calc_chain_after_write or set_full_calc_on_load or clear_formula_cached_values:
        apply_recalc_policy(
            output_path,
            remove_calc_chain_file=remove_calc_chain_after_write,
            set_full_calc_on_load=set_full_calc_on_load,
            clear_formula_cached_values=clear_formula_cached_values,
            clear_formula_cache_sheets=(sheet_name,),
        )

    return report


def inject_cells_mixed(
    input_path: str | Path,
    output_path: str | Path,
    *,
    sheet_name: str,
    cell_values: Mapping[str, object],
    guard_cells: list[str] | tuple[str, ...] = (),
    allow_formula_overwrite: bool = False,
    validate_sheet_rules: bool = True,
    remove_calc_chain_after_write: bool = True,
    set_full_calc_on_load: bool = True,
    clear_formula_cached_values: bool = False,
) -> WriteReport:
    """High-level mixed-value injection API for numeric and string cells.

    This API preserves the existing XML-first write strategy while supporting
    text values via `inlineStr` cells. Empty values and NaN are skipped.
    Optionally validates candidate values against direct worksheet validation rules
    before any mutation is written.
    """
    report = write_cells(
        Path(input_path),
        Path(output_path),
        sheet_name=sheet_name,
        cell_values=cell_values,
        guard_cells=guard_cells,
        allow_formula_overwrite=allow_formula_overwrite,
        validate_sheet_rules=validate_sheet_rules,
    )

    if remove_calc_chain_after_write or set_full_calc_on_load or clear_formula_cached_values:
        apply_recalc_policy(
            output_path,
            remove_calc_chain_file=remove_calc_chain_after_write,
            set_full_calc_on_load=set_full_calc_on_load,
            clear_formula_cached_values=clear_formula_cached_values,
            clear_formula_cache_sheets=(sheet_name,),
        )

    return report
