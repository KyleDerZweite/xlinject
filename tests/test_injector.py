from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import Workbook  # type: ignore[import-untyped]

from xlinject.injector import replace_sentinel_in_column_range
from xlinject.workbook_map import map_sheet_name_to_part

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _create_test_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        raise RuntimeError("No active worksheet")
    ws.title = "Eingabemaske"

    ws["A1"] = "Timestamp"
    ws["B1"] = "Wert"

    ws["B2"] = -1
    ws["B3"] = 42
    ws["B4"] = -1
    ws["B5"] = -1
    ws["B6"] = -1

    ws["H2"] = (
        "=LET("
        "werte,B2:B6,"
        "valide,(ISTZAHL(werte))*(werte<>-1),"
        "res,FILTER(werte,valide,\"\"),"
        "WENN(ODER(res=\"\",ANZAHL(res)=0),\"\",MAX(res))"
        ")"
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _read_cell_v_text(xlsx_path: Path, cell_ref: str) -> str | None:
    with zipfile.ZipFile(xlsx_path, "r") as archive:
        sheet_part = map_sheet_name_to_part(archive, "Eingabemaske")
        root = ET.fromstring(archive.read(sheet_part))

    cell = root.find(f".//x:c[@r='{cell_ref}']", NS)
    if cell is None:
        return None
    value = cell.find("x:v", NS)
    return value.text if value is not None else None


def _read_formula_text(xlsx_path: Path, cell_ref: str) -> str | None:
    with zipfile.ZipFile(xlsx_path, "r") as archive:
        sheet_part = map_sheet_name_to_part(archive, "Eingabemaske")
        root = ET.fromstring(archive.read(sheet_part))

    cell = root.find(f".//x:c[@r='{cell_ref}']", NS)
    if cell is None:
        return None
    formula = cell.find("x:f", NS)
    return formula.text if formula is not None else None


def test_replace_only_sentinel_cells_and_preserve_formula(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _create_test_workbook(source)

    before_formula = _read_formula_text(source, "H2")

    report = replace_sentinel_in_column_range(
        source,
        output,
        sheet_name="Eingabemaske",
        range_ref="B2:B6",
        values=[1.1, 2.2, 3.3],
        sentinel=-1.0,
        guard_cells=["H2"],
    )

    assert report.replaced_count == 3
    assert report.consumed_values == 3
    assert report.untouched_sentinel_count == 1

    assert _read_cell_v_text(output, "B2") == "1.1"
    assert _read_cell_v_text(output, "B3") == "42"
    assert _read_cell_v_text(output, "B4") == "2.2"
    assert _read_cell_v_text(output, "B5") == "3.3"
    assert _read_cell_v_text(output, "B6") == "-1"

    after_formula = _read_formula_text(output, "H2")
    assert before_formula == after_formula


def test_guard_cell_change_raises(tmp_path: Path) -> None:
    source = tmp_path / "source_guard.xlsx"
    output = tmp_path / "output_guard.xlsx"
    _create_test_workbook(source)

    try:
        replace_sentinel_in_column_range(
            source,
            output,
            sheet_name="Eingabemaske",
            range_ref="B2:B2",
            values=[8.8],
            sentinel=-1.0,
            guard_cells=["B2"],
        )
    except RuntimeError as exc:
        assert "Guard cell signatures changed" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError when guard cell is modified")
