from __future__ import annotations

from pathlib import Path
from datetime import datetime, timezone
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import Workbook  # type: ignore[import-untyped]

from xlinject.cli_write_cells import _load_cells_from_csv, _load_cells_from_json, _load_cells_from_json_text
from xlinject.highlevel import apply_recalc_policy, build_column_cell_map, inject_cells, merge_cell_maps, normalize_numeric_value, to_excel_serial
from xlinject.injector import replace_sentinel_in_column_range, write_numeric_cells
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


def test_write_numeric_cells_preserve_formula_and_skip_nan(tmp_path: Path) -> None:
    source = tmp_path / "source_write.xlsx"
    output = tmp_path / "output_write.xlsx"
    _create_test_workbook(source)

    before_formula = _read_formula_text(source, "H2")

    report = write_numeric_cells(
        source,
        output,
        sheet_name="Eingabemaske",
        cell_values={"B2": 5.5, "B4": 6.6, "B7": float("nan")},
        guard_cells=["H2"],
    )

    assert report.written_count == 2
    assert report.skipped_nan_count == 1

    assert _read_cell_v_text(output, "B2") == "5.5"
    assert _read_cell_v_text(output, "B4") == "6.6"

    after_formula = _read_formula_text(output, "H2")
    assert before_formula == after_formula


def test_write_numeric_cells_in_place(tmp_path: Path) -> None:
    source = tmp_path / "source_in_place.xlsx"
    _create_test_workbook(source)

    report = write_numeric_cells(
        source,
        source,
        sheet_name="Eingabemaske",
        cell_values={"B5": 9.9},
    )

    assert report.written_count == 1
    assert _read_cell_v_text(source, "B5") == "9.9"


def test_write_numeric_cells_reject_formula_overwrite(tmp_path: Path) -> None:
    source = tmp_path / "source_formula_guard.xlsx"
    output = tmp_path / "output_formula_guard.xlsx"
    _create_test_workbook(source)

    try:
        write_numeric_cells(
            source,
            output,
            sheet_name="Eingabemaske",
            cell_values={"H2": 12.3},
        )
    except RuntimeError as exc:
        assert "Refusing to overwrite formula cell" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError when attempting to overwrite formula cell")


def test_write_numeric_cells_preserves_ignorable_namespaces(tmp_path: Path) -> None:
    source = tmp_path / "source_ns.xlsx"
    output = tmp_path / "output_ns.xlsx"
    _create_test_workbook(source)

    with zipfile.ZipFile(source, "r") as zin:
        sheet_part = map_sheet_name_to_part(zin, "Eingabemaske")
        xml_text = zin.read(sheet_part).decode("utf-8")

    if "mc:Ignorable" not in xml_text:
        xml_text = xml_text.replace(
            "<worksheet ",
            (
                "<worksheet "
                'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
                'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
                'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" '
                'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" '
                'mc:Ignorable="x14ac xr xr2 xr3" '
            ),
            1,
        )

    with zipfile.ZipFile(source, "r") as zin:
        with zipfile.ZipFile(output, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_part:
                    data = xml_text.encode("utf-8")
                zout.writestr(item, data)

    report = write_numeric_cells(
        output,
        output,
        sheet_name="Eingabemaske",
        cell_values={"B2": 7.7},
        guard_cells=["H2"],
    )

    assert report.written_count == 1

    with zipfile.ZipFile(output, "r") as z:
        out_head = z.read(sheet_part).decode("utf-8")[:1800]

    assert "Ignorable=\"x14ac xr xr2 xr3\"" in out_head
    assert "xmlns:x14ac=" in out_head
    assert "xmlns:xr=" in out_head
    assert "xmlns:xr2=" in out_head
    assert "xmlns:xr3=" in out_head


def test_load_cells_from_json_mapping(tmp_path: Path) -> None:
    path = tmp_path / "cells.json"
    path.write_text('{"b45": 1.25, "C46": 2}', encoding="utf-8")

    loaded = _load_cells_from_json(path)
    assert loaded == {"B45": 1.25, "C46": 2.0}


def test_load_cells_from_csv(tmp_path: Path) -> None:
    path = tmp_path / "cells.csv"
    path.write_text("cell,value\n b45 ,1.25\nC46,2\n", encoding="utf-8")

    loaded = _load_cells_from_csv(path)
    assert loaded == {"B45": 1.25, "C46": 2.0}


def test_load_cells_from_json_text_mapping() -> None:
    loaded = _load_cells_from_json_text('{"b45": 1.25, "C46": 2}')
    assert loaded == {"B45": 1.25, "C46": 2.0}


def test_inject_cells_highlevel_accepts_str_values(tmp_path: Path) -> None:
    source = tmp_path / "source_highlevel.xlsx"
    output = tmp_path / "output_highlevel.xlsx"
    _create_test_workbook(source)

    report = inject_cells(
        source,
        output,
        sheet_name="Eingabemaske",
        cell_values={"b2": "5,5", "B4": 6.6, "B7": None, "B8": float("nan")},
        guard_cells=["H2"],
    )

    assert report.written_count == 2
    assert _read_cell_v_text(output, "B2") == "5.5"
    assert _read_cell_v_text(output, "B4") == "6.6"


def test_normalize_numeric_value() -> None:
    assert normalize_numeric_value(None) is None
    assert normalize_numeric_value("") is None
    assert normalize_numeric_value(" 12,34 ") == 12.34
    assert normalize_numeric_value(float("nan")) is None
    assert normalize_numeric_value("abc") is None


def test_to_excel_serial_with_datetime_like() -> None:
    naive = datetime(2025, 1, 1, 0, 0, 0)
    aware = datetime(2025, 1, 1, 0, 0, 0, tzinfo=timezone.utc)

    serial_naive = to_excel_serial(naive)
    serial_aware = to_excel_serial(aware)

    assert serial_naive is not None
    assert serial_aware is not None
    assert abs(serial_naive - serial_aware) < 1e-9


def test_build_column_map_and_merge() -> None:
    col_b = build_column_cell_map(2, [1.0, None, "3,5"], 45)
    col_c = build_column_cell_map(3, [10, 20], 45)
    merged = merge_cell_maps(col_b, col_c)

    assert col_b == {"B45": 1.0, "B47": 3.5}
    assert col_c == {"C45": 10.0, "C46": 20.0}
    assert merged["B45"] == 1.0
    assert merged["B47"] == 3.5
    assert merged["C46"] == 20.0


def test_inject_cells_sets_calcpr_recalc_flags(tmp_path: Path) -> None:
    source = tmp_path / "source_recalc.xlsx"
    output = tmp_path / "output_recalc.xlsx"
    _create_test_workbook(source)

    report = inject_cells(
        source,
        output,
        sheet_name="Eingabemaske",
        cell_values={"B2": 5.5},
    )

    assert report.written_count == 1

    with zipfile.ZipFile(output, "r") as z:
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        calc_pr = wb_root.find("x:calcPr", NS)
        assert calc_pr is not None
        assert calc_pr.attrib.get("fullCalcOnLoad") == "1"
        assert calc_pr.attrib.get("forceFullCalc") == "1"
        assert calc_pr.attrib.get("calcOnSave") == "1"
        assert calc_pr.attrib.get("calcCompleted") == "0"


def test_apply_recalc_policy_can_clear_formula_cached_values(tmp_path: Path) -> None:
    source = tmp_path / "source_formula_cache.xlsx"
    _create_test_workbook(source)

    # Ensure formula cell has cached value node for the test scenario
    with zipfile.ZipFile(source, "r") as zin:
        sheet_part = map_sheet_name_to_part(zin, "Eingabemaske")
        xml_text = zin.read(sheet_part).decode("utf-8")

    if "<f>LET(" in xml_text and "<c r=\"H2\"" in xml_text and "<v>" not in xml_text:
        xml_text = xml_text.replace("<c r=\"H2\"><f>", "<c r=\"H2\"><f>", 1)
        xml_text = xml_text.replace("</f></c>", "</f><v>999</v></c>", 1)

        with zipfile.ZipFile(source, "r") as zin:
            temp = tmp_path / "patched.xlsx"
            with zipfile.ZipFile(temp, "w") as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == sheet_part:
                        data = xml_text.encode("utf-8")
                    zout.writestr(item, data)
        source = temp

    apply_recalc_policy(
        source,
        remove_calc_chain_file=False,
        set_full_calc_on_load=False,
        clear_formula_cached_values=True,
        clear_formula_cache_sheets=("Eingabemaske",),
    )

    with zipfile.ZipFile(source, "r") as z:
        sheet_part = map_sheet_name_to_part(z, "Eingabemaske")
        root = ET.fromstring(z.read(sheet_part))
        formula_cell = root.find(".//x:c[@r='H2']", NS)
        assert formula_cell is not None
        assert formula_cell.find("x:f", NS) is not None
        assert formula_cell.find("x:v", NS) is None
