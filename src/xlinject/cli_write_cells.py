from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Sequence

from xlinject.injector import write_numeric_cells


def _load_cells_from_json(path: Path) -> dict[str, float]:
    raw = json.loads(path.read_text(encoding="utf-8"))

    if isinstance(raw, dict):
        return {str(cell).strip().upper(): float(value) for cell, value in raw.items()}

    if isinstance(raw, list):
        loaded: dict[str, float] = {}
        for entry in raw:
            if not isinstance(entry, dict):
                raise ValueError("JSON list entries must be objects with 'cell' and 'value'.")
            if "cell" not in entry or "value" not in entry:
                raise ValueError("JSON list entries must include 'cell' and 'value' keys.")
            loaded[str(entry["cell"]).strip().upper()] = float(entry["value"])
        return loaded

    raise ValueError("Unsupported JSON format. Use object mapping or list of {cell, value} objects.")


def _load_cells_from_csv(path: Path) -> dict[str, float]:
    loaded: dict[str, float] = {}
    with path.open("r", encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        if not reader.fieldnames:
            raise ValueError("CSV file is empty or missing header.")

        lower_fields = {name.lower().strip(): name for name in reader.fieldnames}
        if "cell" not in lower_fields or "value" not in lower_fields:
            raise ValueError("CSV header must include 'cell' and 'value' columns.")

        cell_key = lower_fields["cell"]
        value_key = lower_fields["value"]

        for row in reader:
            cell = str(row.get(cell_key, "")).strip().upper()
            value_text = str(row.get(value_key, "")).strip().replace(",", ".")

            if not cell:
                continue
            if value_text == "":
                continue

            loaded[cell] = float(value_text)

    return loaded


def _load_cells(path: Path) -> dict[str, float]:
    suffix = path.suffix.lower()
    if suffix == ".json":
        return _load_cells_from_json(path)
    if suffix == ".csv":
        return _load_cells_from_csv(path)
    raise ValueError("Unsupported cells file extension. Use .json or .csv")


def _load_cells_from_json_text(json_text: str) -> dict[str, float]:
    raw = json.loads(json_text)
    if isinstance(raw, dict):
        return {str(cell).strip().upper(): float(value) for cell, value in raw.items()}
    if isinstance(raw, list):
        loaded: dict[str, float] = {}
        for entry in raw:
            if not isinstance(entry, dict):
                raise ValueError("JSON list entries must be objects with 'cell' and 'value'.")
            if "cell" not in entry or "value" not in entry:
                raise ValueError("JSON list entries must include 'cell' and 'value' keys.")
            loaded[str(entry["cell"]).strip().upper()] = float(entry["value"])
        return loaded
    raise ValueError("Unsupported JSON format. Use object mapping or list of {cell, value} objects.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlinject-write-cells",
        description=(
            "Write numeric values to explicit A1 cells inside an .xlsx file "
            "without workbook object-model reserialization."
        ),
    )

    parser.add_argument("--input", required=True, help="Path to source .xlsx")
    parser.add_argument("--output", required=True, help="Path to output .xlsx")
    parser.add_argument("--sheet", required=True, help="Sheet name (e.g. Eingabemaske)")
    parser.add_argument(
        "--cells-file",
        help="Path to .json or .csv mapping file. JSON object {\"C45\": 12.3} or CSV with columns cell,value.",
    )
    parser.add_argument(
        "--cells-json",
        help="Inline JSON mapping, e.g. '{\"C45\":12.3,\"D45\":11.9}'.",
    )
    parser.add_argument(
        "--guard-cells",
        default="",
        help="Comma-separated guard cell refs that must not change (e.g. H2,B35188)",
    )
    parser.add_argument(
        "--allow-formula-overwrite",
        action="store_true",
        help="Allow writing to cells containing formulas.",
    )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.cells_file and not args.cells_json:
        parser.error("Provide either --cells-file or --cells-json")

    if args.cells_file and args.cells_json:
        parser.error("Use only one of --cells-file or --cells-json")

    if args.cells_file:
        cells_file = Path(args.cells_file)
        if not cells_file.exists():
            parser.error(f"cells file not found: {cells_file}")
        cell_values = _load_cells(cells_file)
    else:
        try:
            cell_values = _load_cells_from_json_text(str(args.cells_json))
        except Exception as exc:
            parser.error(f"Invalid --cells-json payload: {exc}")

    if not cell_values:
        parser.error("No writable cell values found in cells file.")

    guard_cells = [entry.strip().upper() for entry in args.guard_cells.split(",") if entry.strip()]

    report = write_numeric_cells(
        Path(args.input),
        Path(args.output),
        sheet_name=args.sheet,
        cell_values=cell_values,
        guard_cells=guard_cells,
        allow_formula_overwrite=bool(args.allow_formula_overwrite),
    )

    print(f"Output: {report.output_file}")
    print(f"Requested writes: {len(cell_values)}")
    print(f"Written cells: {report.written_count}")
    print(f"Skipped NaN: {report.skipped_nan_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
