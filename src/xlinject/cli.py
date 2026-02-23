from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from xlinject.injector import replace_sentinel_in_column_range


def _load_values_from_file(values_file: Path) -> list[float]:
    values: list[float] = []
    for line in values_file.read_text(encoding="utf-8").splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        values.append(float(cleaned.replace(",", ".")))
    return values


def _load_values_from_arg(values_arg: str) -> list[float]:
    values: list[float] = []
    for raw_value in values_arg.split(","):
        cleaned = raw_value.strip()
        if not cleaned:
            continue
        values.append(float(cleaned.replace(",", ".")))
    return values


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlinject-replace",
        description=(
            "Replace sentinel values in a single-column range inside an .xlsx file "
            "without reserializing the full workbook object model."
        ),
    )

    parser.add_argument("--input", required=True, help="Path to source .xlsx")
    parser.add_argument("--output", required=True, help="Path to output .xlsx")
    parser.add_argument("--sheet", required=True, help="Sheet name (e.g. Eingabemaske)")
    parser.add_argument("--range", dest="range_ref", required=True, help="Single-column range (e.g. C45:C35181)")

    parser.add_argument(
        "--values-file",
        help="Text file with one numeric value per line",
    )
    parser.add_argument(
        "--values",
        help="Comma-separated numeric values (e.g. 1.2,3.4,5.6)",
    )
    parser.add_argument(
        "--sentinel",
        type=float,
        default=-1.0,
        help="Numeric sentinel to replace (default: -1)",
    )
    parser.add_argument(
        "--guard-cells",
        default="",
        help="Comma-separated guard cell refs that must not change (e.g. B35188,C35188)",
    )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.values_file and not args.values:
        parser.error("Provide either --values-file or --values")

    if args.values_file and args.values:
        parser.error("Use only one of --values-file or --values")

    if args.values_file:
        values = _load_values_from_file(Path(args.values_file))
    else:
        values = _load_values_from_arg(args.values)

    guard_cells = [entry.strip().upper() for entry in args.guard_cells.split(",") if entry.strip()]

    report = replace_sentinel_in_column_range(
        Path(args.input),
        Path(args.output),
        sheet_name=args.sheet,
        range_ref=args.range_ref,
        values=values,
        sentinel=args.sentinel,
        guard_cells=guard_cells,
    )

    print(f"Output: {report.output_file}")
    print(f"Replaced cells: {report.replaced_count}")
    print(f"Values consumed: {report.consumed_values}")
    print(f"Remaining sentinel cells in range: {report.untouched_sentinel_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
