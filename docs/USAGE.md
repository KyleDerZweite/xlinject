# Usage

Primary interface is the Python API (`inject_cells`, `inject_cells_mixed` + helper utilities).
For architecture and design constraints, refer to `ARCHITECTURE.md`.

## Install / prepare

```bash
uv sync --dev
```

## Python API

```python
from xlinject import inject_cells

report = inject_cells(
    "source.xlsx",
    "output.xlsx",
    sheet_name="Eingabemaske",
    cell_values={"C45": 12.34, "D45": "11,90"},
)

print(report)
```

Notes:

- `inject_cells` is the recommended entry point for most integrations.
- Lower-level APIs (`write_numeric_cells`, sentinel replacement) remain available for advanced cases.

## Python API: mixed text and numeric writes

```python
from xlinject import inject_cells_mixed

report = inject_cells_mixed(
  "source.xlsx",
  "output.xlsx",
  sheet_name="Template",
  cell_values={
    "B10": "BK4S1-0008738",
    "B16": 310,
    "B17": 129,
    "B20": "ja",
  },
  guard_cells=["B19", "B25"],
  validate_sheet_rules=True,
)

print(report)
```

Notes:

- `inject_cells_mixed` writes numbers as numeric `<v>` nodes and strings as `inlineStr` nodes.
- Guard cells are checked before and after the mutation, just like the numeric API.
- When `validate_sheet_rules=True`, direct worksheet validations are checked before writing.
- The current validation helper supports worksheet rules of type `list`, `textLength`, `whole`, and `decimal`.

## Python API: low-level mixed write

```python
from pathlib import Path
from xlinject import write_cells

report = write_cells(
  Path("source.xlsx"),
  Path("output.xlsx"),
  sheet_name="Template",
  cell_values={
    "B10": "BK4S1-0008738",
    "B16": 310,
    "B17": 129,
    "B20": "ja",
  },
  guard_cells=["B19", "B25"],
  validate_sheet_rules=True,
)

print(report)
```

## Python API: direct cell writes

```python
from pathlib import Path
from xlinject import write_numeric_cells

report = write_numeric_cells(
  Path("source.xlsx"),
  Path("output.xlsx"),
  sheet_name="Eingabemaske",
  cell_values={
    "B45": 45717.25,
    "C45": 12.34,
    "D45": 15.67,
  },
  guard_cells=["H2"],
)

print(report)
```

Notes:

- `cell_values` is a mapping of `A1` references to numeric values.
- `NaN` values are skipped.
- Formula cells are protected by default and raise an error if targeted.
- Use `guard_cells` to assert key formulas remain unchanged.

## Quickstart: pass dict directly

```python
from xlinject import inject_cells

report = inject_cells(
  "source.xlsx",
  "output.xlsx",
  sheet_name="Eingabemaske",
  cell_values={
    "C45": 12.34,
    "D45": "11,90",
  },
)

print(report)
```

`inject_cells` accepts plain strings for paths and values. Empty values and `NaN` are skipped.

## Recalculation behavior

`inject_cells` applies a safe recalculation policy by default:

- remove `xl/calcChain.xml`
- set workbook `calcPr` flags for full recalculation on load/save

This is designed to avoid manual "click and Enter" recalculation in Excel after injection.

Advanced options:

```python
inject_cells(
  "source.xlsx",
  "output.xlsx",
  sheet_name="Eingabemaske",
  cell_values={"C45": 12.34},
  set_full_calc_on_load=True,
  remove_calc_chain_after_write=True,
  clear_formula_cached_values=False,
)
```

If Excel still shows stale formula results in a special workbook, try:

```python
inject_cells(
  "source.xlsx",
  "output.xlsx",
  sheet_name="Eingabemaske",
  cell_values={"C45": 12.34},
  clear_formula_cached_values=True,
)
```

## Generalized helpers for larger integrations

When you build data from tables (for example API dataframes), use helper functions
to avoid custom mapping logic in your app code.

```python
from xlinject import (
  build_column_cell_map,
  inject_cells,
  merge_cell_maps,
  to_excel_serial,
)

timestamp_map = build_column_cell_map(
  2,
  timestamp_values,
  45,
  value_transform=to_excel_serial,
)

site_c_map = build_column_cell_map(3, site_c_values, 45)
site_d_map = build_column_cell_map(4, site_d_values, 45)

all_cells = merge_cell_maps(timestamp_map, site_c_map, site_d_map)

inject_cells(
  "source.xlsx",
  "output.xlsx",
  sheet_name="Eingabemaske",
  cell_values=all_cells,
)
```

This keeps application code short and moves conversion details into `xlinject`.

## CLI: direct cell writes from JSON or CSV

```bash
uv run xlinject-write-cells \
  --input /path/to/source.xlsx \
  --output /path/to/output.xlsx \
  --sheet Eingabemaske \
  --cells-file /path/to/cells.json \
  --guard-cells H2,B35188
```

JSON format options:

```json
{
  "C45": 12.34,
  "D45": 11.9
}
```

or

```json
[
  {"cell": "C45", "value": 12.34},
  {"cell": "D45", "value": 11.9}
]
```

CSV format:

```csv
cell,value
C45,12.34
D45,11.90
```

Inline JSON mode (no file needed):

```bash
uv run xlinject-write-cells \
  --input /path/to/source.xlsx \
  --output /path/to/output.xlsx \
  --sheet Eingabemaske \
  --cells-json '{"C45":12.34,"D45":11.90}'
```

## Production checklist

1. Keep source workbook immutable and always write to a new output path.
2. Set guard cells for critical formulas.
3. Validate workbook opens in Excel without recovery message.
4. Remove sensitive IDs and tokens from logs before sharing.
5. Use synthetic sample files for bug reports and pull requests.

## Additional references

- Architecture details: `ARCHITECTURE.md`
- Release planning: `ROADMAP.md`
- Publishing flow: `PUBLISHING.md`

## Integration pattern

1. Keep this repo next to your existing data-fetch script or install from PyPI.
2. Build a `{A1: value}` map from your fetched data.
3. Call `inject_cells(...)` once per worksheet.
4. Continue your normal workflow with the generated output workbook.
