# Usage

## Install / prepare

```bash
uv sync --dev
```

## CLI: replace `-1` placeholders

```bash
uv run xlinject-replace \
  --input /path/to/source.xlsx \
  --output /path/to/output.xlsx \
  --sheet Eingabemaske \
  --range C45:C35181 \
  --values-file /path/to/values.txt \
  --sentinel -1 \
  --guard-cells B35188,C35188,B35196,C35196
```

`values.txt` format:

```text
12.34
11.90
7.05
```

Values are consumed in order and written only where the current cell value equals the sentinel.

## Python API

```python
from pathlib import Path
from xlinject import replace_sentinel_in_column_range

report = replace_sentinel_in_column_range(
    Path("source.xlsx"),
    Path("output.xlsx"),
    sheet_name="Eingabemaske",
    range_ref="C45:C35181",
    values=[12.34, 11.9, 7.05],
    sentinel=-1.0,
    guard_cells=["B35188", "C35188", "B35196", "C35196"],
)

print(report)
```

## Integration pattern on your other machine

1. Keep this repo next to your existing data-fetch script.
2. Let your fetch script write numeric values to a text file (one per line).
3. Call `xlinject-replace` after fetch completes.
4. Continue your normal workflow with the generated output workbook.
