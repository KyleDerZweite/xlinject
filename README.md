# xlinject

[![CI](https://github.com/KyleDerZweite/xlinject/actions/workflows/ci.yml/badge.svg)](https://github.com/KyleDerZweite/xlinject/actions/workflows/ci.yml)

A surgical XML injector for `.xlsx` files.

`xlinject` is designed to read and write specific cell values directly inside the XLSX archive while preserving formatting, metadata, and modern Excel dynamic array semantics.

> Status: **Alpha testing/release - API-first**

## Release note

`0.2.0b2` fixes an Excel compatibility issue in the worksheet XML serializer.
Workbooks produced by the previous `xml.etree` round-trip could still be read by ZIP tools
and `openpyxl`, but some files were rejected by Excel itself after write operations.
The serializer path now uses `lxml`, which preserves the required worksheet markup more safely.

## Who this is for

`xlinject` is useful when you have Excel templates with modern formulas (`LET`, `FILTER`, dynamic arrays, custom metadata)
and need to inject measured values without changing workbook structure or formatting.

Typical workflow:

1. Fetch data from an external API.
2. Build a mapping from `A1` cell references to numeric values.
3. Inject values with `xlinject`.
4. Open workbook in Excel with formulas and layout preserved.

## Why this project exists

Object-model-based libraries often deserialize and reserialize full workbook structures. During that process, unsupported XML tags/attributes can be dropped. `xlinject` will instead target specific XML nodes in-place to minimize collateral changes.

## Direct write API

```python
from pathlib import Path
from xlinject import inject_cells

report = inject_cells(
	"source.xlsx",
	"output.xlsx",
	sheet_name="Eingabemaske",
	cell_values={
		"B45": 45717.25,
		"C45": "12,34",
		"D45": 15.67,
	},
	guard_cells=["H2"],
)

print(report)
```

This writes only the listed cells and keeps formula XML intact unless `allow_formula_overwrite=True` is explicitly set.

## Mixed write API

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

`inject_cells_mixed` uses the same XML-first strategy as the numeric writer,
but can also write string cells as `inlineStr` while preserving neighboring XML.
When `validate_sheet_rules=True`, direct worksheet validations such as `list`,
`textLength`, `whole`, and `decimal` are checked before the workbook is mutated.


## Development setup (uv)

### Prerequisites

- Python 3.11+
- [uv](https://docs.astral.sh/uv/)

### Quick start

```bash
uv sync --dev
uv run pre-commit install
```

### Quality checks

```bash
uv run pre-commit run --all-files
uv run ruff check .
uv run mypy .
uv run pytest
```


## License

`xlinject` is open source under the GNU General Public License v3.0 or later (GPL-3.0-or-later).
See [LICENSE](LICENSE) for the full license text.
