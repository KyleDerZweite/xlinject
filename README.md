# xlinject

[![CI](https://github.com/KyleDerZweite/xlinject/actions/workflows/ci.yml/badge.svg)](https://github.com/KyleDerZweite/xlinject/actions/workflows/ci.yml)

A surgical XML injector for `.xlsx` files.

`xlinject` is designed to read and write specific cell values directly inside the XLSX archive while preserving formatting, metadata, and modern Excel dynamic array semantics.

> Status: **Early development (alpha) — not fully released yet**
>
> Current MVP supports sentinel replacement in a single-column range, but the project is still pre-release.

## Why this project exists

Object-model-based libraries often deserialize and reserialize full workbook structures. During that process, unsupported XML tags/attributes can be dropped. `xlinject` will instead target specific XML nodes in-place to minimize collateral changes.

## Scope (current + planned)

- Implemented: sheet name → internal XML mapping (`workbook.xml` + relationships)
- Implemented: surgical numeric sentinel replacement in a target single-column range
- Implemented: optional guard cells to verify protected formulas did not change
- Planned: broader read support (`sharedStrings` decoding)
- Planned: range and sparse-matrix optimization for very large sheets
- Planned: optional calc chain invalidation on write

## Roadmap

See:

- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/USAGE.md`

## Development setup (uv)

### Prerequisites

- Python 3.11+
- [uv](https://docs.astral.sh/uv/)

### Quick start

```bash
uv sync --dev
```

### Quality checks

```bash
uv run ruff check .
uv run mypy .
uv run pytest
```

## Production replacement command

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

This updates only sentinel cells (for example `-1`) in the specified range and preserves formula cells unless explicitly targeted.

## License

`xlinject` is open source under the GNU General Public License v3.0 or later (GPL-3.0-or-later).
See `LICENSE`.

## Project layout

```text
xlinject/
├── src/xlinject/
├── docs/
├── tests/
├── AGENTS.md
├── CONTRIBUTING.md
├── pyproject.toml
└── README.md
```
