# Contributing

Thanks for contributing to `xlinject`.

Current direction: API-first Python library with one optional generic CLI command.

## License and contributions

This project is licensed under GPLv3 (or later).
By submitting contributions, you agree your changes are provided under the same license.

## Setup

```bash
uv sync --dev
```

## Before opening a PR

```bash
uv run ruff check .
uv run mypy .
uv run pytest -q
uv build
uv run twine check dist/*
```

## Contribution rules

- Keep changes focused and small.
- Update relevant docs under `docs/` when behavior or plan changes.
- Add tests with behavior changes.
- Avoid unrelated cleanup in feature PRs.
- Prefer extending reusable high-level helpers over adding app-specific logic.
- Keep CLI changes minimal and generic. The library API is the primary interface.
- Use anonymized/synthetic data only in docs, tests, and issue reports.

## API-first guidance

- Favor `inject_cells` + helper functions (`build_column_cell_map`, `merge_cell_maps`, `to_excel_serial`).
- Preserve formula and workbook metadata safety guarantees.
- If recalculation behavior changes, include tests covering workbook `calcPr` and cache behavior.

## Branch naming

Suggested prefixes:

- `feat/`
- `fix/`
- `docs/`
- `chore/`
