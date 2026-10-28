# Contributing

Thanks for contributing to `xlinject`.

Current direction: API-first Python library with one optional generic CLI command.

## License and contributions

This project is licensed under GPLv3 (or later).
By submitting contributions, you agree your changes are provided under the same license.

## Setup

```bash
uv sync --locked --dev
```

Use uv 0.11.16 as pinned in `pyproject.toml`. The default Python version is 3.11.
Run `uv lock --upgrade` for intentional dependency updates, then review and commit
`uv.lock` with the manifest changes. Dependabot checks the native `uv` ecosystem
and GitHub Actions weekly.
Ruff uses an explicit rule selection matching the
[previous defaults](https://github.com/astral-sh/ruff/blob/0.15.13/docs/configuration.md).

## Before opening a PR

```bash
uv run --locked ruff check .
uv run --locked mypy .
uv run --locked pytest -q
uv build
uv run --locked twine check dist/*
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
