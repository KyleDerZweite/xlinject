# Contributing

Thanks for contributing to `xlinject`.

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
uv run pytest
```

## Contribution rules

- Keep changes focused and small.
- Update relevant docs under `docs/` when behavior or plan changes.
- Add tests with behavior changes.
- Avoid unrelated cleanup in feature PRs.

## Branch naming

Suggested prefixes:

- `feat/`
- `fix/`
- `docs/`
- `chore/`
