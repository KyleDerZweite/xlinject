# Publishing to PyPI

This document describes the release flow for `xlinject`.

## One-time setup

1. Create the project on PyPI (`xlinject`) if it does not exist yet.
2. In PyPI, configure a Trusted Publisher for this GitHub repository:
   - Owner: `KyleDerZweite`
   - Repository: `xlinject`
   - Workflow file: `.github/workflows/publish-pypi.yml`
   - Environment (recommended): `pypi`
3. In GitHub repository settings, create the `pypi` environment.
4. Optionally add required reviewers for the `pypi` environment.

## Local preflight checks

```bash
uv sync --dev
uv run pre-commit run --all-files
uv run ruff check .
uv run mypy .
uv run pytest -q
uv build
uv run twine check dist/*
```

## Release process

1. Bump `version` in `pyproject.toml`.
2. Update release notes/changelog.
3. Commit and push to `main`.
4. Create a GitHub Release with a matching tag (for example `v0.1.1`).
5. The `publish-pypi.yml` workflow runs automatically and publishes to PyPI.

## Verify installation

```bash
python -m pip install --upgrade xlinject
python -c "import xlinject; print(xlinject.__version__)"
```

## TestPyPI (optional)

You can duplicate the publish workflow for TestPyPI before your first real release.
Use `repository-url: https://test.pypi.org/legacy/` in the publish step.

## Security notes

- Do not store PyPI API tokens in repository secrets when using Trusted Publisher.
- Keep example workbooks synthetic and anonymized.
- Never publish customer workbooks or logs with sensitive IDs.
