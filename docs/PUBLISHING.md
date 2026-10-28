# Publishing to PyPI

This document describes the release flow for `xlinject`.

## One-time setup

The project already exists on [PyPI](https://pypi.org/project/xlinject/).
Version `0.2.0b2` was uploaded on 2026-04-01, verified on 2026-10-28.
The steps below describe publisher configuration for maintainers.

1. Confirm access to the existing PyPI project (`xlinject`).
2. In PyPI, configure a Trusted Publisher for this GitHub repository:
   - Owner: `KyleDerZweite`
   - Repository: `xlinject`
   - Workflow file: `.github/workflows/publish-pypi.yml`
   - Environment (recommended): `pypi`
3. In GitHub repository settings, create the `pypi` environment.
4. Optionally add required reviewers for the `pypi` environment.

## Local preflight checks

```bash
uv sync --locked --dev
uv run --locked pre-commit run --all-files
uv run --locked ruff check .
uv run --locked mypy .
uv run --locked pytest -q
uv build
uv run --locked twine check dist/*
```

## Release process

1. Bump `version` in `pyproject.toml` and run `uv lock` to update package metadata.
2. Update release notes/changelog.
3. Commit and push to `main`.
4. Create a GitHub Release with a matching tag, such as `v0.2.0b3` for a beta. Mark beta releases as prereleases.
5. The `publish-pypi.yml` workflow runs automatically and publishes to PyPI.

## Verify installation

```bash
python -m pip install --upgrade --pre xlinject
python -c "import xlinject; print(xlinject.__version__)"
```

The `--pre` flag is required while only prereleases are available. Consumer
installation with pip is supported; repository development uses uv.

## TestPyPI (optional)

You can duplicate the publish workflow for TestPyPI before your first real release.
Use `repository-url: https://test.pypi.org/legacy/` in the publish step.

## Security notes

- Do not store PyPI API tokens in repository secrets when using Trusted Publisher.
- Keep example workbooks synthetic and anonymized.
- Never publish customer workbooks or logs with sensitive IDs.
