# AGENTS.md

This repository is prepared for human + AI collaboration.

## Current phase

- MVP implemented for production sentinel replacement.
- CLI/API available for single-column range replacement with guard-cell checks.
- Contributors should avoid adding unplanned code paths without first updating `docs/ROADMAP.md`.

## Working agreements for coding agents

1. Keep edits surgical and scope-limited to the requested task.
2. Prefer preserving existing structure and naming.
3. Do not introduce broad refactors during feature work.
4. When touching behavior, update docs in `docs/` in the same change.
5. If changing workbook XML write logic in future phases, include tests that validate unchanged neighboring XML structure.
6. Do not add dependencies without documenting rationale in PR description.
7. Avoid em dashes (`--`) and avoid emoji usage in repository text/docs.

## Implementation guardrails (future code)

- Prioritize XML node-level mutation over full workbook reconstruction.
- Preserve unsupported/unknown tags and attributes exactly.
- Keep generated XML ordering deterministic and Excel-compatible.
- Treat dynamic array and formula metadata as immutable unless explicitly targeted.

## Release and packaging expectations

- Use `uv` for dependency and environment management.
- Keep the package install/build flow in `pyproject.toml`.
- Prefer small, incremental PRs aligned with roadmap phases.
