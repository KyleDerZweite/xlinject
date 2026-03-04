# AGENTS.md

This repository is prepared for human + AI collaboration.

## Current phase

- API-first library phase with generalized cell-injection helpers.
- One optional generic CLI command is available (`xlinject-write-cells`).
- Recalculation safety policy is part of default high-level write behavior.
- Contributors should avoid adding unplanned code paths without first updating `docs/ROADMAP.md`.

## Working agreements for coding agents

1. Keep edits surgical and scope-limited to the requested task.
2. Prefer preserving existing structure and naming.
3. Do not introduce broad refactors during feature work.
4. When touching behavior, update docs in `docs/` in the same change.
5. If changing workbook XML write logic in future phases, include tests that validate unchanged neighboring XML structure.
6. Do not add dependencies without documenting rationale in PR description.
7. Avoid em dashes (`--`) and avoid emoji usage in repository text/docs.
8. Keep the library API as the primary interface; avoid adding task-specific wrappers in CLI.

## Implementation guardrails (future code)

- Prioritize XML node-level mutation over full workbook reconstruction.
- Preserve unsupported/unknown tags and attributes exactly.
- Keep generated XML ordering deterministic and Excel-compatible.
- Treat dynamic array and formula metadata as immutable unless explicitly targeted.
- Preserve `mc:Ignorable` namespace compatibility on all worksheet rewrites.
- Maintain recalculation safety (`calcChain` handling + workbook `calcPr` policy).

## Release and packaging expectations

- Use `uv` for dependency and environment management.
- Keep the package install/build flow in `pyproject.toml`.
- Prefer small, incremental PRs aligned with roadmap phases.
- Keep PyPI release workflow healthy (CI build + twine checks + trusted publisher flow).
