# Roadmap

Current release stage: **early development (alpha)**. The project is not fully released yet.

## Phase 1 - MVP (Proof of Concept)

- [x] Add XLSX zip wrapper (in-memory read/repack)
- [x] Implement sheet name resolution to worksheet XML target
- [x] Implement targeted sentinel replacement in a single-column range (numeric `<v>` writes)
- [x] Add guard-cell protection checks for formula integrity
- [ ] Extend read path for shared strings and additional cell types
- [ ] Extend write path to broader value types where needed

## Phase 2 - Ranges and Sparse Matrices

- Add range write support (example: `B1:B35000`)
- Create missing `<row>` / `<c>` nodes for sparse sheets
- Enforce deterministic cell ordering inside rows

## Phase 3 - Performance and Refinement

- Evaluate `lxml` or lower-level parsers for large sheets
- Remove `xl/calcChain.xml` on save to trigger recalc
- Benchmark with large and sparse workbooks

## Phase 4 - Library Release

- Finalize package structure and public API
- [x] Add CI workflow for lint/type/test
- Add release workflow for tagged builds/publish
- Publish to PyPI
