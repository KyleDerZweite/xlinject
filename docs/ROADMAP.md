# Roadmap

Current release stage: **alpha testing/release**.
The project is usable and published-flow ready, but additional hardening is planned.

## Phase 1 - MVP (Proof of Concept)

- [x] Add XLSX zip wrapper (in-memory read/repack)
- [x] Implement sheet name resolution to worksheet XML target
- [x] Implement targeted sentinel replacement in a single-column range (numeric `<v>` writes)
- [x] Implement direct A1-based numeric cell writes
- [x] Provide CLI for direct writes from JSON/CSV mapping files
- [x] Add guard-cell protection checks for formula integrity
- [x] Preserve `mc:Ignorable` namespace declarations required by modern Excel files
- [x] Apply workbook recalculation policy (`calcPr` flags + optional cache clearing)
- [ ] Extend read path for shared strings and additional cell types
- [ ] Extend write path to broader value types where needed

## Phase 2 - Ranges and Sparse Matrices

- Add range write support (example: `B1:B35000`)
- Create missing `<row>` / `<c>` nodes for sparse sheets
- Enforce deterministic cell ordering inside rows

## Phase 3 - Performance and Refinement

- Evaluate `lxml` or lower-level parsers for large sheets
- [x] Remove `xl/calcChain.xml` on save to trigger recalc
- Benchmark with large and sparse workbooks

## Phase 4 - Library Release

- Finalize package structure and public API
- [x] Add CI workflow for lint/type/test
- [x] Add release workflow for tagged builds/publish
- [ ] Publish first public PyPI release (production index)
