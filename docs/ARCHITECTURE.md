# Architecture Notes

## Core design goal

Mutate only targeted worksheet XML nodes while preserving workbook structure, formula semantics,
formatting, and metadata whenever feasible.

## Workbook targeting strategy

1. Parse `xl/workbook.xml` to locate sheet metadata
2. Parse `xl/_rels/workbook.xml.rels` to resolve relationship targets
3. Map human sheet name -> worksheet part path (`xl/worksheets/sheetN.xml`)

## Read path

- Parse worksheet XML and find target `<c r=\"A1\">`
- Resolve value from `<v>` and cell `t` type
- Shared-string decoding is planned for broader read support

## Write path

- Ensure target row/cell nodes exist
- Write numeric values into `<v>` nodes without object-model reserialization
- Block formula overwrite by default (`allow_formula_overwrite=False`)
- Support guard cells to verify signatures did not change
- Preserve `mc:Ignorable` namespace declarations for compatibility with modern Excel files
- Apply workbook recalculation policy (`calcPr` flags + optional formula cache clearing)

## Non-goals

- No formula authoring
- No style/table/pivot manipulation
- No full workbook object model

## Integration pattern

1. Read config and fetch interval data in your application code.
2. Build a deterministic `{A1_ref: numeric_value}` mapping.
3. Call `inject_cells` once per worksheet for safer high-level orchestration.
4. Keep a small set of guard cells around critical formulas.
