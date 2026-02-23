# Architecture Notes (Planned)

## Core design goal

Mutate only targeted XML nodes while preserving all unrelated workbook content byte-for-byte whenever feasible.

## Workbook targeting strategy

1. Parse `xl/workbook.xml` to locate sheet metadata
2. Parse `xl/_rels/workbook.xml.rels` to resolve relationship targets
3. Map human sheet name -> worksheet part path (`xl/worksheets/sheetN.xml`)

## Read path

- Parse worksheet XML and find target `<c r=\"A1\">`
- Resolve value from `<v>` and cell `t` type
- If `t=\"s\"`, resolve through `xl/sharedStrings.xml`

## Write path (MVP)

- Ensure target row/cell nodes exist
- Use `t=\"inlineStr\"` with `<is><t>...</t></is>` for direct write
- Preserve surrounding row and worksheet structures

## Non-goals for scaffold phase

- No formula authoring
- No style/table/pivot manipulation
- No full workbook object model
