# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2026-05-30

### Added

- **Schema-first sheet definitions** with `defineSheet()` API
  - `Schema.text()`, `Schema.number()`, `Schema.date()`, `Schema.boolean()`, `Schema.enum()`, `Schema.currency()`, `Schema.percent()` field builders
  - `SchemaDefinition` and `InferRowType<T>` for fully-typed sheet schemas
  - `validateRow()` for runtime validation with type checks, min/max bounds, enum membership
  - `schemaToColumnDefs()` to convert schema to column definitions with auto-width and default styles
- **Typed column references** with `ColumnMap` / `ColumnRef`
  - `.letter()`, `.index()`, `.range()`, `.cell()`, `.toColumnDef()` methods
  - `createColumnMap()` factory function
- **Fluent range operations** with `RangeBuilder`
  - `.style()`, `.validation()`, `.cellIs()`, `.dataBar()`, `.iconSet()`, `.top10()`, `.aboveAverage()`, `.containsText()`, `.timePeriod()`, `.expression()`, `.colorScale()`, `.merge()`, `.formula()`
  - `sheet.range()` entry point for fluent range API
- **Range namespace** for semantic coordinate helpers
  - `Range.cell()`, `Range.column()`, `Range.rect()`, `Range.fromTuple()`, `Range.fullColumn()`, `Range.row()`, `Range.offset()`, `Range.expand()`
  - `col()` and `row()` helper functions
- **Typed formula AST** with `FormulaNode` / `Expr` / `Formula`
  - Composable formula expressions with type safety
  - `Formula.sum()`, `Formula.avg()`, `Formula.if()`, `Formula.raw()` and more

### Changed

- **BREAKING**: Reorganized source files into subdirectories
  - `builders/` - WorkbookBuilder, SheetBuilder, RangeBuilder, ColumnMap
  - `coords/` - Coordinate utilities and Range namespace
  - `formulas/` - Formula helpers and AST
  - `styles/` - Style system and presets
  - `schema/` - Schema-first definitions
  - `mixins/` - SheetBuilder extensions
- **BREAKING**: Removed unsafe casts and inlined import types
- **BREAKING**: `defineColumns()` now returns `ColumnMap` instead of array
- Improved type safety throughout the codebase
- Updated documentation with new APIs and examples

### Removed

- **BREAKING**: Removed `utils.ts` barrel file (imports should use specific modules)

## [0.3.0] - 2026-05-20

- Data validation: `sheet.addDataValidation()`, `sheet.addListValidation()`, `sheet.addRangeValidation()`
- Conditional formatting: `sheet.addConditionalFormatting()`, `sheet.addCellIsRule()`, `sheet.addExpressionRule()`
- Conditional formatting: `sheet.addDataBar()`, `sheet.addColorScale()`, `sheet.addIconSet()`
- Conditional formatting: `sheet.addTop10Rule()`, `sheet.addAboveAverageRule()`, `sheet.addContainsTextRule()`, `sheet.addTimePeriodRule()`
- Exported `DataValidation`, `ConditionalFormattingOptions`, and related types from `@quadro/core`
- Simplified API: `headers()` replaces `columns()` + `writeHeaders()` in one call
- `merge()` accepts optional `height` property (no separate `rowHeight()` call needed)
- `headers()` accepts optional `height` parameter
- Columns auto-inferred from first object row when not explicitly defined
- `columnRange()` auto-detects header offset (row 1 vs row 2)
- CLI commands exposed via `pnpm cli` and `pnpm example`
- Numeric coordinate API: `setCellRC()`, `mergeRC()`, `styleRangeRC()`, `colWidthRC()`
- Numeric validation/formatting helpers: `cellRef()`, `colRange()`, `rangeRef()`
- `addListValidationRC()`, `addRangeValidationRC()`, `addDataValidationRC()`
- `addCellIsRuleRC()`, `addExpressionRuleRC()`, `addDataBarRC()`, `addColorScaleRC()`
- `addIconSetRC()`, `addTop10RuleRC()`, `addAboveAverageRuleRC()`
- `addContainsTextRuleRC()`, `addTimePeriodRuleRC()`
- Full zero-A1-string sheet building with `for` loops
- Conditional formatting removal: `sheet.removeConditionalFormatting()`
- Rewrote assignments example using pure numeric coordinates (no A1 strings)
- Fixed pre-existing merge crash in assignments (off-by-one `FIXED_COLS`)
- Fixed pre-existing TypeScript errors in seat filter type guards
- `sheet.columnIndex()` resolves column key to 1‑based index for RC API
- `sheet.removeDataValidation()` to remove validation rules
- Updated `@quadro/core` README with RC API docs, package managers, author attribution

## [0.2.0] - unreleased

- Workbook reading: `WorkbookBuilder.load()`, `WorkbookBuilder.fromFile()`
- Sheet-to-JSON: `sheet.toJSON()`, `sheet.toAOA()`, `sheet.addJSON()`, `sheet.addAOA()`
- CSV import/export: `workbook.toCsv()`, `WorkbookBuilder.fromCsv()`, `fromCsvFile()`
- Template workflow: load → modify → re-save
- Command injection fix in debug CLI command (`execFileSync`)
- Path containment via `allowedBase` option
- CLI now has lint + test scripts and handler tests

## [0.1.0] - unreleased

- Initial release of `@quadro/core`
- Workbook builder with fluent API
- Sheet builder with columns, headers, rows, cells, merged cells
- Full style system (font, fill, border, alignment, number formats)
- Preset styles (header, subHeader, totalRow, currency, percent, date, etc.)
- Formula helpers (SUM, AVERAGE, COUNT, MAX, MIN, ADD, SUB, MUL, DIV, PCT, IF)
- Freeze panes, auto-filter
- Page setup, headers/footers
- Streaming mode for large datasets
- Dual ESM/CJS builds
- CLI app (`@quadro/cli`) with `example`, `assignments`, `debug` commands
