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

## [1.0.0] - 2026-05-29

### Added

- Release checklist and roadmap finalization
- `RELEASE.md` with release process documentation
- Coverage verification thresholds (83% lines, 78% functions, 75% branches)

### Changed

- JSDoc added to all public exports (WorkbookBuilder, SheetBuilder, formulas)
- Split sheet-builder.ts into sheet-builder + format helpers for better separation of concerns
- Addressed code review issues: path traversal protection, ARGB validation, markdown escaping

## [0.8.0] - 2026-05-28

### Added

- Sheet name validation (length, invalid characters, duplicates)
- `WorkbookBuilder.getSheet()` to retrieve a sheet builder by name

### Changed

- Freeze pane preservation improved
- DX and design improvements across the API

## [0.7.0] - 2026-05-27

### Added

- `WorkbookBuilder.calculate()` — trigger full formula recalculation on next open
- `WorkbookBuilder.registerFunction()` — register custom UDFs for formulas
- `WorkbookBuilder.addSheetFromMarkdown()` — import markdown tables as sheets
- `WorkbookBuilder.toMarkdown()` — export sheet data as markdown tables
- `SheetBuilder.fillFormula()` — write formulas into cells by A1 reference
- `SheetBuilder.fillFormulaRC()` — write formulas by column/row numbers

## [0.6.0] - 2026-05-26

### Added

- 61 new formula builders (342 tests total):
  - **Lookup**: `vlookup`, `hlookup`, `index`, `match`, `xlookup`, `offset`, `indirect`
  - **Conditional**: `and`, `or`, `not`, `switch`, `ifs`, `iferror`, `ifna`
  - **Math**: `round`, `roundup`, `rounddown`, `ceiling`, `floor`, `int`, `abs`, `mod`, `power`, `sqrt`, `log`, `log10`, `exp`, `pi`, `rand`, `randbetween`
  - **Text**: `concat`, `left`, `right`, `mid`, `len`, `upper`, `lower`, `trim`, `substitute`, `text`, `value`, `exact`, `find`, `search`, `rept`
  - **Date**: `date`, `today`, `now`, `year`, `month`, `day`, `hour`, `minute`, `second`, `eomonth`, `datedif`, `networkdays`, `workday`
  - **Aggregate**: `sumif`, `sumifs`, `countif`, `countifs`, `averageif`, `averageifs`, `sumproduct`
  - **Array**: `filter`, `unique`, `sort`, `sortby`, `sequence`, `cse`
- `F` namespace expanded with all new formula helpers

## [0.5.0] - 2026-05-25

### Added

- **Comments & Notes**: `sheet.addNote()` for cell comments, `sheet.addThreadedComment()` for threaded discussions
- **Controls**: `sheet.addFormCheckbox()`, `sheet.getFormCheckboxes()`
- **Sheet Management**: `sheet.insertRow()`, `sheet.duplicateRow()`, `sheet.removeRow()`, `sheet.insertColumn()`, `sheet.removeColumn()`, `sheet.addPageBreak()`, `sheet.addColumnPageBreak()`
- **Data I/O**: `sheet.eachRow()`, `sheet.toJSON()`, `sheet.toAOA()`, `sheet.addJSON()`, `sheet.addAOA()`
- **Ignored Errors**: `sheet.addIgnoredError()` for suppressing Excel warnings
- `coords.ts` module with `cellRef()`, `colLetter()`, `colRange()`, `rangeRef()`, `resolveAddr()`, `resolveRange()`
- `style-presets.ts` module with `Styles`, `applyStyle`, `toExcelFont`, `toExcelFill`, `toExcelBorder`, `toExcelAlignment`
- Benchmark test suite
- Integration test suite

### Changed

- Refactored SheetBuilder into mixin architecture:
  - `mixins/charts.ts` — chart methods
  - `mixins/conditional-formatting.ts` — conditional formatting rules
  - `mixins/data-validation.ts` — data validation rules
  - `mixins/media.ts` — notes, hyperlinks, images, watermarks, sparklines, form controls
- Moved style utilities from `utils.ts` to `style-presets.ts`
- Moved coordinate utilities from `utils.ts` to `coords.ts`

## [0.4.1] - 2026-05-24

### Added

- Consolidated RC API with `Addr`/`CellRange` overloads
- Object-form validation and conditional formatting options
- CI release workflow with npm Trusted Publishing (OIDC)

### Changed

- Renamed package from `@quadro/core` to `@qquadro/core` (npm org registration)
- Updated READMEs with v0.4.1 API changes

### Fixed

- Removed `NODE_AUTH_TOKEN` from CI (using Trusted Publishing)
- Added `--no-git-checks` to publish command (detached HEAD on tag push)

## [0.4.0] - 2026-05-23

### Added

- **Charts API**: 20 chart types
  - Standard: bar, column, line, area, pie, doughnut, scatter, bubble, radar, stock, surface, combo
  - Modern ChartEx: histogram, pareto, waterfall, funnel, treemap, sunburst, boxWhisker, regionMap
- **Sparklines**: `sheet.addSparklineGroup()` — line/column/stacked in-cell mini charts with full styling
- **Images**: `sheet.addImage()`, `sheet.addBackgroundImage()`, `sheet.addWatermark()`, `sheet.removeWatermark()`
- **Hyperlinks**: `sheet.setCellHyperlink()` with text, tooltip, RC variants
- **Rich Text**: `sheet.setCellRichText()` — mixed bold/italic/color within a single cell
- **Named Ranges**: `workbook.addNamedRange()` for workbook-scoped named ranges
- **Tables**: `sheet.addTable()` with structured references and styling
- New types exported: `AddImageRange`, `CellHyperlinkValue`, `CellRichTextValue`, `RichTextRun`, `Sparkline`, `SparklineGroup`, `WatermarkOptions`, chart types
- 19 new tests for all features

### Changed

- Chart support auto-installed from `@cj-tech-master/excelts/chart`

## [0.3.0] - 2026-05-20

### Added

- Data validation: `sheet.addDataValidation()`, `sheet.addListValidation()`, `sheet.addRangeValidation()`
- Conditional formatting: `sheet.addConditionalFormatting()`, `sheet.addCellIsRule()`, `sheet.addExpressionRule()`
- Conditional formatting: `sheet.addDataBar()`, `sheet.addColorScale()`, `sheet.addIconSet()`
- Conditional formatting: `sheet.addTop10Rule()`, `sheet.addAboveAverageRule()`, `sheet.addContainsTextRule()`, `sheet.addTimePeriodRule()`
- Conditional formatting removal: `sheet.removeConditionalFormatting()`
- Exported `DataValidation`, `ConditionalFormattingOptions`, and related types
- Numeric coordinate API: `setCellRC()`, `mergeRC()`, `styleRangeRC()`, `colWidthRC()`
- Numeric validation/formatting helpers: `cellRef()`, `colRange()`, `rangeRef()`
- RC variants for all validation and formatting methods
- Full zero-A1-string sheet building with `for` loops
- `sheet.removeDataValidation()` to remove validation rules
- `sheet.columnIndex()` resolves column key to 1-based index for RC API

### Changed

- Simplified API: `headers()` replaces `columns()` + `writeHeaders()` in one call
- `merge()` accepts optional `height` property (no separate `rowHeight()` call needed)
- `headers()` accepts optional `height` parameter
- Columns auto-inferred from first object row when not explicitly defined
- `columnRange()` auto-detects header offset (row 1 vs row 2)

### Fixed

- Pre-existing merge crash in assignments (off-by-one `FIXED_COLS`)
- Pre-existing TypeScript errors in seat filter type guards

## [0.2.0] - unreleased

### Added

- Workbook reading: `WorkbookBuilder.load()`, `WorkbookBuilder.fromFile()`
- Sheet-to-JSON: `sheet.toJSON()`, `sheet.toAOA()`, `sheet.addJSON()`, `sheet.addAOA()`
- CSV import/export: `workbook.toCsv()`, `WorkbookBuilder.fromCsv()`, `fromCsvFile()`
- Template workflow: load → modify → re-save
- CLI lint + test scripts and handler tests

### Changed

- Converted to turborepo monorepo with core package and citty CLI

### Fixed

- Command injection in debug CLI command (`execFileSync`)
- Path containment via `allowedBase` option

## [0.1.0] - unreleased

### Added

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
- CLI app with `example`, `assignments`, `debug` commands
