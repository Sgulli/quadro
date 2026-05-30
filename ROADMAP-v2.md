# QQuadro Core – Improvement Roadmap

## Executive Summary

The strongest ideas in QQuadro are:

- Typed workbook generation
- Tuple-based coordinates
- Inferred column metadata
- Fluent API with TypeScript-first ergonomics

The main risk is that the API is evolving into both:

1. A high-level business reporting DSL
2. A low-level Excel abstraction

This increases complexity, expands the public API surface, and may make long-term maintenance more difficult.

---

# 1. Introduce a First-Class Range Abstraction

## Current Problem

The API exposes many specialized methods:

```ts
addListValidation(...)
addRangeValidation(...)
addCellIsRule(...)
addExpressionRule(...)
addTop10Rule(...)
addAboveAverageRule(...)
addContainsTextRule(...)
```

This causes API surface growth as new Excel features are added.

## Proposed Direction

Make `Range` the central abstraction.

```ts
sheet
  .range(...)
  .validation(...)
  .style(...)
  .conditional(...)
```

## Benefits

- Smaller API surface
- Better composability
- Easier maintenance
- More consistent mental model

---

# 2. Reduce Hidden Builder State

## Current Problem

Methods such as:

```ts
sheet.columnRange("score")
sheet.columnIndex("score")
```

depend on previous calls like:

```ts
sheet.headers(...)
sheet.addRows(...)
```

This creates implicit coupling and ordering constraints.

## Proposed Direction

Introduce explicit column definitions.

```ts
const columns = sheet.columns({
  score: {},
  name: {},
});
```

Usage:

```ts
columns.score.range()
columns.score.index()
```

## Benefits

- More predictable behavior
- Better discoverability
- Easier static analysis
- Reduced runtime surprises

---

# 3. Invest in a Formula AST

## Current Problem

Formula helpers can grow indefinitely:

```ts
F.sum(...)
F.avg(...)
F.if(...)
F.concat(...)
```

Eventually users still need arbitrary Excel formulas.

## Proposed Direction

Separate:

```ts
Formula.sum(...)
```

from:

```ts
Formula.raw(...)
```

Long-term goal:

- Typed formula AST
- Formula composition
- Formula validation
- Formula transformations

## Benefits

- Better type safety
- Easier refactoring
- Advanced tooling opportunities

---

# 4. Improve Coordinate Ergonomics

## Current Problem

Tuple coordinates are safe but difficult to read.

```ts
[4, 2, 4, 10]
```

## Proposed Direction

Introduce semantic helpers.

```ts
Range.column("score", 2, 10)
```

or

```ts
range(
  col("score"),
  row(2),
  col("score"),
  row(10),
)
```

## Benefits

- Improved readability
- Self-documenting code
- Reduced mistakes

---

# 5. Move Toward Schema-First Sheets

## Current Problem

The API is currently data-driven.

```ts
sheet.addRows([
  { name: "Alice", score: 85 }
])
```

Column structure is inferred afterwards.

## Proposed Direction

Define sheets using schemas.

```ts
const schema = defineSheet({
  name: text(),
  score: number(),
  status: enum(["Pass", "Fail"]),
});
```

## Potential Features

- Automatic validations
- Automatic formatting
- Width inference
- Import validation
- Typed exports
- Stronger editor tooling

## Benefits

This is likely the biggest differentiator versus traditional Excel wrappers.

---

# 6. Improve Architectural Documentation

## Current Problem

Documentation focuses heavily on features and examples.

Architecture is less visible.

## Proposed Direction

Add architecture-focused documentation.

Example:

```text
Workbook
├── Worksheet
├── Range
├── Formula Engine
├── Validation Engine
└── Streaming Writer
```

Document:

- Responsibilities
- Extension points
- Internal boundaries
- Streaming limitations
- Performance guarantees

## Benefits

- Easier onboarding
- Better contributor experience
- Clearer long-term direction

---

# Plugin System (Future)

Potential extension points:

## Validators

```ts
registerValidator(...)
```

## Style Presets

```ts
registerStylePreset(...)
```

## Exporters

```ts
registerExporter(...)
```

## Importers

```ts
registerImporter(...)
```

Benefits:

- Smaller core
- Community extensibility
- Reduced maintenance burden

---

# Recommended Priority Order

## Phase 1

1. Range abstraction
2. Reduce hidden state
3. Architectural documentation

## Phase 2

4. Schema-first sheet definitions
5. Formula AST

## Phase 3

6. Plugin system
7. Ecosystem extensions

---

# Strategic Positioning

QQuadro should avoid becoming another thin Excel wrapper.

The most unique direction appears to be:

- Type-safe spreadsheet generation
- Strong metadata awareness
- Schema-driven workbook design
- Developer-first TypeScript ergonomics

That combination has significantly more differentiation potential than simply exposing more Excel features.
