// ─── Overload helper types ────────────────────────────────────────────────────

/** Single cell address: A1 string or `[col, row]` numeric pair. */
export type Addr = string | number[];

/** Cell range: A1 string or `[col1, row1, col2, row2]` numeric quadruple. */
export type CellRange = string | number[];

// ─── Data validation ─────────────────────────────────────────────────────────

/** Grouped input for `SheetBuilder.addRangeValidation()`. */
export interface RangeValidationDef {
  type: "whole" | "decimal" | "date" | "textLength";
  operator:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqual"
    | "lessThanOrEqual";
  formulae: (string | number | Date)[];
  allowBlank?: boolean;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
}
