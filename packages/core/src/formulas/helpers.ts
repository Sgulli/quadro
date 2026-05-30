import type { CellFormulaValue } from "@cj-tech-master/excelts";
import type { CellPrimitive, CellValue, FormulaValue } from "../types.js";
import * as Fn from "./functions.js";
import { range, rect, ref } from "./refs.js";

export * from "./functions.js";
export { esc, range, rect, ref } from "./refs.js";

export function isFormula(val: CellValue): val is FormulaValue {
  return val !== null && typeof val === "object" && !(val instanceof Date);
}

function normalizeFormula(f: string): string {
  return f.startsWith("=") ? f.slice(1) : f;
}

export function toFormulaValue(v: { formula: string; result?: CellPrimitive }): CellFormulaValue {
  const fv: CellFormulaValue = { formula: normalizeFormula(v.formula) };
  if (v.result) fv.result = v.result;
  return fv;
}

export function toExcelValue(val: CellValue): CellPrimitive | CellFormulaValue {
  if (isFormula(val)) return toFormulaValue(val);
  return val;
}

export const F = {
  ...Fn,
  ref,
  range,
  rect,
  if: Fn.ifExpr,
  switch: Fn.switchExpr,
} as const;
