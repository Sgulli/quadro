import type { CellFormulaValue } from "@cj-tech-master/excelts";
import { colLetter } from "./coords.js";
import type { CellPrimitive, CellValue, FormulaValue } from "./types.js";

type Ref = string | number;

function resolveCol(col: Ref): string {
  return typeof col === "number" ? colLetter(col) : col;
}

function esc(v: string | number): string {
  return typeof v === "string" ? `"${v.replace(/"/g, '""')}"` : String(v);
}

export function ref(col: Ref, row: number): string {
  return `${resolveCol(col)}${row}`;
}

export function range(col: Ref, fromRow: number, toRow: number): string {
  const c = resolveCol(col);
  return `${c}${fromRow}:${c}${toRow}`;
}

export function rect(fromCol: Ref, fromRow: number, toCol: Ref, toRow: number): string {
  return `${resolveCol(fromCol)}${fromRow}:${resolveCol(toCol)}${toRow}`;
}

export function sum(rangeStr: string, result?: number): FormulaValue {
  return { formula: `SUM(${rangeStr})`, result };
}

export function average(rangeStr: string, result?: number): FormulaValue {
  return { formula: `AVERAGE(${rangeStr})`, result };
}

export function count(rangeStr: string, result?: number): FormulaValue {
  return { formula: `COUNT(${rangeStr})`, result };
}

export function max(rangeStr: string, result?: number): FormulaValue {
  return { formula: `MAX(${rangeStr})`, result };
}

export function min(rangeStr: string, result?: number): FormulaValue {
  return { formula: `MIN(${rangeStr})`, result };
}

export function add(...cells: (string | number)[]): FormulaValue {
  return { formula: cells.join("+") };
}

export function sub(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}-${b}` };
}

export function mul(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}*${b}` };
}

export function div(a: string | number, b: string | number): FormulaValue {
  return { formula: `${a}/${b}` };
}

export function pct(current: string | number, previous: string | number): FormulaValue {
  return { formula: `(${current}-${previous})/${previous}` };
}

export function ifExpr(
  condition: string,
  ifTrue: string | number | FormulaValue,
  ifFalse: string | number | FormulaValue,
): FormulaValue {
  const thenVal = typeof ifTrue === "object" && "formula" in ifTrue ? ifTrue.formula : esc(ifTrue);
  const elseVal =
    typeof ifFalse === "object" && "formula" in ifFalse ? ifFalse.formula : esc(ifFalse);
  return { formula: `IF(${condition},${thenVal},${elseVal})` };
}

// ─── Cell value conversion helpers ─────────────────────────────────────────────

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
  ref,
  range,
  rect,
  sum,
  average,
  count,
  max,
  min,
  add,
  sub,
  mul,
  div,
  pct,
  if: ifExpr,
  ifExpr,
} as const;
