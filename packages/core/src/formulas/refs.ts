import { colLetter } from "../coords/coords.js";

type Ref = string | number;

function resolveCol(col: Ref): string {
  return typeof col === "number" ? colLetter(col) : col;
}

export function esc(v: string | number): string {
  if (typeof v === "string") {
    const escaped = v.replace(/"/g, '""');
    return /^[=+\-@]/.test(v) ? `"'${escaped}"` : `"${escaped}"`;
  }
  return String(v);
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
