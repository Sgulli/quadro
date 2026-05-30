import type { Addr, CellRange } from "../types.js";

export function colLetter(n: number): string {
  let r = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    r = String.fromCharCode(65 + rem) + r;
    n = Math.floor((n - 1) / 26);
  }
  return r;
}

export function cellRef(col: number, row: number): string {
  return `${colLetter(col)}${row}`;
}

export function colRange(col: number, startRow: number, endRow: number): string {
  const c = colLetter(col);
  return `${c}${startRow}:${c}${endRow}`;
}

export function rangeRef(col1: number, row1: number, col2: number, row2: number): string {
  return `${cellRef(col1, row1)}:${cellRef(col2, row2)}`;
}

export function resolveAddr(addr: Addr): string {
  return Array.isArray(addr) ? cellRef(addr[0], addr[1]) : addr;
}

export function resolveRange(range: CellRange): string {
  return Array.isArray(range) ? rangeRef(range[0], range[1], range[2], range[3]) : range;
}
