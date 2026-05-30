import { cellRef, colLetter, rangeRef } from "./coords.js";

export function col(ref: number | string): string {
  return typeof ref === "number" ? colLetter(ref) : ref;
}

export function row(n: number): number {
  return n;
}

function colToNum(c: string): number {
  let n = 0;
  for (let i = 0; i < c.length; i++) {
    n = n * 26 + (c.charCodeAt(i) - 64);
  }
  return n;
}

function parseCell(ref: string): { col: number; row: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) throw new Error(`[Range] Invalid cell reference: ${ref}`);
  return { col: colToNum(match[1]), row: Number.parseInt(match[2], 10) };
}

export const Range = {
  cell(c: number | string, r: number): string {
    return typeof c === "number" ? cellRef(c, r) : `${c}${r}`;
  },

  column(c: number | string, fromRow: number, toRow: number): string {
    const letter = typeof c === "number" ? colLetter(c) : c;
    return `${letter}${fromRow}:${letter}${toRow}`;
  },

  rect(fromCol: number | string, fromRow: number, toCol: number | string, toRow: number): string {
    const fc = typeof fromCol === "number" ? colLetter(fromCol) : fromCol;
    const tc = typeof toCol === "number" ? colLetter(toCol) : toCol;
    return `${fc}${fromRow}:${tc}${toRow}`;
  },

  fromTuple(tuple: [number, number, number, number]): string {
    return rangeRef(tuple[0], tuple[1], tuple[2], tuple[3]);
  },

  fullColumn(c: number | string, startRow = 1, endRow = 1_048_576): string {
    return Range.column(c, startRow, endRow);
  },

  row(rowNum: number, fromCol: number | string = 1, toCol: number | string = 16_384): string {
    const fc = typeof fromCol === "number" ? colLetter(fromCol) : fromCol;
    const tc = typeof toCol === "number" ? colLetter(toCol) : toCol;
    return `${fc}${rowNum}:${tc}${rowNum}`;
  },

  offset(base: string, rowOffset: number, colOffset: number): string {
    const { col: colNum, row: baseRow } = parseCell(base);
    const newCol = colNum + colOffset;
    const newRow = baseRow + rowOffset;
    if (newCol < 1 || newRow < 1) {
      throw new Error(
        `[Range] Offset produces invalid reference: ${base} + (${rowOffset}, ${colOffset})`,
      );
    }
    return cellRef(newCol, newRow);
  },

  expand(ref: string, rows: number, cols = 0): string {
    const { col: colNum, row: baseRow } = parseCell(ref);
    return rangeRef(colNum, baseRow, colNum + cols, baseRow + rows - 1);
  },
} as const;
