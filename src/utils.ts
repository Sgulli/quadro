import type ExcelJS from "exceljs";
import {
  BUILTIN_FORMATS,
  type Alignment,
  type Border,
  type CellStyle,
  type Fill,
  type Font,
  type NumberFormat,
} from "./types.js";

// ─── Font ────────────────────────────────────────────────────────────────────

export function toExcelFont(font: Font): Partial<ExcelJS.Font> {
  return {
    name: font.name,
    size: font.size,
    bold: font.bold,
    italic: font.italic,
    underline: font.underline as ExcelJS.Font["underline"],
    strike: font.strike,
    color: font.color ? { argb: normalizeArgb(font.color) } : undefined,
    vertAlign: font.vertAlign,
  };
}

// ─── Fill ────────────────────────────────────────────────────────────────────

export function toExcelFill(fill: Fill): ExcelJS.Fill {
  if (fill.type === "solid") {
    return {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: normalizeArgb(fill.color) },
    };
  }

  const stops = fill.stops.map((s) => ({
    position: s.position,
    color: { argb: normalizeArgb(s.color) },
  }));

  if (fill.gradient === "angle") {
    return {
      type: "gradient",
      gradient: "angle",
      degree: fill.degree ?? 0,
      stops,
    };
  }
  return {
    type: "gradient",
    gradient: "path",
    center: { left: 0.5, top: 0.5 },
    stops,
  };
}

// ─── Border ──────────────────────────────────────────────────────────────────

function toExcelBorderSide(
  side: NonNullable<Border["top"]>,
): Partial<ExcelJS.Border> {
  return {
    style: side.style,
    color: side.color ? { argb: normalizeArgb(side.color) } : undefined,
  };
}

export function toExcelBorder(border: Border): Partial<ExcelJS.Borders> {
  const result: Partial<ExcelJS.Borders> = {};
  if (border.top) result.top = toExcelBorderSide(border.top);
  if (border.bottom) result.bottom = toExcelBorderSide(border.bottom);
  if (border.left) result.left = toExcelBorderSide(border.left);
  if (border.right) result.right = toExcelBorderSide(border.right);
  if (border.diagonal) {
    result.diagonal = {
      ...toExcelBorderSide(border.diagonal),
      up: border.diagonal.up ?? false,
      down: border.diagonal.down ?? false,
    };
  }
  return result;
}

// ─── Alignment ───────────────────────────────────────────────────────────────

export function toExcelAlignment(a: Alignment): Partial<ExcelJS.Alignment> {
  return {
    horizontal: a.horizontal as ExcelJS.Alignment["horizontal"],
    vertical: a.vertical as ExcelJS.Alignment["vertical"],
    wrapText: a.wrapText,
    shrinkToFit: a.shrinkToFit,
    indent: a.indent,
    textRotation: a.textRotation as ExcelJS.Alignment["textRotation"],
  };
}

// ─── Number Format ───────────────────────────────────────────────────────────

export function resolveNumberFormat(fmt: NumberFormat): string {
  return BUILTIN_FORMATS[fmt as keyof typeof BUILTIN_FORMATS] ?? fmt;
}

// ─── Apply Style ─────────────────────────────────────────────────────────────

export function applyStyle(
  cell: ExcelJS.Cell,
  style: CellStyle | undefined,
): void {
  if (!style) return;
  if (style.font) cell.font = toExcelFont(style.font);
  if (style.fill) cell.fill = toExcelFill(style.fill);
  if (style.border) cell.border = toExcelBorder(style.border);
  if (style.alignment) cell.alignment = toExcelAlignment(style.alignment);
  if (style.numberFormat) cell.numFmt = resolveNumberFormat(style.numberFormat);
  if (style.protection) cell.protection = style.protection;
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function normalizeArgb(hex: string): string {
  const clean = hex.replace(/^#/, "");
  if (clean.length === 6) return `FF${clean}`;
  if (clean.length === 8) return clean;
  throw new Error(
    `Invalid colour "${hex}" — must be 6-char RGB or 8-char ARGB hex.`,
  );
}

// ─── Preset Styles ───────────────────────────────────────────────────────────

export const Styles = {
  /** Bold white text on blue — ideal for column headers */
  header: {
    font: { bold: true, color: "FFFFFFFF", name: "Arial", size: 11 },
    fill: { type: "solid", color: "FF2B579A" },
    alignment: { horizontal: "center", vertical: "middle" },
    border: { bottom: { style: "medium", color: "FF1A3A6B" } },
  } satisfies CellStyle,

  /** Lighter blue — sub-section headers */
  subHeader: {
    font: { bold: true, name: "Arial", size: 10, color: "FF1F497D" },
    fill: { type: "solid", color: "FFDCE6F1" },
    alignment: { horizontal: "left", vertical: "middle" },
  } satisfies CellStyle,

  /** Yellow-tinted bold — totals / summary rows */
  totalRow: {
    font: { bold: true, name: "Arial", size: 10 },
    fill: { type: "solid", color: "FFFFF2CC" },
    border: {
      top: { style: "thin", color: "FFBFBFBF" },
      bottom: { style: "double", color: "FF000000" },
    },
  } satisfies CellStyle,

  /** Right-aligned currency format */
  currency: {
    numberFormat: "currency",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "right" },
  } satisfies CellStyle,

  /** Right-aligned percentage format (two decimals) */
  percent: {
    numberFormat: "percentDecimal",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "right" },
  } satisfies CellStyle,

  /** Centered date format */
  date: {
    numberFormat: "date",
    font: { name: "Arial", size: 10 },
    alignment: { horizontal: "center" },
  } satisfies CellStyle,

  /** Thin grey border on all four sides */
  boxBorder: {
    border: {
      top: { style: "thin", color: "FFBFBFBF" },
      bottom: { style: "thin", color: "FFBFBFBF" },
      left: { style: "thin", color: "FFBFBFBF" },
      right: { style: "thin", color: "FFBFBFBF" },
    },
  } satisfies CellStyle,

  /** Blue font — financial-model convention for hardcoded inputs */
  inputCell: {
    font: { color: "FF0000FF", name: "Arial", size: 10 },
  } satisfies CellStyle,

  /** Black font — financial-model convention for formula cells */
  formulaCell: {
    font: { color: "FF000000", name: "Arial", size: 10 },
  } satisfies CellStyle,

  /** Green font — financial-model convention for cross-sheet links */
  linkCell: {
    font: { color: "FF008000", name: "Arial", size: 10 },
  } satisfies CellStyle,
} as const;
