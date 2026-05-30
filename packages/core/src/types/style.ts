// ─── Alignment ───────────────────────────────────────────────────────────────

export type HorizontalAlignment =
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed";

export type VerticalAlignment = "top" | "middle" | "bottom" | "distributed" | "justify";

export interface Alignment {
  horizontal?: HorizontalAlignment;
  vertical?: VerticalAlignment;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  textRotation?: number | "vertical";
}

// ─── Border ──────────────────────────────────────────────────────────────────

export type BorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dotted"
  | "dashed"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot"
  | "double";

export interface BorderSide {
  style?: BorderStyle;
  color?: string; // hex without "#", e.g. "FF0000"
}

export interface Border {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
  diagonal?: BorderSide & { up?: boolean; down?: boolean };
}

// ─── Fill ────────────────────────────────────────────────────────────────────

export interface SolidFill {
  type: "solid";
  color: string; // hex ARGB, e.g. "FFFFFF00" for yellow
}

export interface Stop {
  position: number;
  color: string;
}

export interface GradientFill {
  type: "gradient";
  gradient: "angle" | "path";
  degree?: number;
  stops: Stop[];
}

export type Fill = SolidFill | GradientFill;

// ─── Font ────────────────────────────────────────────────────────────────────

export interface Color {
  argb: string;
  theme: number;
}

export interface Font {
  name?: string;
  size?: number;
  family?: number;
  scheme?: "minor" | "major" | "none";
  charset?: number;
  color?: string | Partial<Color>;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
  vertAlign?: "superscript" | "subscript";
  strike?: boolean;
  outline?: boolean;
  condense?: boolean;
  extend?: boolean;
  shadow?: boolean;
}

// ─── Number Format ───────────────────────────────────────────────────────────

export type BuiltinFormat =
  | "general"
  | "number"
  | "integer"
  | "float"
  | "currency"
  | "accountingUSD"
  | "percent"
  | "percentDecimal"
  | "date"
  | "datetime"
  | "time"
  | "scientific"
  | "text";

export type NumberFormat = BuiltinFormat | (string & {}); // custom format string passthrough

export const BUILTIN_FORMATS: Record<BuiltinFormat, string> = {
  general: "General",
  number: "#,##0",
  integer: "0",
  float: "#,##0.00",
  currency: "#,##0.00",
  accountingUSD: '#,##0.00;(#,##0.00);"-"',
  percent: "0%",
  percentDecimal: "0.00%",
  date: "yyyy-mm-dd",
  datetime: "yyyy-mm-dd hh:mm:ss",
  time: "hh:mm:ss",
  scientific: "0.00E+00",
  text: "@",
};

// ─── Cell Style ──────────────────────────────────────────────────────────────

export interface CellStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numberFormat?: NumberFormat;
  protection?: { locked?: boolean; hidden?: boolean };
}

// ─── Sheet-level Header / Footer ─────────────────────────────────────────────

export interface HeaderFooterSection {
  left?: string;
  center?: string;
  right?: string;
}

export interface SheetHeaderFooter {
  oddHeader?: HeaderFooterSection;
  oddFooter?: HeaderFooterSection;
  evenHeader?: HeaderFooterSection;
  evenFooter?: HeaderFooterSection;
}
