import type { CellStyle, Font, NumberFormat, SheetHeaderFooter } from "./style.js";

// ─── Cell Value ──────────────────────────────────────────────────────────────

export type CellPrimitive = string | number | boolean | Date | null;
export type FormulaValue = { formula: string; result?: CellPrimitive };
export type CellValue = CellPrimitive | FormulaValue;

// ─── Column Definition ───────────────────────────────────────────────────────

export interface ColumnDef {
  /** Header label shown in row 1 */
  header: string;
  /** Unique key used when writing rows via key-value objects */
  key: string;
  /** Column width in character units (default: auto / 15) */
  width?: number;
  /** Style applied to every cell in the column (overridden per-cell) */
  style?: CellStyle;
  /** Style applied only to the header cell */
  headerStyle?: CellStyle;
  /** Whether this column is hidden */
  hidden?: boolean;
  /** Number format shorthand — wraps a `numberFormat` in the column style. */
  format?: NumberFormat;
}

// ─── Merge Region ────────────────────────────────────────────────────────────

export interface MergeRange {
  /** e.g. "A1:C3" or computed from row/col numbers */
  range: string;
  /** Value/formula to write into the top-left cell before merging */
  value?: CellValue;
  /** Style for the merged cell */
  style?: CellStyle;
  /** Row height in points */
  height?: number;
}

// ─── Row ─────────────────────────────────────────────────────────────────────

export type RowData = CellValue[] | Record<string, CellValue>;

export interface RowOptions {
  height?: number;
  style?: CellStyle;
  hidden?: boolean;
  outlineLevel?: number;
}

// ─── Sheet Options ───────────────────────────────────────────────────────────

export interface SheetOptions {
  name: string;
  /** Default column style */
  defaultColStyle?: CellStyle;
  /** Default row height */
  defaultRowHeight?: number;
  /** Tab color (hex ARGB) */
  tabColor?: string;
  /** Freeze panes */
  freeze?: { row?: number; col?: number };
  /** Print / page setup */
  pageSetup?: {
    paperSize?: number; // 9 = A4, 1 = Letter
    orientation?: "portrait" | "landscape";
    fitToPage?: boolean;
    fitToWidth?: number;
    fitToHeight?: number;
    margins?: {
      left?: number;
      right?: number;
      top?: number;
      bottom?: number;
      header?: number;
      footer?: number;
    };
    /** Rows to repeat at top when printing, e.g. "1:3" */
    printTitlesRow?: string;
    /** Columns to repeat at left when printing, e.g. "A:C" */
    printTitlesColumn?: string;
  };
  headerFooter?: SheetHeaderFooter;
  /** Whether the sheet is protected */
  protection?: { password?: string; sheet?: boolean };
  /** Show/hide grid lines in view */
  showGridLines?: boolean;
  /** Zoom level (10–400) */
  zoom?: number;
  /** Sheet visibility state */
  state?: "visible" | "hidden" | "veryHidden";
}

// ─── Workbook Options ────────────────────────────────────────────────────────

export interface WorkbookOptions {
  /** Author name embedded in the file metadata */
  author?: string;
  /** Company embedded in the file metadata */
  company?: string;
  /** Created date (defaults to now) */
  created?: Date;
  /** Document title */
  title?: string;
  /** Document subject */
  subject?: string;
  /** Document keywords */
  keywords?: string;
  /** Document category */
  category?: string;
  /** Document manager */
  manager?: string;
  /** Document description */
  description?: string;
  /** Document language */
  language?: string;
  /** Whether to use shared strings table (reduces file size for repeated strings) */
  useSharedStrings?: boolean;
  /**
   * Stream mode: write directly to a file path.
   * When set, the workbook operates in streaming mode for
   * memory-efficient output of large datasets.
   */
  useStreaming?: boolean;
  /** Base directory for output path containment (prevents path traversal) */
  allowedBase?: string;
}

// ─── Builder Return Types ────────────────────────────────────────────────────

export interface WriteResult {
  /** Absolute path of the file written */
  filePath: string;
  /** Approximate file size in bytes (-1 when streaming, size reported after flush) */
  sizeBytes: number;
}

// ─── Comment / Note Types ────────────────────────────────────────────────────

export interface NoteText {
  text: string;
  font?: Partial<Font>;
}

export interface NoteConfig {
  texts?: NoteText[];
  margins?: { insetmode?: string; inset?: number[] };
  protection?: { locked?: string; lockText?: string };
  editAs?: string;
  anchor?: string;
}

// ─── Sparkline Types (defined locally; not exported from excelts index) ─────

export interface SparklineColor {
  theme?: number;
  rgb?: string;
  tint?: number;
  auto?: boolean;
}

export interface Sparkline {
  /** Data reference (e.g. "Sheet1!B2:G2") */
  dataRef: string;
  /** Anchor cell reference (e.g. "H2") */
  cellRef: string;
}

export interface SparklineGroup {
  type?: SparklineType;
  lineWeight?: number;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  displayXAxis?: boolean;
  displayHidden?: boolean;
  minAxisType?: SparklineAxisType;
  maxAxisType?: SparklineAxisType;
  manualMin?: number;
  manualMax?: number;
  rightToLeft?: boolean;
  colorSeries?: SparklineColor;
  colorNegative?: SparklineColor;
  colorAxis?: SparklineColor;
  colorMarkers?: SparklineColor;
  colorFirst?: SparklineColor;
  colorLast?: SparklineColor;
  colorHigh?: SparklineColor;
  colorLow?: SparklineColor;
  dateAxis?: string;
  sparklines: Sparkline[];
}

export type SparklineType = "line" | "column" | "stacked";

export type SparklineAxisType = "individual" | "group" | "custom";

export interface AddSparklineGroupOptions {
  type: SparklineType;
  sparklines: Sparkline[];
  lineWeight?: number;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  lineColor?: string;
  negativeColor?: string;
  axisColor?: string;
  markerColor?: string;
  highColor?: string;
  lowColor?: string;
  firstColor?: string;
  lastColor?: string;
  minAxisType?: SparklineAxisType;
  maxAxisType?: SparklineAxisType;
  manualMin?: number;
  manualMax?: number;
  displayXAxis?: boolean;
  rightToLeft?: boolean;
  displayEmptyCellsAs?: "gap" | "zero" | "span";
  dateAxis?: string;
}

// ─── Ignored Errors ──────────────────────────────────────────────────────────

export interface IgnoredErrorDef {
  /** Cell reference range, e.g. "A1:B10" */
  ref: string;
  /** Ignore "Number Stored as Text" errors */
  numberStoredAsText?: boolean;
  /** Ignore formula errors */
  formula?: boolean;
  /** Ignore formula range errors */
  formulaRange?: boolean;
  /** Ignore unlocked formula errors */
  unlockedFormula?: boolean;
  /** Ignore empty cell reference errors */
  emptyCellReference?: boolean;
  /** Ignore list data validation errors */
  listDataValidation?: boolean;
  /** Ignore calculated column errors */
  calculatedColumn?: boolean;
  /** Ignore eval errors */
  evalError?: boolean;
  /** Ignore two-digit text year errors */
  twoDigitTextYear?: boolean;
}

// ─── External Workbook References ────────────────────────────────────────────

export interface ExternalLinkInput {
  /** Path or URL to external workbook */
  target: string;
  /** Exposed sheet names in the external workbook */
  sheetNames?: string[];
  /** Cached values for formula references */
  cachedValues?: Record<string, Record<string, string | number | boolean | null>>;
  /** Link mode (default: "External") */
  targetMode?: "External" | "Internal";
}
