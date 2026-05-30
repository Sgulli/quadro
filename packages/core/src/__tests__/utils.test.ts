import { describe, expect, it } from "vitest";
import {
  accounting,
  align,
  applyStyle,
  border,
  colLetter,
  currency,
  fill,
  font,
  formatHeaderFooterSection,
  numFmt,
  resolveNumberFormat,
  Styles,
  style,
  toExcelAlignment,
  toExcelBorder,
  toExcelFill,
  toExcelFont,
} from "../utils.js";

describe("utils barrel", () => {
  it("re-exports coords helpers", () => {
    expect(colLetter(1)).toBe("A");
  });

  it("re-exports style helpers", () => {
    expect(style({ font: { bold: true } })).toEqual({ font: { bold: true } });
    expect(border.thinBlack).toBeDefined();
    expect(align.center).toBeDefined();
    expect(Styles.header).toBeDefined();
    expect(font({ bold: true })).toEqual({ font: { bold: true } });
    expect(fill({ type: "solid", color: "FF0000" })).toEqual({
      fill: { type: "solid", color: "FF0000" },
    });
    expect(numFmt("percent")).toEqual({ numberFormat: "percent" });
  });

  it("re-exports converter functions", () => {
    expect(toExcelFont({})).toEqual({});
    expect(toExcelFill({ type: "solid", color: "FF0000" })).toBeDefined();
    expect(toExcelBorder({})).toEqual({});
    expect(toExcelAlignment({})).toEqual({});
    expect(resolveNumberFormat("general")).toBe("General");
    expect(formatHeaderFooterSection({})).toBe("");
  });

  it("re-exports format helpers", () => {
    expect(currency("$")).toBe('"$"#,##0.00');
    expect(accounting("$")).toBe('"$"#,##0.00;("$"#,##0.00);"-"');
  });

  it("re-exports applyStyle", () => {
    const cell: Record<string, unknown> = {};
    applyStyle(cell as never, { font: { bold: true } });
    expect(cell.font).toBeDefined();
  });
});
