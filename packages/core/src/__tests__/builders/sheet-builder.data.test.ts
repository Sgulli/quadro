import { Workbook } from "@cj-tech-master/excelts";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { SheetBuilder } from "../../builders/sheet-builder.js";
import type { SheetOptions } from "../../types.js";

function makeSheet(opts: SheetOptions) {
  const wb = new Workbook();
  const ws = wb.addWorksheet(opts.name);
  const sheet = new SheetBuilder(ws, opts);
  return { ws, sheet };
}

beforeEach(() => {
  vi.restoreAllMocks();
});

describe("SheetBuilder", () => {
  describe("addRow", () => {
    it("writes array data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRow(["a", "b"]);
      expect(ws.addRow).toHaveBeenCalledWith(["a", "b"]);
    });

    it("writes object data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.columns([
        { key: "name", header: "Name" },
        { key: "val", header: "Val" },
      ]);
      sheet.addRow({ name: "Alice", val: 42 });
      expect(ws.addRow).toHaveBeenCalledWith({ name: "Alice", val: 42 });
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRow([{ formula: "SUM(A1:A10)" }]);
      expect(ws.addRow).toHaveBeenCalledWith([{ formula: "SUM(A1:A10)" }]);
    });
  });

  describe("addRows", () => {
    it("adds multiple rows", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.addRows([[1], [2], [3]]);
      expect(ws.addRow).toHaveBeenCalledTimes(3);
    });
  });

  describe("setCell", () => {
    it("accepts a tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell([4, 5], 42);
      expect(ws.getCell("D5").value).toBe(42);
    });

    it("sets a value at the given address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("B3", 42);
      expect(ws.getCell("B3").value).toBe(42);
    });

    it("applies style when provided", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "hello", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
    });

    it("skips undefined value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1");
      expect(ws.getCell("A1").value).toBeFalsy();
    });

    it("writes formula values", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("C5", { formula: "SUM(A1:A10)" });
      expect(ws.getCell("C5").value).toEqual({ formula: "SUM(A1:A10)" });
    });
  });

  describe("styleRange", () => {
    it("applies style to a range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "a");
      sheet.setCell("B2", "b");
      sheet.styleRange("A1:B2", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
      expect(ws.getCell("B2").font?.bold).toBe(true);
    });

    it("handles single cell range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCell("A1", "a");
      const result = sheet.styleRange("A1:A1", { font: { bold: true } });
      expect(ws.getCell("A1").font?.bold).toBe(true);
      expect(result).toBe(sheet);
    });
  });

  describe("merge", () => {
    it("merges cells and writes value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "mergeCells");
      sheet.merge({ range: "A1:C1", value: "Title", style: { font: { bold: true } } });
      expect(ws.mergeCells).toHaveBeenCalledWith("A1:C1");
      expect(ws.getCell("A1").value).toBe("Title");
      expect(ws.getCell("A1").font?.bold).toBe(true);
    });

    it("accepts a tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "mergeCells");
      sheet.merge([1, 1, 3, 1], { value: "Title" });
      expect(ws.mergeCells).toHaveBeenCalledWith("A1:C1");
    });
  });

  describe("rowHeight / colWidth", () => {
    it("sets row height", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.rowHeight(1, 30);
      expect(ws.getRow(1).height).toBe(30);
    });

    it("sets column width", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.colWidth("A", 40);
      expect(ws.getColumn("A").width).toBe(40);
    });
  });
});
