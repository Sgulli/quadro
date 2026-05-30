import { Workbook } from "@cj-tech-master/excelts";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { SheetBuilder } from "../../builders/sheet-builder.js";
import type { NoteConfig, SheetOptions } from "../../types.js";

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
  describe("tables", () => {
    it("addTable delegates to ws.addTable", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("MyTable", "A1:C10", [{ name: "Name" }, { name: "Score" }]);
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({
        name: "MyTable",
        ref: "A1:C10",
        headerRow: true,
      });
    });

    it("addTable with rows and options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("Sales", "A1:D5", [{ name: "Product" }, { name: "Revenue" }], {
        rows: [
          ["Widget", 100],
          ["Gadget", 200],
        ],
        totalsRow: true,
        style: { theme: "TableStyleMedium9", showRowStripes: true },
      });
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({
        name: "Sales",
        ref: "A1:D5",
        totalsRow: true,
        style: { theme: "TableStyleMedium9", showRowStripes: true },
      });
    });

    it("addTable accepts tuple range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addTable");
      sheet.addTable("Data", [1, 1, 3, 10], [{ name: "A" }, { name: "B" }, { name: "C" }]);
      expect(ws.addTable).toHaveBeenCalled();
      const args = vi.mocked(ws.addTable).mock.calls[0][0];
      expect(args).toMatchObject({ name: "Data", ref: "A1:C10" });
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addTable("T1", "A1:B2", [{ name: "X" }])
        .addTable("T2", "D1:E2", [{ name: "Y" }]);
      expect(result).toBe(sheet);
    });
  });

  describe("hyperlinks", () => {
    it("setCellHyperlink sets cell value with hyperlink", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCellHyperlink("A1", "https://example.com", "Click here", "Visit site");
      const cell = ws.getCell("A1");
      expect(cell.isHyperlink).toBe(true);
      expect(cell.hyperlink).toBe("https://example.com");
    });

    it("setCellHyperlink uses hyperlink as default text", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCellHyperlink("B2", "https://example.com");
      const cell = ws.getCell("B2");
      expect(cell.isHyperlink).toBe(true);
      expect(cell.hyperlink).toBe("https://example.com");
    });

    it("setCellHyperlink accepts tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getCell");
      sheet.setCellHyperlink([3, 5], "https://example.com");
      expect(ws.getCell).toHaveBeenCalledWith("C5");
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .setCellHyperlink("A1", "https://a.com")
        .setCellHyperlink([2, 1], "https://b.com");
      expect(result).toBe(sheet);
    });
  });

  describe("comments / notes", () => {
    it("addNote sets cell note text", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addNote("A1", "Reviewed by John");
      const cell = ws.getCell("A1");
      expect(cell.note).toBeDefined();
    });

    it("addNote accepts tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getCell");
      sheet.addNote([2, 3], "Note text");
      expect(ws.getCell).toHaveBeenCalledWith("B3");
    });

    it("addThreadedComment pushes to ws.threadedComments", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.threadedComments, "push");
      sheet.addThreadedComment("A1", { personId: "{person-1}", text: "Great work!" });
      expect(ws.threadedComments.push).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addNote("A1", "Note 1").addNote([2, 2], "Note 2");
      expect(result).toBe(sheet);
    });
  });

  describe("rich text", () => {
    it("setCellRichText sets rich text value", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.setCellRichText("A1", [{ text: "Hello ", font: { bold: true } }, { text: "World" }]);
      const cell = ws.getCell("A1");
      expect(cell.value).toBeDefined();
    });

    it("setCellRichText accepts tuple address", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getCell");
      sheet.setCellRichText([1, 5], [{ text: "Test" }]);
      expect(ws.getCell).toHaveBeenCalledWith("A5");
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.setCellRichText("A1", [{ text: "Hi" }]);
      expect(result).toBe(sheet);
    });
  });

  describe("images", () => {
    it("addImage delegates to ws.addImage", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addImage");
      sheet.addImage(1, "A1:B5");
      expect(ws.addImage).toHaveBeenCalledWith(1, "A1:B5");
    });

    it("addBackgroundImage delegates to ws.addBackgroundImage", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addBackgroundImage");
      sheet.addBackgroundImage(1);
      expect(ws.addBackgroundImage).toHaveBeenCalledWith(1);
    });

    it("addWatermark delegates to ws.addWatermark", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addWatermark");
      sheet.addWatermark({ imageId: 1, opacity: 0.15 });
      expect(ws.addWatermark).toHaveBeenCalledWith({ imageId: 1, opacity: 0.15 });
    });

    it("removeWatermark delegates to ws.removeWatermark", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "removeWatermark");
      sheet.removeWatermark();
      expect(ws.removeWatermark).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addImage(1, "A1:B5").addBackgroundImage(2);
      expect(result).toBe(sheet);
    });
  });

  describe("sparklines", () => {
    it("addSparklineGroup delegates to ws.addSparklineGroup", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addSparklineGroup");
      sheet.addSparklineGroup({
        type: "line",
        sparklines: [{ dataRef: "A1:C1", cellRef: "D1" }],
      });
      expect(ws.addSparklineGroup).toHaveBeenCalledWith({
        type: "line",
        sparklines: [{ dataRef: "A1:C1", cellRef: "D1" }],
      });
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addSparklineGroup({
        type: "column",
        sparklines: [{ dataRef: "A1:B1", cellRef: "C1" }],
      });
      expect(result).toBe(sheet);
    });
  });

  describe("charts", () => {
    it("addChart delegates to ws.addChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addChart");
      sheet.addChart({ type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addChart).toHaveBeenCalledWith(
        { type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] },
        "A1:D15",
      );
    });

    it("addBarChart delegates to ws.addBarChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addBarChart");
      sheet.addBarChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addBarChart).toHaveBeenCalledWith(
        { series: [{ values: "Sheet1!$B$2:$B$5" }] },
        "A1:D15",
      );
    });

    it("addColumnChart delegates to ws.addColumnChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addColumnChart");
      sheet.addColumnChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addColumnChart).toHaveBeenCalledWith(
        { series: [{ values: "Sheet1!$B$2:$B$5" }] },
        "A1:D15",
      );
    });

    it("addLineChart delegates to ws.addLineChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addLineChart");
      sheet.addLineChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addLineChart).toHaveBeenCalled();
    });

    it("addPieChart delegates to ws.addPieChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addPieChart");
      sheet.addPieChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(ws.addPieChart).toHaveBeenCalled();
    });

    it("addComboChart delegates to ws.addComboChart", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addComboChart");
      sheet.addComboChart(
        {
          groups: [{ type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] }],
        },
        "A1:D15",
      );
      expect(ws.addComboChart).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addChart({ type: "bar", series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15")
        .addLineChart({ series: [{ values: "Sheet1!$B$2:$B$5" }] }, "A1:D15");
      expect(result).toBe(sheet);
    });
  });

  describe("form controls", () => {
    it("addFormCheckbox delegates to ws.addFormCheckbox", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addFormCheckbox");
      sheet.addFormCheckbox("B2:C3", { text: "Approve", checked: true });
      expect(ws.addFormCheckbox).toHaveBeenCalledWith("B2:C3", { text: "Approve", checked: true });
    });

    it("getFormCheckboxes delegates to ws.getFormCheckboxes", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "getFormCheckboxes");
      sheet.addFormCheckbox("A1", { text: "Yes" });
      sheet.getFormCheckboxes();
      expect(ws.getFormCheckboxes).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addFormCheckbox("A1", { text: "X" })
        .addFormCheckbox("B1", { text: "Y" });
      expect(result).toBe(sheet);
    });
  });

  describe("addNote (v0.5 richer)", () => {
    it("accepts NoteConfig with texts", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      const config: NoteConfig = {
        texts: [{ text: "Line 1" }, { text: "Line 2" }],
      };
      sheet.addNote("A1", config);
      const cell = ws.getCell("A1");
      expect(cell.note).toBeDefined();
    });

    it("accepts string directly (backward compat)", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.addNote("B2", "Simple note");
      expect(ws.getCell("B2").note).toBe("Simple note");
    });
  });

  describe("insertRow", () => {
    it("inserts a row at the given position", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "insertRow");
      sheet.insertRow(1, ["a", "b"]);
      expect(ws.insertRow).toHaveBeenCalledWith(1, ["a", "b"]);
    });

    it("inserts a row and updates rowCount", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow([1]);
      expect(sheet.rowCount).toBe(1);
      sheet.insertRow(1, ["a"]);
      expect(sheet.rowCount).toBe(2);
    });

    it("accepts row options", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.insertRow(1, ["data"], { height: 30 });
      expect(ws.getRow(1).height).toBe(30);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.insertRow(1, ["a"]).insertRow(2, ["b"]);
      expect(result).toBe(sheet);
    });
  });

  describe("duplicateRow", () => {
    it("duplicates a row and updates rowCount", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "duplicateRow");
      sheet.addRow([1]);
      sheet.duplicateRow(1, 2, true);
      expect(ws.duplicateRow).toHaveBeenCalledWith(1, 2, true);
      expect(sheet.rowCount).toBe(3);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow([1]);
      const result = sheet.duplicateRow(1, 1);
      expect(result).toBe(sheet);
      expect(sheet.rowCount).toBe(2);
    });
  });

  describe("removeRow", () => {
    it("removes rows and updates rowCount", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceRows");
      sheet.addRows([[1], [2], [3]]);
      expect(sheet.rowCount).toBe(3);
      sheet.removeRow(2, 2);
      expect(ws.spliceRows).toHaveBeenCalledWith(2, 2);
      expect(sheet.rowCount).toBe(1);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([[1], [2]]);
      const result = sheet.removeRow(1, 1);
      expect(result).toBe(sheet);
    });
  });

  describe("insertColumn / removeColumn", () => {
    it("insertColumn delegates to ws.spliceColumns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceColumns");
      sheet.insertColumn(2, 1);
      expect(ws.spliceColumns).toHaveBeenCalledWith(2, 1);
    });

    it("insertColumn with data", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceColumns");
      sheet.insertColumn(2, 1, ["a"], ["b"]);
      expect(ws.spliceColumns).toHaveBeenCalledWith(2, 1, ["a"], ["b"]);
    });

    it("removeColumn delegates to ws.spliceColumns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "spliceColumns");
      sheet.removeColumn(3, 2);
      expect(ws.spliceColumns).toHaveBeenCalledWith(3, 2);
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.insertColumn(1, 1).removeColumn(2, 1);
      expect(result).toBe(sheet);
    });
  });

  describe("page breaks", () => {
    it("addPageBreak adds a row page break", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      const row = ws.getRow(5);
      vi.spyOn(row, "addPageBreak");
      sheet.addPageBreak(5);
      expect(row.addPageBreak).toHaveBeenCalled();
    });

    it("addColumnPageBreak adds a column page break", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      const col = ws.getColumn(3);
      vi.spyOn(col, "addPageBreak");
      sheet.addColumnPageBreak(3);
      expect(col.addPageBreak).toHaveBeenCalled();
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet.addPageBreak(10).addColumnPageBreak(5);
      expect(result).toBe(sheet);
    });
  });

  describe("addIgnoredError", () => {
    it("pushes to ws.ignoredErrors", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.ignoredErrors, "push");
      sheet.addIgnoredError("A1:A100", { numberStoredAsText: true });
      expect(ws.ignoredErrors.push).toHaveBeenCalledWith({
        ref: "A1:A100",
        numberStoredAsText: true,
      });
    });

    it("accepts multiple error types", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws.ignoredErrors, "push");
      sheet.addIgnoredError("C1:C10", {
        formula: true,
        evalError: true,
        emptyCellReference: true,
      });
      expect(ws.ignoredErrors.push).toHaveBeenCalledWith({
        ref: "C1:C10",
        formula: true,
        evalError: true,
        emptyCellReference: true,
      });
    });

    it("methods are chainable", () => {
      const { sheet } = makeSheet({ name: "Test" });
      const result = sheet
        .addIgnoredError("A1:A10", { numberStoredAsText: true })
        .addIgnoredError("B1:B10", { formula: true });
      expect(result).toBe(sheet);
    });
  });

  describe("worksheet getter", () => {
    it("returns the underlying excelts worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      expect(sheet.worksheet).toBe(ws);
    });
  });
});
