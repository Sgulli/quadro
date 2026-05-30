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
  describe("headers", () => {
    it("sets column definitions and writes header row", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.headers([{ key: "a", header: "A" }]);
      expect(ws.columns).toHaveLength(1);
      expect(ws.columns[0].key).toBe("a");
      expect(ws.columns[0].width).toBe(15);
      expect(ws.columns[0].hidden).toBe(false);
      expect(ws.addRow).toHaveBeenCalledWith(["A"]);
    });

    it("calls addRow on the underlying worksheet", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "addRow");
      sheet.headers([{ key: "x", header: "X" }]);
      expect(ws.addRow).toHaveBeenCalledWith(["X"]);
    });
  });

  describe("freeze", () => {
    it("sets freeze pane with positional args", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.freeze(3);
      const view = ws.views?.[0] ?? null;
      expect(view?.state).toBe("frozen");
      if (view?.state === "frozen") {
        expect(view.ySplit).toBe(3);
      }
    });

    it("sets freeze pane with object overload", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.freeze({ row: 2, col: 1 });
      const view = ws.views?.[0] ?? null;
      expect(view?.state).toBe("frozen");
      if (view?.state === "frozen") {
        expect(view.ySplit).toBe(2);
        expect(view.xSplit).toBe(1);
      }
    });
  });

  describe("autoFilter", () => {
    it("sets auto filter range", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      sheet.autoFilter("A1:C1");
      expect(ws.autoFilter).toBe("A1:C1");
    });
  });

  describe("autoFitColumns", () => {
    it("auto-fits columns", () => {
      const { ws, sheet } = makeSheet({ name: "Test" });
      vi.spyOn(ws, "autoFitColumns");
      sheet.autoFitColumns();
      expect(ws.autoFitColumns).toHaveBeenCalledWith(undefined, undefined);
    });
  });

  describe("rowCount / columnRange", () => {
    it("tracks row count after addRow", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow([1]);
      expect(sheet.rowCount).toBe(1);
      sheet.addRow([2]);
      expect(sheet.rowCount).toBe(2);
    });

    it("tracks row count after addRows", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([[1], [2], [3]]);
      expect(sheet.rowCount).toBe(3);
    });

    it("columnRange resolves column key to range string", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet
        .headers([
          { key: "name", header: "Name" },
          { key: "score", header: "Score" },
        ])
        .addRows([
          { name: "Alice", score: 85 },
          { name: "Bob", score: 42 },
        ]);
      expect(sheet.columnRange("name")).toBe("A2:A3");
      expect(sheet.columnRange("score")).toBe("B2:B3");
    });

    it("columnRange accepts custom startRow", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.headers([{ key: "x", header: "X" }]);
      sheet.addRow(["a"]);
      expect(sheet.columnRange("x", 1)).toBe("A1:A1");
    });

    it("columnRange throws for unknown key", () => {
      const { sheet } = makeSheet({ name: "Test" });
      expect(() => sheet.columnRange("missing")).toThrow("No column with key");
    });
  });

  describe("auto-inferred columns", () => {
    it("infers columns from first object row", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRow({ name: "Alice", score: 85 });
      expect(sheet.columnRange("name")).toBe("A1:A1");
      expect(sheet.columnRange("score")).toBe("B1:B1");
    });

    it("columnRange works after addRows with inferred columns", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.addRows([
        { product: "Widget", quantity: 10 },
        { product: "Gadget", quantity: 20 },
        { product: "Doohickey", quantity: 30 },
      ]);
      expect(sheet.rowCount).toBe(3);
      expect(sheet.columnRange("product")).toBe("A1:A3");
      expect(sheet.columnRange("quantity")).toBe("B1:B3");
    });

    it("explicit columns take precedence over inference", () => {
      const { sheet } = makeSheet({ name: "Test" });
      sheet.headers([{ key: "custom", header: "Custom" }]);
      sheet.addRow({ name: "Alice", score: 85 });
      expect(sheet.columnRange("custom")).toBe("A2:A2");
      expect(() => sheet.columnRange("name")).toThrow("No column with key");
    });
  });
});
