import fs from "node:fs";
import path from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { F, Styles, WorkbookBuilder } from "../../index.js";

const outputDir = path.resolve("output/test");
const outputPath = (name: string) => path.join(outputDir, name);

afterAll(() => {
  fs.rmSync(outputDir, { recursive: true, force: true });
});

describe("WorkbookBuilder", () => {
  it("creates a workbook with a single sheet", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Sheet1" }, (sheet) => {
        sheet
          .headers([
            { key: "name", header: "Name", width: 20, headerStyle: Styles.header },
            {
              key: "value",
              header: "Value",
              width: 15,
              style: Styles.currency,
              headerStyle: Styles.header,
            },
          ])

          .addRow({ name: "Alpha", value: 100 })
          .addRow({ name: "Beta", value: 200 });
      })
      .write(outputPath("single-sheet.xlsx"));

    expect(result.filePath).toMatch(/single-sheet\.xlsx$/);
    expect(result.sizeBytes).toBeGreaterThan(0);
    expect(fs.existsSync(result.filePath)).toBe(true);
  });

  it("creates a multi-sheet workbook", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Summary" }, (sheet) => {
        sheet
          .headers([
            { key: "item", header: "Item", width: 15 },
            { key: "qty", header: "Qty", width: 10 },
          ])

          .addRow({ item: "Widget", qty: 42 });
      })
      .addSheet({ name: "Details" }, (sheet) => {
        sheet.setCell("A1", "Hello");
      })
      .write(outputPath("multi-sheet.xlsx"));

    expect(result.filePath).toMatch(/multi-sheet\.xlsx$/);
    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports formula cells", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Formulas" }, (sheet) => {
        sheet.setCell("A1", 10);
        sheet.setCell("A2", 20);
        sheet.setCell("A3", { formula: "=SUM(A1:A2)", result: 30 });
      })
      .write(outputPath("formulas.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports merged cells", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Merged" }, (sheet) => {
        sheet.merge({ range: "A1:C1", value: "Title", style: Styles.header });
      })
      .write(outputPath("merged.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports freeze panes", async () => {
    const result = await new WorkbookBuilder({ author: "Test" })
      .addSheet({ name: "Frozen", freeze: { row: 1 } }, (sheet) => {
        sheet.setCell("A1", "Header").setCell("A2", "Data");
      })
      .write(outputPath("frozen.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports auto-filter", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Filtered" }, (sheet) => {
        sheet
          .headers([
            { key: "x", header: "X", width: 10 },
            { key: "y", header: "Y", width: 10 },
          ])

          .addRow({ x: 1, y: 2 })
          .autoFilter();
      })
      .write(outputPath("filter.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports row options (height, style)", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Styled" }, (sheet) => {
        sheet
          .headers([{ key: "a", header: "A", width: 10 }])

          .addRow({ a: "Tall" }, { height: 30, style: Styles.totalRow });
      })
      .write(outputPath("row-options.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports styleRange", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Range" }, (sheet) => {
        sheet.setCell("A1", "Cell 1").setCell("B1", "Cell 2").styleRange("A1:B1", Styles.boxBorder);
      })
      .write(outputPath("range-style.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports F formula shortcuts", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "F-Sheet" }, (sheet) => {
        sheet
          .headers([
            { key: "a", header: "A", width: 10 },
            { key: "b", header: "B", width: 10 },
            { key: "total", header: "Total", width: 15 },
          ])

          .addRow({ a: 5, b: 10, total: F.add(F.ref("A", 2), F.ref("B", 2)) });
      })
      .write(outputPath("f-formulas.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("supports header/footer", async () => {
    const result = await new WorkbookBuilder()
      .addSheet(
        {
          name: "HF",
          headerFooter: {
            oddHeader: { left: "Title", right: "&D" },
            oddFooter: { center: "Page &P" },
          },
        },
        (sheet) => {
          sheet.setCell("A1", "Data");
        },
      )
      .write(outputPath("header-footer.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("returns buffer via toBuffer()", async () => {
    const buf = await new WorkbookBuilder({ author: "BufferTest" })
      .addSheet({ name: "Buf" }, (sheet) => {
        sheet.setCell("A1", "Hello");
      })
      .toBuffer();

    expect(Buffer.isBuffer(buf)).toBe(true);
    expect(buf.length).toBeGreaterThan(0);
  });

  it("exposes excelts workbook via getter", () => {
    const builder = new WorkbookBuilder();
    expect(builder.workbook).toBeDefined();
    expect(typeof builder.workbook.addWorksheet).toBe("function");
  });

  it("addSheet with callback returns this for chaining", () => {
    const builder = new WorkbookBuilder();
    const result = builder.addSheet({ name: "Imperative" }, (sheet) => {
      sheet.setCell("A1", "hello");
    });
    expect(result).toBe(builder);
    expect(builder.sheets[0]).toBeDefined();
    expect(typeof builder.sheets[0].setCell).toBe("function");
  });

  it("getSheet retrieves a sheet by name", () => {
    const builder = new WorkbookBuilder();
    builder.addSheet({ name: "Data" }, () => {});
    expect(builder.getSheet("Data")).toBeDefined();
    expect(builder.getSheet("Data")?.name).toBe("Data");
    expect(builder.getSheet("Missing")).toBeUndefined();
  });

  it("sheet retrieves a sheet by index", () => {
    const builder = new WorkbookBuilder();
    builder.addSheet({ name: "First" }, () => {});
    builder.addSheet({ name: "Second" }, () => {});
    expect(builder.sheet(0).name).toBe("First");
    expect(builder.sheet(1).name).toBe("Second");
    expect(() => builder.sheet(99)).toThrow("No sheet at index");
  });

  it("tabColor applies to sheet", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Colored", tabColor: "FFFF0000" }, (sheet) => {
        sheet.setCell("A1", "Red tab");
      })
      .write(outputPath("tab-color.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("writes object-array rows", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Batch" }, (sheet) => {
        sheet
          .headers([
            { key: "name", header: "Name", width: 15 },
            { key: "age", header: "Age", width: 8 },
          ])

          .addRows([
            { name: "Alice", age: 30 },
            { name: "Bob", age: 25 },
          ]);
      })
      .write(outputPath("batch-rows.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("writes positional array rows", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Array" }, (sheet) => {
        sheet
          .headers([
            { key: "a", header: "A", width: 10 },
            { key: "b", header: "B", width: 10 },
          ])

          .addRows([
            [1, 2],
            [3, 4],
          ]);
      })
      .write(outputPath("array-rows.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("colWidth and rowHeight work", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Sizing" }, (sheet) => {
        sheet.setCell("A1", "Wide column").colWidth("A", 40).rowHeight(1, 50);
      })
      .write(outputPath("sizing.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("autoFitColumns works", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "AutoFit" }, (sheet) => {
        sheet.setCell("A1", "Long text here").autoFitColumns();
      })
      .write(outputPath("autofit.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("pageSetup applies", async () => {
    const result = await new WorkbookBuilder()
      .addSheet(
        {
          name: "Page",
          pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true },
        },
        (sheet) => {
          sheet.setCell("A1", "Page setup");
        },
      )
      .write(outputPath("page-setup.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("zoom applies", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "Zoomed", zoom: 150 }, (sheet) => {
        sheet.setCell("A1", "Zoom 150%");
      })
      .write(outputPath("zoom.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("hide gridlines", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "NoGrid", showGridLines: false }, (sheet) => {
        sheet.setCell("A1", "No grid");
      })
      .write(outputPath("no-grid.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("default row height", async () => {
    const result = await new WorkbookBuilder()
      .addSheet({ name: "DefaultH", defaultRowHeight: 30 }, (sheet) => {
        sheet.setCell("A1", "Taller rows");
      })
      .write(outputPath("default-row-height.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  it("toCsvString returns csv as string", async () => {
    const csv = await new WorkbookBuilder()
      .addSheet({ name: "Data" }, (sheet) => {
        sheet
          .headers([
            { key: "name", header: "Name" },
            { key: "age", header: "Age" },
          ])

          .addRows([
            { name: "Alice", age: 30 },
            { name: "Bob", age: 25 },
          ]);
      })
      .toCsvString();

    expect(typeof csv).toBe("string");
    expect(csv).toContain("Alice");
    expect(csv).toContain("30");
  });

  it("writeCsv writes to file", async () => {
    await new WorkbookBuilder()
      .addSheet({ name: "Data" }, (sheet) => {
        sheet
          .headers([{ key: "x", header: "X" }])

          .addRow({ x: 42 });
      })
      .writeCsv(outputPath("test-out.csv"));

    expect(fs.existsSync(outputPath("test-out.csv"))).toBe(true);
  });

  it("load and read back", async () => {
    const written = outputPath("roundtrip.xlsx");
    await new WorkbookBuilder({ author: "Roundtrip" })
      .addSheet({ name: "ReadTest" }, (sheet) => {
        sheet
          .headers([{ key: "val", header: "Value" }])

          .addRow({ val: 99 });
      })
      .write(written);

    const loaded = await WorkbookBuilder.load(written);
    expect(loaded.sheets.length).toBe(1);
  });

  it("template workflow: load → modify → re-save", async () => {
    const original = outputPath("template-original.xlsx");
    const modified = outputPath("template-modified.xlsx");

    await new WorkbookBuilder({ author: "Template" })
      .addSheet({ name: "Sheet1" }, (sheet) => {
        sheet
          .headers([
            { key: "item", header: "Item", width: 20 },
            { key: "qty", header: "Qty", width: 10 },
          ])

          .addRow({ item: "Widget", qty: 10 });
      })
      .write(original);

    const loaded = await WorkbookBuilder.load(original);
    const sheet = loaded.sheets[0];

    expect(sheet).toBeDefined();

    await loaded.write(modified);

    const reloaded = await WorkbookBuilder.load(modified);
    expect(reloaded.sheets.length).toBe(1);
    expect(reloaded.sheets[0]).toBeDefined();
  });

  it("fromFile is an alias for load", async () => {
    const written = outputPath("fromfile.xlsx");
    await new WorkbookBuilder()
      .addSheet({ name: "S" }, (sheet) => sheet.setCell("A1", "hi"))
      .write(written);

    const loaded = await WorkbookBuilder.fromFile(written);
    expect(loaded.sheets.length).toBe(1);
  });

  it("custom author and company metadata", async () => {
    const result = await new WorkbookBuilder({
      author: "Quadro Test",
      company: "Quadro Corp",
    })
      .addSheet({ name: "Meta" }, (sheet) => {
        sheet.setCell("A1", "Meta data");
      })
      .write(outputPath("metadata.xlsx"));

    expect(result.sizeBytes).toBeGreaterThan(0);
  });

  describe("defined names", () => {
    it("defineName registers a named range", async () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Sheet1" }, () => {});
      wb.defineName("MyRange", "A1:B10", "Sheet1");
      const names = wb.getDefinedNames();
      expect(names).toHaveLength(1);
      expect(names[0].name).toBe("MyRange");
    });

    it("defineNameRC registers by numeric coordinates", async () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Sheet1" }, () => {});
      wb.defineNameRC("Data", 1, 1, 5, 10, "Sheet1");
      const names = wb.getDefinedNames();
      expect(names).toHaveLength(1);
      expect(names[0].ranges[0]).toContain("$A$1:$E$10");
    });

    it("defineFormula registers a formula-based name", async () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Sheet1" }, () => {});
      wb.defineFormula("MyLambda", "LAMBDA(x,y,x+y)");
      const names = wb.getDefinedNames();
      expect(names).toHaveLength(1);
      expect(names[0].name).toBe("MyLambda");
    });

    it("removeDefinedName clears ranges", async () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Sheet1" }, () => {});
      wb.defineName("ToRemove", "$A$1:$B$10", "Sheet1");
      expect(wb.getDefinedNames()).toHaveLength(1);
      expect(wb.getDefinedNames()[0].ranges).not.toHaveLength(0);
      wb.removeDefinedName("ToRemove", "Sheet1!$A$1:$B$10");
      expect(wb.getDefinedNames()[0].ranges).toHaveLength(0);
    });

    it("methods are chainable", async () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Sheet1" }, () => {});
      const result = wb
        .defineName("A", "A1", "Sheet1")
        .defineName("B", "B1", "Sheet1")
        .defineFormula("C", "42");
      expect(wb.getDefinedNames()).toHaveLength(3);
      expect(result).toBe(wb);
    });

    it("persists named ranges through write and load", async () => {
      const file = outputPath("named-ranges.xlsx");
      await new WorkbookBuilder()
        .addSheet({ name: "Data" }, (sheet) => {
          sheet.setCell("A1", 10);
          sheet.setCell("B1", 20);
        })
        .defineName("ValueA", "A1", "Data")
        .defineName("ValueB", "B1", "Data")
        .write(file);

      const loaded = await WorkbookBuilder.load(file);
      const names = loaded.workbook.definedNames.getAllEntries();
      expect(names).toHaveLength(2);
      expect(names.find((n) => n.name === "ValueA")).toBeDefined();
      expect(names.find((n) => n.name === "ValueB")).toBeDefined();
    });
  });

  // ── v0.5: Sheet State ────────────────────────────────────────────────────────

  describe("sheet state", () => {
    it("sets sheet state to hidden", async () => {
      const result = await new WorkbookBuilder()
        .addSheet({ name: "Visible" }, () => {})
        .addSheet({ name: "Hidden", state: "hidden" }, () => {})
        .write(outputPath("sheet-state.xlsx"));
      expect(result.sizeBytes).toBeGreaterThan(0);
    });

    it("sets sheet state to veryHidden", async () => {
      const result = await new WorkbookBuilder()
        .addSheet({ name: "VeryHidden", state: "veryHidden" }, () => {})
        .write(outputPath("sheet-veryhidden.xlsx"));
      expect(result.sizeBytes).toBeGreaterThan(0);
    });
  });

  // ── v0.5: Print Titles ───────────────────────────────────────────────────────

  describe("print titles", () => {
    it("applies print titles row", async () => {
      const result = await new WorkbookBuilder()
        .addSheet(
          {
            name: "Print",
            pageSetup: { printTitlesRow: "1:3", printTitlesColumn: "A:B" },
          },
          () => {},
        )
        .write(outputPath("print-titles.xlsx"));
      expect(result.sizeBytes).toBeGreaterThan(0);
    });
  });

  // ── v0.5: Remove Sheet ───────────────────────────────────────────────────────

  describe("removeSheet", () => {
    it("removes a sheet by name", () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Keep" }, () => {});
      wb.addSheet({ name: "Remove" }, () => {});
      expect(wb.sheets).toHaveLength(2);
      wb.removeSheet("Remove");
      expect(wb.sheets).toHaveLength(1);
      expect(wb.sheets[0].name).toBe("Keep");
    });

    it("throws for missing sheet", () => {
      const wb = new WorkbookBuilder();
      expect(() => wb.removeSheet("Missing")).toThrow("No sheet named");
    });

    it("returns this for chaining", () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "S" }, () => {});
      const result = wb.removeSheet("S");
      expect(result).toBe(wb);
    });
  });

  // ── v0.5: Duplicate Sheet ────────────────────────────────────────────────────

  describe("duplicateSheet", () => {
    it("duplicates an existing sheet", () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Original" }, (sheet) => {
        sheet.setCell("A1", 42);
      });
      wb.duplicateSheet("Original", "Copy");
      expect(wb.sheets).toHaveLength(2);
      expect(wb.sheets[1].name).toBe("Copy");
    });

    it("auto-generates name when not provided", () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "Data" }, () => {});
      wb.duplicateSheet("Data");
      expect(wb.sheets).toHaveLength(2);
      expect(wb.sheets[1].name).toMatch(/^Data /);
    });

    it("throws for missing sheet", () => {
      const wb = new WorkbookBuilder();
      expect(() => wb.duplicateSheet("Missing")).toThrow("No sheet named");
    });

    it("returns this for chaining", () => {
      const wb = new WorkbookBuilder();
      wb.addSheet({ name: "S" }, () => {});
      const result = wb.duplicateSheet("S");
      expect(result).toBe(wb);
    });
  });

  // ── v0.5: External Links ─────────────────────────────────────────────────────

  describe("external links", () => {
    it("addExternalLink registers a link", () => {
      const wb = new WorkbookBuilder();
      const link = wb.addExternalLink({
        target: "external_data.xlsx",
        sheetNames: ["Sheet1"],
      });
      expect(link.target).toBe("external_data.xlsx");
    });

    it("getExternalLinks returns registered links", () => {
      const wb = new WorkbookBuilder();
      wb.addExternalLink({ target: "data.xlsx", sheetNames: ["Data"] });
      const links = wb.getExternalLinks();
      expect(links).toHaveLength(1);
      expect(links[0].target).toBe("data.xlsx");
    });

    it("adds link with cached values", () => {
      const wb = new WorkbookBuilder();
      const link = wb.addExternalLink({
        target: "pricing.xlsx",
        sheetNames: ["Sheet1"],
        cachedValues: { Sheet1: { A1: 100, B1: "hello" } },
      });
      expect(link.cachedValues?.Sheet1?.A1).toBe(100);
    });
  });

  // ── v0.5: Custom Workbook Properties ─────────────────────────────────────────

  describe("custom workbook properties", () => {
    it("sets title, subject, keywords, category, manager, description, language", async () => {
      const result = await new WorkbookBuilder({
        title: "Test Report",
        subject: "Monthly Summary",
        keywords: "report, monthly, summary",
        category: "Finance",
        manager: "Jane Doe",
        description: "A test workbook",
        language: "en-US",
      })
        .addSheet({ name: "Props" }, (sheet) => {
          sheet.setCell("A1", "Metadata test");
        })
        .write(outputPath("custom-props.xlsx"));
      expect(result.sizeBytes).toBeGreaterThan(0);
    });
  });

  // ── v0.8: Sheet Name Validation ─────────────────────────────────────────────

  describe("sheet name validation", () => {
    it("rejects name longer than 31 chars", () => {
      expect(() => new WorkbookBuilder().addSheet({ name: "A".repeat(32) }, () => {})).toThrow(
        "exceeds 31",
      );
    });
    it("rejects name with invalid characters", () => {
      expect(() => new WorkbookBuilder().addSheet({ name: "Sheet[1]" }, () => {})).toThrow(
        "invalid characters",
      );
    });
    it("rejects duplicate sheet name", () => {
      expect(() =>
        new WorkbookBuilder()
          .addSheet({ name: "Same" }, () => {})
          .addSheet({ name: "Same" }, () => {}),
      ).toThrow("already exists");
    });
    it("accepts valid short name", () => {
      const wb = new WorkbookBuilder().addSheet({ name: "Valid" }, () => {});
      expect(wb.getSheet("Valid")).toBeDefined();
    });
  });
});
