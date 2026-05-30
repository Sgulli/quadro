import { describe, expect, it } from "vitest";
import { col, Range, row } from "../../coords/range.js";

describe("col()", () => {
  it("converts number to column letter", () => {
    expect(col(1)).toBe("A");
    expect(col(2)).toBe("B");
    expect(col(26)).toBe("Z");
    expect(col(27)).toBe("AA");
  });

  it("passes through string column letters", () => {
    expect(col("A")).toBe("A");
    expect(col("AZ")).toBe("AZ");
  });
});

describe("row()", () => {
  it("returns the row number unchanged", () => {
    expect(row(1)).toBe(1);
    expect(row(100)).toBe(100);
  });
});

describe("Range.cell()", () => {
  it("builds A1 reference from number + row", () => {
    expect(Range.cell(1, 1)).toBe("A1");
    expect(Range.cell(3, 10)).toBe("C10");
  });

  it("builds A1 reference from letter + row", () => {
    expect(Range.cell("A", 1)).toBe("A1");
    expect(Range.cell("AZ", 99)).toBe("AZ99");
  });
});

describe("Range.column()", () => {
  it("builds a column range from number", () => {
    expect(Range.column(2, 2, 10)).toBe("B2:B10");
  });

  it("builds a column range from letter", () => {
    expect(Range.column("B", 2, 10)).toBe("B2:B10");
    expect(Range.column("D", 1, 50)).toBe("D1:D50");
  });
});

describe("Range.rect()", () => {
  it("builds a rectangular range from numbers", () => {
    expect(Range.rect(1, 1, 3, 10)).toBe("A1:C10");
  });

  it("builds a rectangular range from letters", () => {
    expect(Range.rect("A", 1, "C", 10)).toBe("A1:C10");
  });

  it("handles mixed number/letter inputs", () => {
    expect(Range.rect(1, 1, "C", 10)).toBe("A1:C10");
    expect(Range.rect("B", 2, 5, 20)).toBe("B2:E20");
  });
});

describe("Range.fromTuple()", () => {
  it("converts [col1, row1, col2, row2] to A1 range", () => {
    expect(Range.fromTuple([1, 1, 3, 10])).toBe("A1:C10");
    expect(Range.fromTuple([2, 2, 2, 10])).toBe("B2:B10");
  });
});

describe("Range.fullColumn()", () => {
  it("builds a full-column range with default rows", () => {
    expect(Range.fullColumn("A")).toBe("A1:A1048576");
  });

  it("builds a full-column range with custom start row", () => {
    expect(Range.fullColumn("B", 2)).toBe("B2:B1048576");
  });

  it("accepts numeric column", () => {
    expect(Range.fullColumn(3, 1, 100)).toBe("C1:C100");
  });
});

describe("Range.row()", () => {
  it("builds a full-row range with defaults", () => {
    expect(Range.row(1)).toBe("A1:XFD1");
  });

  it("builds a bounded row range", () => {
    expect(Range.row(5, "A", "D")).toBe("A5:D5");
    expect(Range.row(3, 1, 5)).toBe("A3:E3");
  });
});

describe("Range.offset()", () => {
  it("offsets a cell reference by rows and columns", () => {
    expect(Range.offset("A1", 1, 0)).toBe("A2");
    expect(Range.offset("A1", 0, 1)).toBe("B1");
    expect(Range.offset("B5", 3, 2)).toBe("D8");
  });

  it("throws on invalid base reference", () => {
    expect(() => Range.offset("invalid", 1, 0)).toThrow("Invalid cell reference");
  });

  it("throws when offset produces invalid reference", () => {
    expect(() => Range.offset("A1", -1, 0)).toThrow("Offset produces invalid reference");
  });
});

describe("Range.expand()", () => {
  it("expands a single cell into a range", () => {
    expect(Range.expand("A1", 10)).toBe("A1:A10");
    expect(Range.expand("B2", 5, 3)).toBe("B2:E6");
  });

  it("throws on invalid reference", () => {
    expect(() => Range.expand("bad", 5)).toThrow("Invalid cell reference");
  });
});
