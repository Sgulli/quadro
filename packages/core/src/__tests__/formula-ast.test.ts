import { describe, expect, it } from "vitest";
import { Formula } from "../formula-ast.js";
import { WorkbookBuilder } from "../workbook-builder.js";

describe("Formula AST — basic nodes", () => {
  it("creates a ref node", () => {
    expect(Formula.ref("A1").toString()).toBe("A1");
  });

  it("creates a range node", () => {
    expect(Formula.range("A1:A10").toString()).toBe("A1:A10");
  });

  it("creates a number node", () => {
    expect(Formula.num(42).toString()).toBe("42");
  });

  it("creates a string node", () => {
    expect(Formula.str("hello").toString()).toBe('"hello"');
  });

  it("creates a raw node", () => {
    expect(Formula.raw("CUSTOM_FN(1,2)").toString()).toBe("CUSTOM_FN(1,2)");
  });
});

describe("Formula AST — function calls", () => {
  it("builds SUM formula", () => {
    expect(Formula.sum("A1:A10").toString()).toBe("SUM(A1:A10)");
  });

  it("builds AVERAGE formula", () => {
    expect(Formula.average("B2:B20").toString()).toBe("AVERAGE(B2:B20)");
  });

  it("builds COUNT formula", () => {
    expect(Formula.count("C1:C100").toString()).toBe("COUNT(C1:C100)");
  });

  it("builds MAX formula", () => {
    expect(Formula.max("D1:D50").toString()).toBe("MAX(D1:D50)");
  });

  it("builds MIN formula", () => {
    expect(Formula.min("E1:E50").toString()).toBe("MIN(E1:E50)");
  });

  it("builds IF formula", () => {
    expect(Formula.if("A1>10", Formula.str("Yes"), Formula.str("No")).toString()).toBe(
      'IF(A1>10,"Yes","No")',
    );
  });

  it("builds IFERROR formula", () => {
    expect(Formula.iferror(Formula.ref("A1"), 0).toString()).toBe("IFERROR(A1,0)");
  });

  it("builds VLOOKUP formula", () => {
    expect(Formula.vlookup("A1", "B1:D100", 2).toString()).toBe("VLOOKUP(A1,B1:D100,2,FALSE)");
  });

  it("builds INDEX formula", () => {
    expect(Formula.index("A1:D10", 3, 2).toString()).toBe("INDEX(A1:D10,3,2)");
  });

  it("builds MATCH formula", () => {
    expect(Formula.match("A1", "B1:B100").toString()).toBe("MATCH(A1,B1:B100,0)");
  });

  it("builds XLOOKUP formula", () => {
    expect(Formula.xlookup("A1", "B1:B10", "C1:C10", "N/A").toString()).toBe(
      'XLOOKUP(A1,B1:B10,C1:C10,"N/A")',
    );
  });

  it("builds ROUND formula", () => {
    expect(Formula.round("A1", 2).toString()).toBe("ROUND(A1,2)");
  });

  it("builds SUMIF formula", () => {
    expect(Formula.sumif("A1:A10", ">5", "B1:B10").toString()).toBe('SUMIF(A1:A10,">5",B1:B10)');
  });

  it("builds COUNTIF formula", () => {
    expect(Formula.countif("A1:A10", ">0").toString()).toBe('COUNTIF(A1:A10,">0")');
  });

  it("builds CONCAT formula", () => {
    expect(Formula.concat("Hello", " ", "World").toString()).toBe('CONCAT("Hello"," ","World")');
  });
});

describe("Formula AST — binary operations", () => {
  it("adds two values", () => {
    expect(Formula.add("A1", "B1").toString()).toBe("A1+B1");
  });

  it("adds multiple values", () => {
    expect(Formula.add("A1", "B1", "C1").toString()).toBe("A1+B1+C1");
  });

  it("subtracts two values", () => {
    expect(Formula.subtract("A1", "B1").toString()).toBe("A1-B1");
  });

  it("multiplies two values", () => {
    expect(Formula.multiply("A1", "B1").toString()).toBe("A1*B1");
  });

  it("divides two values", () => {
    expect(Formula.divide("A1", "B1").toString()).toBe("A1/B1");
  });

  it("computes percentage change", () => {
    expect(Formula.pctChange("A2", "A1").toString()).toBe("(A2-A1)/A1");
  });
});

describe("Formula AST — composition", () => {
  it("composes SUM + SUM via add", () => {
    const total = Formula.add(Formula.sum("A1:A10"), Formula.sum("B1:B10"));
    expect(total.toString()).toBe("SUM(A1:A10)+SUM(B1:B10)");
  });

  it("chains Expr methods", () => {
    const result = Formula.sum("A1:A10").add("B1").mul(2);
    expect(result.toString()).toBe("SUM(A1:A10)+B1*2");
  });

  it("composes nested IF with SUM", () => {
    const f = Formula.if(Formula.raw("SUM(A1:A10)>100"), Formula.sum("A1:A10"), 0);
    expect(f.toString()).toBe("IF(SUM(A1:A10)>100,SUM(A1:A10),0)");
  });

  it("composes IFERROR around VLOOKUP", () => {
    const f = Formula.iferror(Formula.vlookup("A1", "B1:C10", 2), Formula.str("Not Found"));
    expect(f.toString()).toBe('IFERROR(VLOOKUP(A1,B1:C10,2,FALSE),"Not Found")');
  });
});

describe("Formula AST — toFormula()", () => {
  it("converts to FormulaValue", () => {
    const fv = Formula.sum("A1:A10").toFormula();
    expect(fv).toEqual({ formula: "SUM(A1:A10)", result: undefined });
  });

  it("converts to FormulaValue with result", () => {
    const fv = Formula.sum("A1:A10").toFormula(100);
    expect(fv).toEqual({ formula: "SUM(A1:A10)", result: 100 });
  });
});

describe("Formula AST — integration with WorkbookBuilder", () => {
  it("uses Formula AST as cell value", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      sheet
        .headers([
          { key: "val", header: "Value", width: 15 },
          { key: "total", header: "Total", width: 15 },
        ])
        .addRows([
          { val: 10, total: 0 },
          { val: 20, total: 0 },
        ])
        .setCell("B4", Formula.sum("B2:B3").toFormula());
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
  });

  it("uses composed formula in cell", async () => {
    const wb = new WorkbookBuilder();
    wb.addSheet({ name: "Test" }, (sheet) => {
      const total = Formula.add(Formula.sum("A1:A5"), Formula.sum("B1:B5"));
      sheet.setCell("C1", total.toFormula());
    });
    const result = await wb.toBuffer();
    expect(result).toBeInstanceOf(Buffer);
  });
});
