import { describe, expect, it } from "vitest";
import { cellRef, colLetter, colRange, rangeRef, resolveAddr, resolveRange } from "../coords.js";

describe("coords", () => {
  describe("colLetter", () => {
    it("returns empty string for 0", () => {
      expect(colLetter(0)).toBe("");
    });

    it("returns A for 1, Z for 26", () => {
      expect(colLetter(1)).toBe("A");
      expect(colLetter(26)).toBe("Z");
    });

    it("wraps to AA, AB, ZZ", () => {
      expect(colLetter(27)).toBe("AA");
      expect(colLetter(28)).toBe("AB");
      expect(colLetter(702)).toBe("ZZ");
    });

    it("handles three-letter columns", () => {
      expect(colLetter(703)).toBe("AAA");
    });
  });

  describe("cellRef", () => {
    it("builds A1 references", () => {
      expect(cellRef(1, 1)).toBe("A1");
      expect(cellRef(2, 3)).toBe("B3");
      expect(cellRef(27, 10)).toBe("AA10");
    });
  });

  describe("colRange", () => {
    it("builds a single-column range", () => {
      expect(colRange(2, 1, 10)).toBe("B1:B10");
    });
  });

  describe("rangeRef", () => {
    it("builds a rectangular range", () => {
      expect(rangeRef(1, 1, 3, 5)).toBe("A1:C5");
    });
  });

  describe("resolveAddr", () => {
    it("returns string addresses unchanged", () => {
      expect(resolveAddr("D5")).toBe("D5");
    });

    it("converts tuple addresses to A1", () => {
      expect(resolveAddr([4, 5])).toBe("D5");
    });
  });

  describe("resolveRange", () => {
    it("returns string ranges unchanged", () => {
      expect(resolveRange("A1:B2")).toBe("A1:B2");
    });

    it("converts tuple ranges to A1 format", () => {
      expect(resolveRange([1, 1, 3, 5])).toBe("A1:C5");
    });
  });
});
