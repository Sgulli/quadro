import { describe, expect, it } from "vitest";
import { handler } from "../debug.js";

describe("debug handler", () => {
  it("rejects unsupported file types", () => {
    expect(() => handler("output/demo.txt")).toThrow("Unsupported file type");
    expect(() => handler("output/demo.csv")).toThrow("Unsupported file type");
    expect(() => handler("output/demo")).toThrow("Unsupported file type");
  });

  it("rejects non-existent file", () => {
    expect(() => handler("/nonexistent/file.xlsx")).toThrow();
  });
});
