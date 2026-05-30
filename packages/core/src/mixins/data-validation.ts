import type { DataValidation, DataValidationWithFormulae } from "@cj-tech-master/excelts";
import { resolveAddr, resolveRange } from "../coords.js";
import type { Addr, CellRange, RangeValidationDef, SheetBuilderExtension } from "../types.js";

declare module "../sheet-builder.js" {
  interface SheetBuilder {
    addDataValidation(addr: Addr, validation: DataValidation): this;
    removeDataValidation(address: string): this;
    addListValidation(
      range: CellRange,
      list: (string | number | Date)[],
      options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
    ): this;
    addRangeValidation(range: CellRange, validation: RangeValidationDef): this;
  }
}

interface DvWs {
  dataValidations: {
    add(address: string, validation: DataValidation): void;
    remove(address: string): void;
  };
}

interface DvMethods {
  worksheet: DvWs;
}

function formatListValue(v: string | number | Date): string | number | Date {
  if (typeof v === "string") return `"${v.replace(/"/g, '""')}"`;
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const d = String(v.getDate()).padStart(2, "0");
    return `"${y}-${m}-${d}"`;
  }
  return v;
}

export function applyDataValidationMixin(proto: SheetBuilderExtension): void {
  proto.addDataValidation = function (this: DvMethods, addr: Addr, validation: DataValidation) {
    this.worksheet.dataValidations.add(resolveAddr(addr), validation);
    return this;
  };

  proto.removeDataValidation = function (this: DvMethods, address: string) {
    this.worksheet.dataValidations.remove(address);
    return this;
  };

  proto.addListValidation = function (
    this: DvMethods,
    range: CellRange,
    list: (string | number | Date)[],
    options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
  ) {
    this.worksheet.dataValidations.add(resolveRange(range), {
      type: "list",
      formulae: list.map(formatListValue),
      ...options,
    });
    return this;
  };

  proto.addRangeValidation = function (
    this: DvMethods,
    range: CellRange,
    validation: RangeValidationDef,
  ) {
    this.worksheet.dataValidations.add(resolveRange(range), validation);
    return this;
  };
}
