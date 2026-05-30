import type {
  AddImageRange,
  CellHyperlinkValueInput,
  CellRichTextValue,
  FormCheckbox as ExcelFormCheckbox,
  FormCheckboxOptions,
  FormControlRange,
  RichText as RichTextRun,
  ThreadedComment,
  WatermarkOptions,
} from "@cj-tech-master/excelts";
import { resolveAddr } from "../coords/coords.js";
import type {
  Addr,
  AddSparklineGroupOptions,
  NoteConfig,
  SheetBuilderExtension,
} from "../types.js";

declare module "../builders/sheet-builder.js" {
  interface SheetBuilder {
    addNote(addr: Addr, text: string): this;
    addNote(addr: Addr, config: NoteConfig): this;
    addThreadedComment(ref: string, comment: ThreadedComment): this;
    setCellHyperlink(addr: Addr, hyperlink: string, text?: string, tooltip?: string): this;
    setCellRichText(addr: Addr, richText: RichTextRun[]): this;
    addImage(imageId: string | number, range: AddImageRange): this;
    addBackgroundImage(imageId: string | number): this;
    addWatermark(options: WatermarkOptions): this;
    removeWatermark(): this;
    addSparklineGroup(options: AddSparklineGroupOptions): this;
    addFormCheckbox(range: FormControlRange, options?: FormCheckboxOptions): this;
    getFormCheckboxes(): ExcelFormCheckbox[];
  }
}

interface MediaWs {
  threadedComments: Array<{ ref: string; comment: ThreadedComment }>;
  addImage(imageId: string | number, range: AddImageRange): void;
  addBackgroundImage(imageId: string | number): void;
  addWatermark(options: WatermarkOptions): void;
  removeWatermark(): void;
  addSparklineGroup(options: AddSparklineGroupOptions): void;
  addFormCheckbox(range: FormControlRange, options?: FormCheckboxOptions): void;
  getFormCheckboxes(): ExcelFormCheckbox[];
  getCell(address: string): {
    value: CellHyperlinkValueInput | CellRichTextValue | null;
    note: string;
  };
}

interface MediaMethods {
  worksheet: MediaWs;
}

export function applyMediaMixin(proto: SheetBuilderExtension): void {
  proto.addNote = function (this: MediaMethods, addr: Addr, textOrConfig: string | NoteConfig) {
    const address = resolveAddr(addr);
    const cell = this.worksheet.getCell(address);
    if (typeof textOrConfig === "string") {
      cell.note = textOrConfig;
    } else {
      const texts = textOrConfig.texts?.map((t) => t.text).join("\n") ?? "";
      cell.note = texts;
    }
    return this;
  };

  proto.addThreadedComment = function (this: MediaMethods, ref: string, comment: ThreadedComment) {
    const entry = { ref, comment: { ...comment, date: comment.date ?? new Date().toISOString() } };
    this.worksheet.threadedComments.push(entry);
    return this;
  };

  proto.setCellHyperlink = function (
    this: MediaMethods,
    addr: Addr,
    hyperlink: string,
    text?: string,
    tooltip?: string,
  ) {
    const cell = this.worksheet.getCell(resolveAddr(addr));
    cell.value = { text: text ?? hyperlink, hyperlink, tooltip } as CellHyperlinkValueInput;
    return this;
  };

  proto.setCellRichText = function (this: MediaMethods, addr: Addr, richText: RichTextRun[]) {
    const cell = this.worksheet.getCell(resolveAddr(addr));
    cell.value = { richText } as CellRichTextValue;
    return this;
  };

  proto.addImage = function (this: MediaMethods, imageId: string | number, range: AddImageRange) {
    this.worksheet.addImage(imageId, range);
    return this;
  };

  proto.addBackgroundImage = function (this: MediaMethods, imageId: string | number) {
    this.worksheet.addBackgroundImage(imageId);
    return this;
  };

  proto.addWatermark = function (this: MediaMethods, options: WatermarkOptions) {
    this.worksheet.addWatermark(options);
    return this;
  };

  proto.removeWatermark = function (this: MediaMethods) {
    this.worksheet.removeWatermark();
    return this;
  };

  proto.addSparklineGroup = function (this: MediaMethods, options: AddSparklineGroupOptions) {
    this.worksheet.addSparklineGroup(options);
    return this;
  };

  proto.addFormCheckbox = function (
    this: MediaMethods,
    range: FormControlRange,
    options?: FormCheckboxOptions,
  ) {
    this.worksheet.addFormCheckbox(range, options);
    return this;
  };

  proto.getFormCheckboxes = function (this: MediaMethods): ExcelFormCheckbox[] {
    return this.worksheet.getFormCheckboxes();
  };
}
