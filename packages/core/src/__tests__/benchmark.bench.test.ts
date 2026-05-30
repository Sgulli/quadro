import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, it } from "vitest";
import { WorkbookBuilder } from "../index.js";

const outputDir = path.resolve("output/benchmark");

function stamp(): string {
  const cpus = os.cpus();
  return [
    "── Benchmark ──────────────────────────────",
    `  Date:      ${new Date().toISOString()}`,
    `  Platform:  ${os.platform()} ${os.arch()}`,
    `  CPU:       ${cpus[0].model} (${cpus.length} cores)`,
    `  RAM:       ${Math.round(os.totalmem() / 1024 / 1024 / 1024)} GB`,
    `  Hostname:  ${os.hostname()}`,
    `  Node:      ${process.version}`,
    "───────────────────────────────────────────",
  ].join("\n");
}

function fmt(ms: number): string {
  if (ms >= 60_000) return `${(ms / 60_000).toFixed(2)}m`;
  if (ms >= 1_000) return `${(ms / 1_000).toFixed(2)}s`;
  return `${ms.toFixed(0)}ms`;
}

const ROWS = 500_000;
const BATCH = 10_000;

function generateChunk(start: number, end: number) {
  const chunk: { item: string; qty: number; price: number }[] = [];
  for (let j = start; j < end; j++) {
    chunk.push({ item: `Item-${j}`, qty: j % 1000, price: (j % 100) + 0.99 });
  }
  return chunk;
}

describe("benchmark", { timeout: 300_000 }, () => {
  it(`streaming: writes ${ROWS.toLocaleString()} rows with streaming mode`, async () => {
    fs.mkdirSync(outputDir, { recursive: true });
    const outputFile = path.join(outputDir, "1m-rows-stream.xlsx");

    const builder = new WorkbookBuilder({
      author: "Benchmark",
      useStreaming: true,
      useSharedStrings: true,
    });
    builder.addSheet({ name: "Data" }, (sheet) => {
      sheet.headers([
        { key: "item", header: "Item", width: 16 },
        { key: "qty", header: "Qty", width: 10 },
        { key: "price", header: "Price", width: 12 },
      ]);
    });

    const t0 = performance.now();
    const sheet = builder.getSheet("Data");
    for (let i = 0; i < ROWS; i += BATCH) {
      const end = Math.min(i + BATCH, ROWS);
      sheet?.addRows(generateChunk(i, end));
    }
    await builder.write(outputFile);
    const elapsed = performance.now() - t0;

    const stats = fs.statSync(outputFile);
    const sizeMb = (stats.size / 1024 / 1024).toFixed(1);

    console.log(`\n${stamp()}`);
    console.log(`  Mode:      streaming`);
    console.log(`  Rows:      ${ROWS.toLocaleString()}`);
    console.log(`  Columns:   3`);
    console.log(`  Batch:     ${BATCH.toLocaleString()}`);
    console.log(`  File:      ${outputFile}`);
    console.log(`  Size:      ${sizeMb} MB`);
    console.log(`  Time:      ${fmt(elapsed)}`);
    console.log("───────────────────────────────────────────\n");
  });

  it(`standard: writes ${ROWS.toLocaleString()} rows (normal mode)`, async () => {
    fs.mkdirSync(outputDir, { recursive: true });
    const outputFile = path.join(outputDir, "1m-rows-standard.xlsx");

    const builder = new WorkbookBuilder({
      author: "Benchmark",
      useSharedStrings: true,
    });
    builder.addSheet({ name: "Data" }, (sheet) => {
      sheet.headers([
        { key: "item", header: "Item", width: 16 },
        { key: "qty", header: "Qty", width: 10 },
        { key: "price", header: "Price", width: 12 },
      ]);

      for (let i = 0; i < ROWS; i += BATCH) {
        const end = Math.min(i + BATCH, ROWS);
        sheet.addRows(generateChunk(i, end));
      }
    });

    const t0 = performance.now();
    await builder.write(outputFile);
    const elapsed = performance.now() - t0;

    const stats = fs.statSync(outputFile);
    const sizeMb = (stats.size / 1024 / 1024).toFixed(1);

    console.log(`\n${stamp()}`);
    console.log(`  Mode:      standard`);
    console.log(`  Rows:      ${ROWS.toLocaleString()}`);
    console.log(`  Columns:   3`);
    console.log(`  Batch:     ${BATCH.toLocaleString()}`);
    console.log(`  File:      ${outputFile}`);
    console.log(`  Size:      ${sizeMb} MB`);
    console.log(`  Time:      ${fmt(elapsed)}`);
    console.log("───────────────────────────────────────────\n");
  });

  it(`read: loads back the ${ROWS.toLocaleString()}-row workbook`, async () => {
    const outputFile = path.join(outputDir, "1m-rows-stream.xlsx");

    const t0 = performance.now();
    const loaded = await WorkbookBuilder.load(outputFile);
    const elapsed = performance.now() - t0;

    const sheet = loaded.getSheet("Data");
    const ws = sheet?.worksheet;
    const rows = ws?.rowCount ?? 0;

    console.log(`\n${stamp()}`);
    console.log(`  File:      ${outputFile}`);
    console.log(`  Rows read: ${rows.toLocaleString()}`);
    console.log(`  Time:      ${fmt(elapsed)}`);
    console.log("───────────────────────────────────────────\n");
  });
});
