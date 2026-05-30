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

const ROWS = 1_000_000;

describe("benchmark", { timeout: 300_000 }, () => {
  it(`build: writes ${ROWS.toLocaleString()} rows via addRows`, async () => {
    fs.mkdirSync(outputDir, { recursive: true });
    const outputFile = path.join(outputDir, "1m-rows.xlsx");

    const builder = new WorkbookBuilder({ author: "Benchmark" });
    builder.addSheet({ name: "Data" }, (sheet) => {
      sheet.headers([
        { key: "item", header: "Item", width: 16 },
        { key: "qty", header: "Qty", width: 10 },
        { key: "price", header: "Price", width: 12 },
      ]);

      const batch = 10_000;
      for (let i = 0; i < ROWS; i += batch) {
        const chunk: { item: string; qty: number; price: number }[] = [];
        const end = Math.min(i + batch, ROWS);
        for (let j = i; j < end; j++) {
          chunk.push({ item: `Item-${j}`, qty: j % 1000, price: (j % 100) + 0.99 });
        }
        sheet.addRows(chunk);
      }
    });

    const t0 = performance.now();
    await builder.write(outputFile);
    const elapsed = performance.now() - t0;

    const stats = fs.statSync(outputFile);
    const sizeMb = (stats.size / 1024 / 1024).toFixed(1);

    console.log(`\n${stamp()}`);
    console.log(`  Rows:      ${ROWS.toLocaleString()}`);
    console.log(`  Columns:   3`);
    console.log(`  Batch:     10_000`);
    console.log(`  File:      ${outputFile}`);
    console.log(`  Size:      ${sizeMb} MB`);
    console.log(`  Time:      ${fmt(elapsed)}`);
    console.log("───────────────────────────────────────────\n");
  });

  it(`read: loads back the ${ROWS.toLocaleString()}-row workbook`, async () => {
    const outputFile = path.join(outputDir, "1m-rows.xlsx");

    const t0 = performance.now();
    const loaded = await WorkbookBuilder.load(outputFile);
    const elapsed = performance.now() - t0;

    const sheet = loaded.getSheet("Data");
    const ws = sheet?.worksheet;
    const rows = (ws as { rowCount?: number })?.rowCount ?? 0;

    console.log(`\n${stamp()}`);
    console.log(`  File:      ${outputFile}`);
    console.log(`  Rows read: ${rows.toLocaleString()}`);
    console.log(`  Time:      ${fmt(elapsed)}`);
    console.log("───────────────────────────────────────────\n");
  });
});
