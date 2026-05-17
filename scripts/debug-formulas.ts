import fs from "node:fs";
import path from "node:path";
import { execSync } from "node:child_process";
import { XMLParser } from "fast-xml-parser";

const xlsxPath = path.resolve(process.argv[2] ?? "output/demo-report.xlsx");
const tmp = path.resolve("output/_debug_fx");
fs.mkdirSync(tmp, { recursive: true });

// Unzip sheet files, workbook, and sharedStrings
execSync(`unzip -o "${xlsxPath}" "xl/worksheets/sheet*.xml" "xl/workbook.xml" "xl/sharedStrings.xml" -d "${tmp}" 2>/dev/null`);

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  isArray: (name: string) => name === "row" || name === "c" || name === "si",
});

// Parse workbook calcPr
const wb = parser.parse(fs.readFileSync(path.join(tmp, "xl/workbook.xml"), "utf-8")).workbook;
const calc = wb.calcPr ?? {};
console.log(`calcPr: calcId=${calc.calcId ?? "N/A"} fullCalcOnLoad=${calc.fullCalcOnLoad ?? "N/A"}\n`);

// Parse shared strings
const ss: string[] = [];
const ssPath = path.join(tmp, "xl/sharedStrings.xml");
if (fs.existsSync(ssPath)) {
  const items = parser.parse(fs.readFileSync(ssPath, "utf-8"))?.sst?.si ?? [];
  for (const si of items) {
    ss.push(si?.t ?? si?.r?.t ?? "");
  }
}

// Parse each sheet
const sheets = fs.readdirSync(path.join(tmp, "xl/worksheets")).sort();
for (const sf of sheets) {
  const sheetName = path.basename(sf, ".xml");
  console.log(`=== ${sheetName} ===`);
  const rows = parser.parse(fs.readFileSync(path.join(tmp, "xl/worksheets", sf), "utf-8")).worksheet?.sheetData?.row ?? [];
  for (const row of (Array.isArray(rows) ? rows : [rows])) {
    const cells = Array.isArray(row.c) ? row.c : row.c ? [row.c] : [];
    for (const c of cells) {
      const formula = c.f != null
        ? (typeof c.f === "object" ? c.f["#text"] ?? "" : c.f)
        : null;
      if (formula == null) continue;
      const cached = c.v != null
        ? c.t === "s" ? `"${ss[Number(c.v)] ?? "???"}"` : c.v
        : "(none)";
      console.log(`  ${c.r}  ${formula}  [cached=${cached}]`);
    }
  }
  console.log();
}

// Cleanup
fs.rmSync(tmp, { recursive: true, force: true });
