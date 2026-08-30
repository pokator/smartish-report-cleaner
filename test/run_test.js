/**
 * Sanity check: runs the JS reconciliation logic (app.js) against the
 * sample CSVs and prints the result, so it can be eyeballed against
 * build_report.py's output on the same files. Not a full test framework —
 * just a fast way to catch drift between the two implementations.
 */
const fs = require("fs");
const path = require("path");
const Papa = require("papaparse");
const { reconcile } = require("../app.js");

function parseCsv(filePath) {
  const text = fs.readFileSync(filePath, "utf8");
  const result = Papa.parse(text, { header: true, skipEmptyLines: true });
  return result.data;
}

const dir = path.join(__dirname, "..", "sample_data");
const orders = parseCsv(path.join(dir, "system1_shopify_orders.csv"));
const warehouse = parseCsv(path.join(dir, "system2_warehouse_export.csv"));
const returns = parseCsv(path.join(dir, "system3_returns.csv"));

const { summaryRows, flags } = reconcile(orders, warehouse, returns);

console.log("=== Summary ===");
console.table(summaryRows);
console.log(`\n=== Flags (${flags.length}) ===`);
flags.forEach((f) => console.log(`[${f.sheet}] ${f.issue} -- ${f.detail}`));
