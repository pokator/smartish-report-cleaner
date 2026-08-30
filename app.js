/**
 * app.js — Cross-system product reconciliation logic.
 *
 * This is a JS port of build_report.py. Same column-mapping config, same
 * SKU-normalization rule, same flagging behavior — kept in one file with
 * no framework dependency so it runs identically in the browser (via
 * <script src="app.js">) and in Node (for the test harness in test/).
 *
 * If you change a mapping or a rule here, make the same change in
 * build_report.py, or the two will silently drift apart.
 */

// ============================== CONFIG ==============================

const ORDERS_COLUMNS = {
  order_id: ["orderid", "order"],
  sku_raw: ["sku", "productsku", "itemsku"],
  product: ["product", "productname", "itemname", "description"],
  qty: ["qty", "quantity", "units", "unitssold"],
  revenue: ["revenue", "sales", "amount", "total"],
};

const WAREHOUSE_COLUMNS = {
  sku_raw: ["itemcode", "sku", "productsku"],
  on_hand: ["onhand", "qtyonhand", "stock", "quantity"],
  location: ["location", "warehouse", "bin"],
};

const RETURNS_COLUMNS = {
  rma: ["rma", "rmaid", "returnid"],
  sku_raw: ["sku", "productsku", "itemsku"],
  reason: ["reason", "returnreason"],
  units: ["units", "qty", "quantity"],
};

const SKU_DISPLAY_PREFIX = "SMR";

// ========================== NORMALIZATION ============================

function normalizeColName(name) {
  return String(name == null ? "" : name).trim().toLowerCase().replace(/[\s_-]+/g, "");
}

/** Fuzzy-match actual CSV headers to canonical field names. Missing fields
 *  are reported in `missing` rather than thrown. */
function buildColumnMap(actualHeaders, mapping) {
  const normalizedActual = {};
  actualHeaders.forEach((h) => {
    normalizedActual[normalizeColName(h)] = h;
  });
  const map = {};
  const missing = [];
  for (const canonical of Object.keys(mapping)) {
    const aliases = [canonical, ...mapping[canonical]];
    let found = null;
    for (const alias of aliases) {
      const key = normalizeColName(alias);
      if (Object.prototype.hasOwnProperty.call(normalizedActual, key)) {
        found = normalizedActual[key];
        break;
      }
    }
    map[canonical] = found; // null if not found
    if (!found) missing.push(canonical);
  }
  return { map, missing };
}

function remapRows(rows, columnMap) {
  return rows.map((row) => {
    const out = {};
    for (const canonical of Object.keys(columnMap)) {
      const actual = columnMap[canonical];
      out[canonical] = actual ? row[actual] : undefined;
    }
    return out;
  });
}

/** Pull the trailing digit run out of a raw SKU/item-code, zero-padded to
 *  4 digits. This is the join key across all three systems' SKU formats. */
function extractSkuKey(raw) {
  if (raw === undefined || raw === null) return null;
  const s = String(raw).trim();
  if (s === "") return null;
  const m = s.match(/(\d+)\s*$/);
  if (!m) return null;
  return m[1].padStart(4, "0");
}

function displaySku(key) {
  return `${SKU_DISPLAY_PREFIX}-${key}`;
}

function toNumber(v) {
  if (v === undefined || v === null) return null;
  const s = String(v).trim();
  if (s === "") return null;
  const n = parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

// ============================= LOADERS ===============================

function loadOrders(rows, flags) {
  const headers = Object.keys(rows[0] || {});
  const { map, missing } = buildColumnMap(headers, ORDERS_COLUMNS);
  if (missing.length) {
    flags.push({
      sheet: "Load",
      issue: `Orders file: could not find column(s) [${missing.join(", ")}] — filled with blanks. Check if the export format changed.`,
      detail: "",
    });
  }
  const mapped = remapRows(rows, map);

  mapped.forEach((r) => {
    r.sku_key = extractSkuKey(r.sku_raw);
    r.qty = toNumber(r.qty);
    r.revenue = toNumber(r.revenue);
  });

  mapped
    .filter((r) => !r.sku_key)
    .forEach((r) => {
      flags.push({
        sheet: "Orders",
        issue: "Missing/unreadable SKU — order excluded from product summary",
        detail: `order_id=${r.order_id ?? ""}, product=${r.product ?? ""}`,
      });
    });

  mapped
    .filter((r) => r.sku_key && r.qty === null)
    .forEach((r) => {
      flags.push({
        sheet: "Orders",
        issue: "Missing quantity — treated as 0 in unit totals",
        detail: `order_id=${r.order_id ?? ""}, sku=${r.sku_raw ?? ""}`,
      });
      r.qty = 0;
    });

  return mapped;
}

function loadWarehouse(rows, flags) {
  const headers = Object.keys(rows[0] || {});
  const { map, missing } = buildColumnMap(headers, WAREHOUSE_COLUMNS);
  if (missing.length) {
    flags.push({
      sheet: "Load",
      issue: `Warehouse file: could not find column(s) [${missing.join(", ")}] — filled with blanks. Check if the export format changed.`,
      detail: "",
    });
  }
  const mapped = remapRows(rows, map);

  // Exact-duplicate rows (same raw sku + on_hand + location, BEFORE numeric
  // conversion) are dropped — mirrors build_report.py's dedupe order.
  const seen = new Set();
  const deduped = [];
  for (const r of mapped) {
    const sig = JSON.stringify([r.sku_raw ?? null, r.on_hand ?? null, r.location ?? null]);
    if (seen.has(sig)) {
      flags.push({
        sheet: "Warehouse",
        issue: "Exact duplicate row — dropped",
        detail: `item_code=${r.sku_raw ?? ""}, on_hand=${r.on_hand ?? ""}`,
      });
      continue;
    }
    seen.add(sig);
    deduped.push(r);
  }

  deduped.forEach((r) => {
    r.sku_key = extractSkuKey(r.sku_raw);
    r.on_hand = toNumber(r.on_hand);
  });

  deduped
    .filter((r) => r.sku_key && r.on_hand === null)
    .forEach((r) => {
      flags.push({
        sheet: "Warehouse",
        issue: "Missing on-hand quantity — shown as 'Unknown', NOT assumed zero",
        detail: `item_code=${r.sku_raw ?? ""}, location=${r.location ?? ""}`,
      });
    });

  return deduped;
}

function loadReturns(rows, flags) {
  const headers = Object.keys(rows[0] || {});
  const { map, missing } = buildColumnMap(headers, RETURNS_COLUMNS);
  if (missing.length) {
    flags.push({
      sheet: "Load",
      issue: `Returns file: could not find column(s) [${missing.join(", ")}] — filled with blanks. Check if the export format changed.`,
      detail: "",
    });
  }
  const mapped = remapRows(rows, map);

  mapped.forEach((r) => {
    r.sku_key = extractSkuKey(r.sku_raw);
    r.units = toNumber(r.units) ?? 0;
  });

  mapped
    .filter((r) => !r.sku_key)
    .forEach((r) => {
      flags.push({
        sheet: "Returns",
        issue: "Missing/unreadable SKU — return excluded from product summary",
        detail: `rma=${r.rma ?? ""}`,
      });
    });

  return mapped;
}

// ============================ RECONCILE ===============================

function buildSummary(orders, warehouse, returns, flags) {
  const ordersOk = orders.filter((r) => r.sku_key);
  const warehouseOk = warehouse.filter((r) => r.sku_key);
  const returnsOk = returns.filter((r) => r.sku_key);

  const ordersAgg = new Map();
  for (const r of ordersOk) {
    const e = ordersAgg.get(r.sku_key) || { product: null, units_sold: 0, revenue: 0 };
    if (!e.product && r.product) e.product = r.product;
    e.units_sold += r.qty || 0;
    e.revenue += r.revenue || 0;
    ordersAgg.set(r.sku_key, e);
  }

  const warehouseAgg = new Map();
  for (const r of warehouseOk) {
    const e = warehouseAgg.get(r.sku_key) || { on_hand_sum: 0, on_hand_known: false, locations: new Set() };
    if (r.on_hand !== null) {
      e.on_hand_sum += r.on_hand;
      e.on_hand_known = true;
    }
    if (r.location) e.locations.add(r.location);
    warehouseAgg.set(r.sku_key, e);
  }

  const returnsAgg = new Map();
  for (const r of returnsOk) {
    const e = returnsAgg.get(r.sku_key) || { units_returned: 0, reasons: new Set() };
    e.units_returned += r.units || 0;
    if (r.reason) e.reasons.add(r.reason);
    returnsAgg.set(r.sku_key, e);
  }

  const allKeys = new Set([...ordersAgg.keys(), ...warehouseAgg.keys(), ...returnsAgg.keys()]);
  const sortedKeys = Array.from(allKeys).sort();

  const summaryRows = [];
  for (const key of sortedKeys) {
    const o = ordersAgg.get(key);
    const w = warehouseAgg.get(key);
    const ret = returnsAgg.get(key);

    if (!o && (w || ret)) {
      flags.push({
        sheet: "Summary",
        issue: "SKU has no matching order record — appears only in warehouse/returns data",
        detail: `sku=${displaySku(key)}`,
      });
    }
    if (!w && o) {
      flags.push({
        sheet: "Summary",
        issue: "SKU sold in orders but absent from warehouse export — stock unknown",
        detail: `sku=${displaySku(key)}`,
      });
    }
    if (ret && !o && !w) {
      flags.push({
        sheet: "Summary",
        issue: "Returned SKU not found in orders or warehouse at all — likely bad SKU or discontinued product",
        detail: `sku=${displaySku(key)}`,
      });
    }

    let onHandDisplay;
    if (!w) onHandDisplay = "No record";
    else if (!w.on_hand_known) onHandDisplay = "Unknown";
    else onHandDisplay = w.on_hand_sum;

    summaryRows.push({
      sku: displaySku(key),
      product: o && o.product ? o.product : "(unknown — not in orders export)",
      units_sold: o ? o.units_sold : 0,
      revenue: o ? Math.round(o.revenue * 100) / 100 : 0,
      on_hand: onHandDisplay,
      location: w ? Array.from(w.locations).sort().join(", ") : "",
      units_returned: ret ? ret.units_returned : 0,
      return_reasons: ret ? Array.from(ret.reasons).sort().join("; ") : "",
    });
  }

  return summaryRows;
}

/** Top-level entry point: three arrays of row-objects (as produced by
 *  Papa.parse(file, {header:true})) in -> {summaryRows, flags} out. */
function reconcile(orderRows, warehouseRows, returnRows) {
  const flags = [];
  const orders = loadOrders(orderRows, flags);
  const warehouse = loadWarehouse(warehouseRows, flags);
  const returns = loadReturns(returnRows, flags);
  const summaryRows = buildSummary(orders, warehouse, returns, flags);
  const today = new Date().toISOString().slice(0, 10);
  flags.forEach((f) => {
    f.run_date = today;
  });
  return { summaryRows, flags };
}

// Export for Node (tests) while staying a plain global in the browser.
if (typeof module !== "undefined" && module.exports) {
  module.exports = { reconcile, extractSkuKey, displaySku, normalizeColName };
}
