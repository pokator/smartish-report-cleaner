/**
 * ui.js — wires the DOM (drop zones, button, results table) to the
 * reconciliation logic in app.js and to SheetJS for the .xlsx download.
 * No logic about SKUs or column-matching lives here — that's all in app.js.
 */

const slots = {
  orders: { file: null, el: document.getElementById("slot-orders"), input: document.getElementById("file-orders") },
  warehouse: { file: null, el: document.getElementById("slot-warehouse"), input: document.getElementById("file-warehouse") },
  returns: { file: null, el: document.getElementById("slot-returns"), input: document.getElementById("file-returns") },
};

const generateBtn = document.getElementById("generate-btn");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");
const downloadBtn = document.getElementById("download-btn");

let lastSummaryRows = null;
let lastFlags = null;

function setSlotFile(key, file) {
  const slot = slots[key];
  slot.file = file;
  slot.el.classList.toggle("filled", !!file);
  const hint = slot.el.querySelector("[data-hint]");
  const nameEl = slot.el.querySelector("[data-filename]");
  if (file) {
    hint.hidden = true;
    nameEl.hidden = false;
    nameEl.textContent = file.name;
  } else {
    hint.hidden = false;
    nameEl.hidden = true;
  }
  generateBtn.disabled = !(slots.orders.file && slots.warehouse.file && slots.returns.file);
}

// Wire up click-to-browse + drag-and-drop for each slot.
Object.entries(slots).forEach(([key, slot]) => {
  slot.input.addEventListener("change", (e) => {
    if (e.target.files[0]) setSlotFile(key, e.target.files[0]);
  });

  ["dragenter", "dragover"].forEach((evt) =>
    slot.el.addEventListener(evt, (e) => {
      e.preventDefault();
      slot.el.classList.add("drag-over");
    })
  );
  ["dragleave", "drop"].forEach((evt) =>
    slot.el.addEventListener(evt, (e) => {
      e.preventDefault();
      slot.el.classList.remove("drag-over");
    })
  );
  slot.el.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    if (file) setSlotFile(key, file);
  });
});

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (result) => resolve(result.data),
      error: reject,
    });
  });
}

generateBtn.addEventListener("click", async () => {
  statusEl.textContent = "Reading files…";
  statusEl.classList.remove("error");
  generateBtn.disabled = true;

  try {
    const [orderRows, warehouseRows, returnRows] = await Promise.all([
      parseCsvFile(slots.orders.file),
      parseCsvFile(slots.warehouse.file),
      parseCsvFile(slots.returns.file),
    ]);

    statusEl.textContent = "Reconciling…";
    const { summaryRows, flags } = reconcile(orderRows, warehouseRows, returnRows);
    lastSummaryRows = summaryRows;
    lastFlags = flags;

    renderResults(summaryRows, flags);
    statusEl.textContent = `Done — ${summaryRows.length} products, ${flags.length} flagged.`;
  } catch (err) {
    console.error(err);
    statusEl.textContent = `Something went wrong reading those files: ${err.message || err}`;
    statusEl.classList.add("error");
  } finally {
    generateBtn.disabled = false;
  }
});

function renderResults(summaryRows, flags) {
  resultsEl.hidden = false;
  document.getElementById("stat-products").textContent = summaryRows.length;
  document.getElementById("stat-flags").textContent = flags.length;

  const body = document.getElementById("preview-body");
  body.innerHTML = "";
  summaryRows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.sku)}</td>
      <td>${escapeHtml(row.product)}</td>
      <td>${row.units_sold}</td>
      <td>${Number(row.revenue).toFixed(2)}</td>
      <td>${escapeHtml(String(row.on_hand))}</td>
      <td>${escapeHtml(row.location)}</td>
      <td>${row.units_returned}</td>
      <td>${escapeHtml(row.return_reasons)}</td>
    `;
    body.appendChild(tr);
  });

  const flagsList = document.getElementById("flags-list");
  flagsList.innerHTML = "";
  if (flags.length === 0) {
    const li = document.createElement("li");
    li.textContent = "Nothing flagged — all three files reconciled cleanly.";
    flagsList.appendChild(li);
  } else {
    flags.forEach((f) => {
      const li = document.createElement("li");
      li.innerHTML = `<span class="tag">${escapeHtml(f.sheet)}</span>${escapeHtml(f.issue)} — <em>${escapeHtml(f.detail)}</em>`;
      flagsList.appendChild(li);
    });
  }
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

downloadBtn.addEventListener("click", () => {
  if (!lastSummaryRows) return;

  const summaryHeader = ["sku", "product", "units_sold", "revenue", "on_hand", "location", "units_returned", "return_reasons"];
  const summarySheet = XLSX.utils.json_to_sheet(lastSummaryRows, { header: summaryHeader });
  summarySheet["!cols"] = [
    { wch: 12 }, { wch: 30 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 30 },
  ];

  const flagsHeader = ["run_date", "sheet", "issue", "detail"];
  const flagsSheet = XLSX.utils.json_to_sheet(lastFlags, { header: flagsHeader });
  flagsSheet["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 60 }, { wch: 40 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");
  XLSX.utils.book_append_sheet(wb, flagsSheet, "Flags");

  XLSX.writeFile(wb, "weekly_report.xlsx");
});
