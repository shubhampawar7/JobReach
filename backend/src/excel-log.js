const fs = require("fs");
const XLSX = require("xlsx");

const { writeFileAtomic } = require("./utils");

const SHEET_NAME = "sent";
const HEADER = ["email", "name", "subject", "error"];

function buildEmptyWorkbook() {
  const ws = XLSX.utils.aoa_to_sheet([HEADER]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
  return wb;
}

function loadWorkbook(filePath) {
  if (!fs.existsSync(filePath)) return buildEmptyWorkbook();
  try {
    return XLSX.readFile(filePath, { cellDates: true });
  } catch {
    return buildEmptyWorkbook();
  }
}

function getOrCreateSheet(wb) {
  if (wb.SheetNames.includes(SHEET_NAME) && wb.Sheets[SHEET_NAME]) {
    return { sheetName: SHEET_NAME, ws: wb.Sheets[SHEET_NAME] };
  }

  const first = wb.SheetNames[0];
  if (first && wb.Sheets[first]) return { sheetName: first, ws: wb.Sheets[first] };

  const ws = XLSX.utils.aoa_to_sheet([HEADER]);
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
  return { sheetName: SHEET_NAME, ws };
}

function normalizeAoaToHeader(aoa) {
  const rows = Array.isArray(aoa) ? aoa : [];
  if (!rows.length) return [HEADER];

  const rawHeader = Array.isArray(rows[0]) ? rows[0] : [];
  const keys = rawHeader.map((v) => String(v ?? "").trim());

  const idxByKey = new Map();
  keys.forEach((k, i) => {
    if (k && !idxByKey.has(k)) idxByKey.set(k, i);
  });

  // Support legacy column names from older versions.
  function pickIdx(headerKey) {
    if (idxByKey.has(headerKey)) return idxByKey.get(headerKey);
    if (headerKey === "email" && idxByKey.has("toEmail")) return idxByKey.get("toEmail");
    if (headerKey === "name" && idxByKey.has("toName")) return idxByKey.get("toName");
    return null;
  }

  const out = [HEADER];
  for (let i = 1; i < rows.length; i++) {
    const r = Array.isArray(rows[i]) ? rows[i] : [];
    const mapped = HEADER.map((k) => {
      const idx = pickIdx(k);
      return idx === null ? "" : (r[idx] ?? "");
    });
    if (mapped.every((v) => String(v ?? "").trim() === "")) continue;
    out.push(mapped);
  }
  return out;
}

function writeWorkbookAtomic(filePath, wb) {
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  writeFileAtomic(filePath, buf);
}

function normalizeWorkbookSheet(wb) {
  const { sheetName, ws } = getOrCreateSheet(wb);
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const normalized = normalizeAoaToHeader(aoa);
  wb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(normalized);
  return { sheetName, ws: wb.Sheets[sheetName] };
}

function appendSentRow(filePath, row) {
  const wb = loadWorkbook(filePath);
  const { sheetName, ws } = normalizeWorkbookSheet(wb);
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  const outRow = HEADER.map((k) => (row && row[k] !== undefined ? row[k] : ""));
  aoa.push(outRow);
  wb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(aoa);
  writeWorkbookAtomic(filePath, wb);
}

function getSentWorkbookBuffer(filePath) {
  const wb = loadWorkbook(filePath);
  normalizeWorkbookSheet(wb);
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

module.exports = {
  SHEET_NAME,
  HEADER,
  appendSentRow,
  getSentWorkbookBuffer,
};


