// api/reports/seo.js
//
// Vercel/Next.js API route:
//
// - Accepts multipart/form-data with field "seo_file"
// - Reads the SEO Excel workbook with exceljs
// - Extracts ranking data from the "Ranking" sheet
// - Extracts backlinks from backlink sheets
// - Generates a branded PDF with pdfkit

import path from "path";
import fs from "fs";

// ---------- Dynamic imports ----------
async function getFormidable() {
  const mod = await import("formidable");
  return mod.default || mod;
}

async function getExcelJS() {
  const mod = await import("exceljs");
  return mod.default || mod;
}

async function getPdfKit() {
  const mod = await import("pdfkit");
  return mod.default || mod;
}

// ---------- Helpers ----------
async function parseForm(req) {
  const formidable = await getFormidable();
  return new Promise((resolve, reject) => {
    const form = formidable({ multiples: false });
    form.parse(req, (err, fields, files) => {
      if (err) return reject(err);
      resolve({ fields, files });
    });
  });
}

function cellToString(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.text) return String(v.text);
    if (v.result) return String(v.result);
    if (Array.isArray(v.richText)) {
      return v.richText.map((p) => p.text || "").join("");
    }
  }
  return String(v);
}

// treat only sensible numbers as ranks (not timestamps etc)
function normaliseRank(value) {
  if (value === null || value === undefined) return null;

  let n = null;
  if (typeof value === "number") {
    n = value;
  } else if (!isNaN(Number(value))) {
    n = Number(value);
  }

  if (n === null || n <= 0 || n > 1000) return null;
  return n;
}

// interpret header as a date – understands "1 Sept-25" etc
function parseHeaderDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  if (typeof value === "object" && value !== null) {
    if (value.text) return parseHeaderDate(value.text);
    if (value.result) return parseHeaderDate(value.result);
    if (Array.isArray(value.richText)) {
      const txt = value.richText.map((p) => p.text || "").join("");
      return parseHeaderDate(txt);
    }
  }

  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const dt = new Date(excelEpoch.getTime() + value * 86400000);
    return isNaN(dt) ? null : dt;
  }

  let text = String(value).trim();
  if (!text) return null;

  text = text.replace(/\bSept\b/gi, "Sep");
  text = text.replace(
    /(\d{1,2})\s+([A-Za-z]{3,})-(\d{2})\b/,
    (m, d, mon, yy) => `${d} ${mon} 20${yy}`
  );

  const looksLikeDate =
    /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(text) ||
    /^\d{8}$/.test(text) ||
    (/\d/.test(text) && (text.includes("/") || text.includes("-")));

  if (!looksLikeDate) return null;

  const dt = new Date(text);
  return isNaN(dt) ? null : dt;
}

// ---------- Ranking extraction ----------
function parseRankingSheet(workbook) {
  let sheet = workbook.getWorksheet("Ranking");
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("No worksheets found in uploaded file.");

  const maxHeaderRows = Math.min(15, sheet.rowCount);

  const dateByCol = new Map();
  const rowFreq = new Map();

  for (let r = 1; r <= maxHeaderRows; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= row.cellCount; c++) {
      const v = row.getCell(c).value;
      const dt = parseHeaderDate(v);
      if (dt) {
        const existing = dateByCol.get(c);
        if (!existing || existing.date < dt) {
          dateByCol.set(c, { col: c, date: dt, row: r });
        }
        rowFreq.set(r, (rowFreq.get(r) || 0) + 1);
      }
    }
  }

  const dateCandidates = Array.from(dateByCol.values());
  if (!dateCandidates.length) {
    throw new Error(
      "Could not find any date headers in the ranking sheet. " +
        "Make sure the top rows contain date labels such as '1 Sept-25', '10-Nov-25', etc."
    );
  }

  dateCandidates.sort((a, b) => a.date - b.date);
  const latest = dateCandidates[dateCandidates.length - 1];
  const previous =
    dateCandidates.length >= 2
      ? dateCandidates[dateCandidates.length - 2]
      : null;

  let headerRowNumber = latest.row;
  if (rowFreq.size) {
    let bestRow = headerRowNumber;
    let bestCount = 0;
    for (const [r, count] of rowFreq.entries()) {
      if (count > bestCount) {
        bestRow = r;
        bestCount = count;
      }
    }
    headerRowNumber = bestRow;
  }
  const firstDataRow = headerRowNumber + 1;

  // Domain / location (optional)
  let domain = "";
  let location = "";
  const domainRegex = /[a-z0-9.-]+\.[a-z]{2,}/i;

  for (let r = 1; r <= maxHeaderRows; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= row.cellCount; c++) {
      const text = cellToString(row.getCell(c).value).trim();
      if (!text) continue;
      if (!domain && domainRegex.test(text)) domain = text;
      if (!location && /new zealand/i.test(text)) location = text;
    }
  }

  const KEYWORD_COL = 1;
  const keywords = [];

  for (let r = firstDataRow; r <= sheet.rowCount; r++) {
    const row = sheet.getRow
