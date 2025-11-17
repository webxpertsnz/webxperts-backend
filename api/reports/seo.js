This is the last script you have me ? // api/reports/seo.js
//
// Reads an uploaded SEO Excel file and generates a PDF summary.
// Uses dynamic imports for exceljs/formidable so the function
// doesn't crash at load time.

import PDFDocument from "pdfkit";

// ---- dynamic imports ----
async function getFormidable() {
  const mod = await import("formidable");
  return mod.default || mod; // handle both ESM/CJS
}

async function getExcelJS() {
  const mod = await import("exceljs");
  return mod.default || mod;
}

// ---- helpers ----
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

function parseHeaderDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    // Excel serial date
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const dt = new Date(excelEpoch.getTime() + value * 86400000);
    return isNaN(dt) ? null : dt;
  }
  const text = typeof value === "string" ? value : value.text || "";
  const dt = new Date(text);
  return isNaN(dt) ? null : dt;
}

// ---- Excel parsing ----
async function parseSeoWorkbook(filePath) {
  const ExcelJS = await getExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Prefer "Ranking" sheet, else first sheet
  let rankingSheet = workbook.getWorksheet("Ranking");
  if (!rankingSheet) rankingSheet = workbook.worksheets[0];
  if (!rankingSheet) {
    throw new Error("No worksheets found in uploaded file.");
  }

  // Domain/location (optional) in first 10 rows
  let domain = "";
  let location = "";
  for (let r = 1; r <= Math.min(10, rankingSheet.rowCount); r++) {
    const a = (rankingSheet.getCell(r, 1).value || "").toString().trim().toLowerCase();
    const b = rankingSheet.getCell(r, 2).value;
    if (a === "domain" || a === "website") domain = b ? String(b) : "";
    if (a === "location") location = b ? String(b) : "";
  }

  // Find header row & keyword column: any cell with "keyword"
  let headerRowNumber = null;
  let keywordCol = null;

  outer: for (let r = 1; r <= rankingSheet.rowCount; r++) {
    const row = rankingSheet.getRow(r);
    for (let c = 1; c <= row.cellCount; c++) {
      const v = row.getCell(c).value;
      const text = (v || "").toString().trim().toLowerCase();
      if (text === "keyword") {
        headerRowNumber = r;
        keywordCol = c;
        break outer;
      }
    }
  }

  if (!headerRowNumber || !keywordCol) {
    throw new Error(
      "Could not find a header cell with text 'Keyword' in any column. " +
        "Check the ranking sheet header row."
    );
  }

  const headerRow = rankingSheet.getRow(headerRowNumber);

  // URL column & date columns from header row
  let urlCol = null;
  const dateCols = [];

  headerRow.eachCell((cell, col) => {
    const raw = cell.value;
    const label = (raw || "").toString().trim().toLowerCase();

    if (["url", "landing page", "page", "target url"].includes(label)) {
      urlCol = col;
      return;
    }

    const dt = parseHeaderDate(raw);
    if (dt) {
      dateCols.push({ col, date: dt });
    }
  });

  if (!dateCols.length) {
    throw new Error(
      "No date-like headers found in the header row. " +
        "Make sure your ranking sheet has dates as column headings (e.g. 2025-11-10)."
    );
  }

  dateCols.sort((a, b) => a.date - b.date);
  const latest = dateCols[dateCols.length - 1];
  const previous = dateCols.length >= 2 ? dateCols[dateCols.length - 2] : null;

  // Collect keyword rows
  const keywords = [];
  for (let r = headerRowNumber + 1; r <= rankingSheet.rowCount; r++) {
    const row = rankingSheet.getRow(r);
    const kwVal = row.getCell(keywordCol).value;
    if (!kwVal) continue;
    const kw = String(kwVal).trim();
    if (!kw) continue;

    const urlVal = urlCol ? row.getCell(urlCol).value : "";
    const url = urlVal ? String(urlVal) : "";

    const curRaw = row.getCell(latest.col).value;
    const prevRaw = previous ? row.getCell(previous.col).value : null;

    const cur =
      typeof curRaw === "number" ? curRaw : curRaw ? Number(curRaw) : null;
    const prev =
      typeof prevRaw === "number" ? prevRaw : prevRaw ? Number(prevRaw) : null;

    keywords.push({
      keyword: kw,
      url,
      current: cur && cur > 0 ? cur : null,
      previous: prev && prev > 0 ? prev : null
    });
  }

  if (!keywords.length) {
    throw new Error(
      "No keyword rows found under the header. " +
        "Check that your ranking sheet has data below the 'Keyword' row."
    );
  }

  const tracked = keywords.length;
  const withCurrent = keywords.filter((k) => k.current !== null);
  const withPrev = keywords.filter((k) => k.previous !== null);

  const avgCurrent =
    withCurrent.reduce((sum, k) => sum + k.current, 0) /
    (withCurrent.length || 1);
  const avgPrev =
    withPrev.reduce((sum, k) => sum + k.previous, 0) /
    (withPrev.length || 1);

  const top10 = withCurrent.filter((k) => k.current <= 10).length;
  const top10Prev = withPrev.filter((k) => k.previous <= 10).length;

  const movers = keywords
    .map((k) => ({
      ...k,
      change:
        k.current !== null && k.previous !== null
          ? k.previous - k.current
          : null
    }))
    .filter((k) => k.change !== null);

  const topWinners = movers
    .filter((k) => k.change > 0)
    .sort((a, b) => b.change - a.change)
    .slice(0, 10);

  const topLosers = movers
    .filter((k) => k.change < 0)
    .sort((a, b) => a.change - b.change)
    .slice(0, 10);

  return {
    domain,
    location,
    latest,
    previous,
    tracked,
    avgCurrent,
    avgPrev,
    top10,
    top10Prev,
    keywords,
    topWinners,
    topLosers
  };
}

// ---- PDF generation ----
function buildSeoPdf(res, summary) {
  const {
    domain,
    location,
    latest,
    previous,
    tracked,
    avgCurrent,
    avgPrev,
    top10,
    top10Prev,
    keywords,
    topWinners,
    topLosers
  } = summary;

  const doc = new PDFDocument({ margin: 40 });

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="SEO-Report-${domain || "site"}.pdf"`
  );

  doc.pipe(res);

  const fmtDate = (obj) =>
    obj && obj.date ? obj.date.toISOString().slice(0, 10) : "-";

  const top10Pct = tracked > 0 ? Math.round((top10 / tracked) * 100) : 0;
  const top10PrevPct =
    tracked > 0 ? Math.round((top10Prev / tracked) * 100) : 0;

  // Cover
  doc.fontSize(18).text("SEO Weekly Report");
  doc.moveDown(0.5);
  doc.fontSize(11).fillColor("#555");
  if (domain) doc.text(`Domain: ${domain}`);
  if (location) doc.text(`Location: ${location}`);
  doc.text(`Week ending: ${fmtDate(latest)}`);
  if (previous) doc.text(`Compared with: ${fmtDate(previous)}`);
  doc.moveDown();

  doc.fontSize(12).fillColor("#000");
  doc.text(`Tracked keywords: ${tracked}`);
  doc.text(
    `Average position: ${avgCurrent.toFixed(1)} (prev ${avgPrev.toFixed(1)})`
  );
  doc.text(
    `Top 10 visibility: ${top10Pct}% (prev ${top10PrevPct}%)`
  );
  doc.moveDown();

  doc.fontSize(11).fillColor("#555");
  doc.text(
    `This week we are tracking ${tracked} keywords. ` +
      `Average ranking changed from ${avgPrev.toFixed(
        1
      )} to ${avgCurrent.toFixed(
        1
      )}, with ${top10Pct}% of keywords in the top 10.`
  );

  // Winners / losers
  doc.addPage();
  doc.fontSize(14).fillColor("#000").text("Top Winners", { underline: true });
  doc.moveDown(0.5);
  doc.fontSize(10);
  if (!topWinners.length) {
    doc.text("No improving keywords this period.");
  } else {
    topWinners.forEach((k) => {
      doc.text(
        `${k.keyword} — ${k.previous ?? "-"} → ${k.current ?? "-"} (↑ ${
          k.change
        })`
      );
    });
  }

  doc.moveDown();
  doc.fontSize(14).text("Top Losers", { underline: true });
  doc.moveDown(0.5);
  doc.fontSize(10);
  if (!topLosers.length) {
    doc.text("No dropping keywords this period.");
  } else {
    topLosers.forEach((k) => {
      doc.text(
        `${k.keyword} — ${k.previous ?? "-"} → ${k.current ?? "-"} (${
          k.change
        })`
      );
    });
  }

  // Full table
  doc.addPage();
  doc.fontSize(14).fillColor("#000").text("Keyword detail", {
    underline: true
  });
  doc.moveDown(0.5);
  doc.fontSize(9);
  doc.text("Keyword                          Prev   Curr   Change");
  doc.text("------------------------------------------------------");

  keywords.forEach((k) => {
    const prevStr = k.previous ? String(k.previous).padStart(4) : "   -";
    const currStr = k.current ? String(k.current).padStart(4) : "   -";
    let changeStr = "   -";
    if (k.current !== null && k.previous !== null) {
      const diff = k.previous - k.current;
      if (diff > 0) changeStr = ` ↑${String(diff).padStart(2)}`;
      else if (diff < 0)
        changeStr = ` ↓${String(Math.abs(diff)).padStart(2)}`;
      else changeStr = "  0";
    }
    const kw =
      k.keyword.length > 30
        ? k.keyword.slice(0, 27) + "..."
        : k.keyword;
    doc.text(`${kw.padEnd(30)} ${prevStr}  ${currStr}  ${changeStr}`);
  });

  doc.end();
}

// ---- main handler ----
export default async function handler(req, res) {
  try {
    if (req.method === "GET") {
      // health check
      return res
        .status(200)
        .json({ ok: true, message: "SEO reports API is alive" });
    }

    if (req.method !== "POST") {
      res.setHeader("Allow", ["GET", "POST"]);
      return res.status(405).json({ error: "Method not allowed" });
    }

    const { files } = await parseForm(req);
    const file = files.seo_file;
    if (!file) {
      return res.status(400).json({ error: "Missing seo_file upload" });
    }

    const filePath = Array.isArray(file) ? file[0].filepath : file.filepath;
    if (!filePath) {
      return res
        .status(400)
        .json({ error: "Could not access uploaded file path" });
    }

    const summary = await parseSeoWorkbook(filePath);
    buildSeoPdf(res, summary);
  } catch (err) {
    console.error("SEO report error:", err);
    if (!res.headersSent) {
      const msg = err && err.message ? err.message : String(err);
      return res
        .status(500)
        .json({ error: "Failed to generate SEO report: " + msg });
    }
  }
}
