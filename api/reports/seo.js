// api/reports/seo.js
//
// Node/Vercel serverless function:
// - Accepts multipart/form-data with field "seo_file"
// - Reads Excel file with exceljs
// - Generates PDF with pdfkit

// ---------- Dynamic imports so top-level never crashes ----------
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

function parseHeaderDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  if (typeof value === "number") {
    // Excel serial date
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const dt = new Date(excelEpoch.getTime() + value * 86400000);
    return isNaN(dt) ? null : dt;
  }

  const text =
    typeof value === "string"
      ? value
      : (value && value.text) ? value.text : "";

  const dt = new Date(text);
  return isNaN(dt) ? null : dt;
}

// ---------- Excel parsing ----------
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
    const a = (rankingSheet.getCell(r, 1).value || "")
      .toString()
      .trim()
      .toLowerCase();
    const b = rankingSheet.getCell(r, 2).value;
    if (a === "domain" || a === "website") domain = b ? String(b) : "";
    if (a === "location") location = b ? String(b) : "";
  }

  // ---------- Find header row & keyword column ----------
  let headerRowNumber = null;
  let keywordCol = null;

  // Pass 1: look for any header that CONTAINS "keyword"
  outer: for (let r = 1; r <= rankingSheet.rowCount; r++) {
    const row = rankingSheet.getRow(r);
    for (let c = 1; c <= row.cellCount; c++) {
      const v = row.getCell(c).value;
      const text = (v || "").toString().trim().toLowerCase();
      if (!text) continue;
      if (text === "keyword" || text === "keywords" || text.includes("keyword")) {
        headerRowNumber = r;
        keywordCol = c;
        break outer;
      }
    }
  }

  // Pass 2: fallback — pick the first "busy" row as header if still not found
  if (!headerRowNumber || !keywordCol) {
    let bestRow = null;
    let bestCount = 0;

    for (let r = 1; r <= rankingSheet.rowCount; r++) {
      const row = rankingSheet.getRow(r);
      let nonEmpty = 0;
      for (let c = 1; c <= row.cellCount; c++) {
        const v = row.getCell(c).value;
        if (v !== null && v !== undefined && String(v).trim() !== "") {
          nonEmpty++;
        }
      }
      if (nonEmpty >= 3) {
        bestRow = r;
        bestCount = nonEmpty;
        break;
      }
    }

    if (bestRow) {
      headerRowNumber = bestRow;
      const row = rankingSheet.getRow(bestRow);
      for (let c = 1; c <= row.cellCount; c++) {
        const v = row.getCell(c).value;
        if (v !== null && v !== undefined && String(v).trim() !== "") {
          keywordCol = c;
          break;
        }
      }
    }
  }

  if (!headerRowNumber || !keywordCol) {
    throw new Error(
      "Could not identify the header row / keyword column. " +
        "Make sure your ranking sheet has a header row (e.g. with 'Keyword' or 'Keywords')."
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
        "Check that your ranking sheet has data below the header row."
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

// ---------- PDF generation ----------
async function buildSeoPdf(res, summary) {
  const PDFKit = await getPdfKit();
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

  const doc = new PDFKit({ margin: 40 });

  res.statusCode = 200;
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
  doc.text(`Top 10 visibility: ${top10Pct}% (prev ${top10PrevPct}%)`);
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
      k.keyword.length > 30 ? k.keyword.slice(0, 27) + "..." : k.keyword;
    doc.text(`${kw.padEnd(30)} ${prevStr}  ${currStr}  ${changeStr}`);
  });

  doc.end();
}

// ---------- Main handler ----------
export default async function handler(req, res) {
  try {
    if (req.method === "GET") {
      // health check
      res.statusCode = 200;
      return res.json({
        ok: true,
        message: "SEO reports API is alive"
      });
    }

    if (req.method !== "POST") {
      res.setHeader("Allow", ["GET", "POST"]);
      res.statusCode = 405;
      return res.json({ error: "Method not allowed" });
    }

    const { files } = await parseForm(req);
    const file = files.seo_file;

    if (!file) {
      res.statusCode = 400;
      return res.json({ error: "Missing seo_file upload" });
    }

    const filePath = Array.isArray(file) ? file[0].filepath : file.filepath;
    if (!filePath) {
      res.statusCode = 400;
      return res.json({ error: "Could not access uploaded file path" });
    }

    const summary = await parseSeoWorkbook(filePath);

    // Stream PDF out
    await buildSeoPdf(res, summary);
  } catch (err) {
    console.error("SEO report error:", err);
    if (!res.headersSent) {
      const msg = err && err.message ? err.message : String(err);
      res.statusCode = 500;
      return res.json({
        error: "Failed to generate SEO report: " + msg
      });
    }
  }
}
