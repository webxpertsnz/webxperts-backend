// api/reports/seo.js
//
// Vercel/Next.js API route:
//
// - Accepts multipart/form-data with field "seo_file"
// - Reads the SEO Excel workbook with exceljs
// - Extracts ranking data from the "Ranking" sheet
// - Extracts backlinks from the backlink sheets
// - Generates a human-readable PDF with pdfkit

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

// “Is this a *position* or rank?”  (1–100 etc, not a timestamp)
function normaliseRank(value) {
  if (value === null || value === undefined) return null;

  let n = null;
  if (typeof value === "number") {
    n = value;
  } else if (!isNaN(Number(value))) {
    n = Number(value);
  }

  // treat zeros / negatives / huge numbers as “not a rank”
  if (n === null || n <= 0 || n > 1000) return null;
  return n;
}

// Try to interpret header value as a date (for nicer labels)
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
    // Excel serial date
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const dt = new Date(excelEpoch.getTime() + value * 86400000);
    return isNaN(dt) ? null : dt;
  }

  const text = String(value).trim();
  if (!text) return null;

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
  // Prefer a sheet called "Ranking", otherwise first sheet
  let sheet = workbook.getWorksheet("Ranking");
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("No worksheets found in uploaded file.");

  // Domain / location in first 10 rows (optional)
  let domain = "";
  let location = "";
  for (let r = 1; r <= Math.min(10, sheet.rowCount); r++) {
    const a = cellToString(sheet.getCell(r, 1).value).trim().toLowerCase();
    const b = sheet.getCell(r, 2).value;
    if (a === "domain" || a === "website") domain = b ? String(b) : "";
    if (a === "location") location = b ? String(b) : "";
  }

  // --- Find header row & keyword column ---
  let headerRowNumber = null;
  let keywordCol = null;

  // Pass 1: row that actually mentions "keyword"
  outer: for (let r = 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= row.cellCount; c++) {
      const text = cellToString(row.getCell(c).value)
        .trim()
        .toLowerCase();
      if (!text) continue;
      if (text.includes("keyword")) {
        headerRowNumber = r;
        keywordCol = c;
        break outer;
      }
    }
  }

  // Pass 2: fallback — first “busy” row with 3+ non-empty cells
  if (!headerRowNumber || !keywordCol) {
    for (let r = 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      let nonEmpty = 0;
      let firstNonEmptyCol = null;
      for (let c = 1; c <= row.cellCount; c++) {
        const text = cellToString(row.getCell(c).value).trim();
        if (text) {
          nonEmpty++;
          if (!firstNonEmptyCol) firstNonEmptyCol = c;
        }
      }
      if (nonEmpty >= 3) {
        headerRowNumber = r;
        keywordCol = firstNonEmptyCol;
        break;
      }
    }
  }

  if (!headerRowNumber || !keywordCol) {
    throw new Error(
      "Could not identify the header row / keyword column. " +
        "Make sure your ranking sheet has a header row with a Keyword column."
    );
  }

  const headerRow = sheet.getRow(headerRowNumber);
  const columnCount = sheet.columnCount || headerRow.cellCount;

  // --- URL column ---
  let urlCol = null;
  for (let c = 1; c <= columnCount; c++) {
    const label = cellToString(headerRow.getCell(c).value)
      .trim()
      .toLowerCase();
    if (
      ["url", "landing page", "page", "target url", "terget url", "address"].includes(
        label
      )
    ) {
      urlCol = c;
      break;
    }
  }

  // --- Detect ranking columns (latest & previous) ---
  const headerCandidates = [];

  for (let c = keywordCol + 1; c <= columnCount; c++) {
    const raw = headerRow.getCell(c).value;
    const label = cellToString(raw).trim();
    if (!label) continue;

    const lower = label.toLowerCase();
    const date = parseHeaderDate(raw);

    const score =
      (lower.includes("current") ? 4 : 0) +
      (lower.includes("prev") || lower.includes("last") ? 3 : 0) +
      (lower.includes("rank") || lower.includes("position") ? 2 : 0) +
      (date ? 1 : 0);

    headerCandidates.push({
      col: c,
      label,
      lower,
      date,
      score
    });
  }

  // Sort by “how likely this is to be the latest column”
  headerCandidates.sort((a, b) => {
    if (a.score !== b.score) return a.score - b.score;
    if (a.date && b.date) return a.date - b.date;
    return a.col - b.col;
  });

  let latest = null;
  let previous = null;

  if (headerCandidates.length) {
    latest = headerCandidates[headerCandidates.length - 1];
    if (headerCandidates.length >= 2) {
      previous = headerCandidates[headerCandidates.length - 2];
    }
  } else {
    // Fallback: look at numeric cells below the header and use rightmost two
    const numericCols = new Set();
    const maxSampleRows = Math.min(sheet.rowCount, headerRowNumber + 50);
    for (let r = headerRowNumber + 1; r <= maxSampleRows; r++) {
      const row = sheet.getRow(r);
      for (let c = keywordCol + 1; c <= columnCount; c++) {
        const rank = normaliseRank(row.getCell(c).value);
        if (rank !== null) numericCols.add(c);
      }
    }
    const cols = Array.from(numericCols).sort((a, b) => a - b);
    if (cols.length) {
      const lastCol = cols[cols.length - 1];
      const prevCol = cols.length >= 2 ? cols[cols.length - 2] : null;

      latest = {
        col: lastCol,
        label:
          cellToString(headerRow.getCell(lastCol).value).trim() ||
          `Column ${lastCol}`,
        date: null
      };
      if (prevCol) {
        previous = {
          col: prevCol,
          label:
            cellToString(headerRow.getCell(prevCol).value).trim() ||
            `Column ${prevCol}`,
          date: null
        };
      }
    }
  }

  if (!latest) {
    throw new Error(
      "Could not find any ranking/position columns to the right of the Keyword column. " +
        "Please make sure your sheet has rank/position columns with numeric values."
    );
  }

  // --- Collect keyword rows ---
  const keywords = [];

  for (let r = headerRowNumber + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const kwVal = row.getCell(keywordCol).value;
    if (!kwVal) continue;

    const kw = cellToString(kwVal).trim();
    if (!kw) continue;

    const urlVal = urlCol ? row.getCell(urlCol).value : "";
    const url = urlVal ? cellToString(urlVal).trim() : "";

    const curRank = normaliseRank(row.getCell(latest.col).value);
    const prevRank = previous
      ? normaliseRank(row.getCell(previous.col).value)
      : null;

    // Skip rows with no usable rank data at all
    if (curRank === null && prevRank === null) continue;

    keywords.push({
      keyword: kw,
      url,
      current: curRank,
      previous: prevRank
    });
  }

  if (!keywords.length) {
    throw new Error(
      "No usable keyword rows found. Check that your ranking sheet has keyword rows with numeric positions."
    );
  }

  const tracked = keywords.length;
  const withCurrent = keywords.filter((k) => k.current !== null);
  const withPrev = keywords.filter((k) => k.previous !== null);

  const avgCurrent =
    withCurrent.reduce((sum, k) => sum + k.current, 0) /
    (withCurrent.length || 1);
  const avgPrev =
    withPrev.length > 0
      ? withPrev.reduce((sum, k) => sum + k.previous, 0) / withPrev.length
      : 0;

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

  const hasPrevData = !!previous && withPrev.length > 0;

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
    hasPrevData,
    keywords,
    topWinners,
    topLosers
  };
}

// ---------- Backlink extraction ----------
function parseBacklinkSheet(sheet, sheetName) {
  // Find header row (search first 10 rows for "Backlinks")
  let headerRowNumber = null;
  let backlinkCol = null;
  let targetCol = null;
  let statusCol = null;

  for (let r = 1; r <= Math.min(10, sheet.rowCount); r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= row.cellCount; c++) {
      const label = cellToString(row.getCell(c).value)
        .trim()
        .toLowerCase();
      if (!label) continue;

      if (!backlinkCol && label.includes("backlink")) {
        backlinkCol = c;
        headerRowNumber = r;
      }
      if (
        !targetCol &&
        (label.includes("target") || label.includes("terget"))
      ) {
        targetCol = c;
        headerRowNumber = headerRowNumber || r;
      }
      if (!statusCol && label.includes("status")) {
        statusCol = c;
        headerRowNumber = headerRowNumber || r;
      }
    }
    if (backlinkCol && targetCol && headerRowNumber) break;
  }

  if (!headerRowNumber || !backlinkCol || !targetCol) {
    return { name: sheetName, total: 0, rows: [] };
  }

  const rows = [];
  for (let r = headerRowNumber + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const backlink = cellToString(row.getCell(backlinkCol).value).trim();
    const target = cellToString(row.getCell(targetCol).value).trim();
    const status = statusCol
      ? cellToString(row.getCell(statusCol).value).trim()
      : "";

    if (!backlink && !target) continue;

    rows.push({ backlink, target, status });
  }

  return {
    name: sheetName,
    total: rows.length,
    rows
  };
}

function parseBacklinks(workbook) {
  const sheetNames = [
    "All Backlinks",
    "Profile Backlinks",
    "Web 2.0 Backlinks",
    "Syndication Backlinks",
    "Article Submission",
    "Social Bookmarking Backlinks"
  ];

  const sections = [];
  for (const name of sheetNames) {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) continue;
    const parsed = parseBacklinkSheet(sheet, name);
    if (parsed.total > 0) sections.push(parsed);
  }

  const totalBacklinks = sections.reduce((sum, s) => sum + s.total, 0);

  return {
    totalBacklinks,
    sections
  };
}

// ---------- Read whole workbook ----------
async function parseSeoWorkbook(filePath) {
  const ExcelJS = await getExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const ranking = parseRankingSheet(workbook);
  const backlinks = parseBacklinks(workbook);

  return { ...ranking, backlinks };
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
    hasPrevData,
    keywords,
    topWinners,
    topLosers,
    backlinks
  } = summary;

  const doc = new PDFKit({ margin: 40 });

  res.statusCode = 200;
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="SEO-Report-${domain || "site"}.pdf"`
  );

  doc.pipe(res);

  const fmtPeriod = (info) => {
    if (!info) return "-";
    if (info.date) return info.date.toISOString().slice(0, 10);
    if (info.label) return info.label;
    return `Column ${info.col}`;
  };

  const top10Pct = tracked > 0 ? Math.round((top10 / tracked) * 100) : 0;
  const top10PrevPct =
    tracked > 0 ? Math.round((top10Prev / tracked) * 100) : 0;

  // ---- Cover page ----
  doc.fontSize(22).text("SEO Performance Report", { align: "center" });
  doc.moveDown(1.5);

  doc.fontSize(12);
  if (domain) doc.text(`Domain: ${domain}`);
  if (location) doc.text(`Location: ${location}`);
  doc.text(`Current period: ${fmtPeriod(latest)}`);
  if (previous) doc.text(`Compared with: ${fmtPeriod(previous)}`);
  doc.moveDown();

  doc.fontSize(12).text(`Tracked keywords: ${tracked}`);
  doc.text(`Average position (current): ${avgCurrent.toFixed(1)}`);
  if (hasPrevData) {
    doc.text(`Average position (previous): ${avgPrev.toFixed(1)}`);
    doc.text(
      `Top 10 visibility: ${top10Pct}% (previously ${top10PrevPct}%)`
    );
  } else {
    doc.text(`Top 10 visibility (current): ${top10Pct}%`);
  }

  if (backlinks && backlinks.totalBacklinks) {
    doc.moveDown();
    doc.fontSize(12).text(
      `Total backlinks created this period: ${backlinks.totalBacklinks}`
    );
  }

  doc.moveDown(1);
  doc.fontSize(10).fillColor("#555");
  doc.text(
    "This report was generated automatically from your uploaded SEO workbook."
  );

  // ---- Winners / losers page (if we have comparison data) ----
  if (topWinners.length || topLosers.length) {
    doc.addPage();
    doc.fontSize(16).fillColor("#000").text("Ranking Movement", {
      underline: true
    });
    doc.moveDown();

    doc.fontSize(13).text("Top Winners");
    doc.moveDown(0.5);
    doc.fontSize(10);
    if (!topWinners.length) {
      doc.text("No improving keywords this period.");
    } else {
      topWinners.forEach((k) => {
        doc.text(
          `• ${k.keyword}: ${k.previous ?? "-"} → ${k.current ?? "-"} (up ${
            k.change
          } places)`
        );
      });
    }

    doc.moveDown();
    doc.fontSize(13).text("Top Losers");
    doc.moveDown(0.5);
    doc.fontSize(10);
    if (!topLosers.length) {
      doc.text("No dropping keywords this period.");
    } else {
      topLosers.forEach((k) => {
        doc.text(
          `• ${k.keyword}: ${k.previous ?? "-"} → ${k.current ?? "-"} (down ${
            Math.abs(k.change)
          } places)`
        );
      });
    }
  }

  // ---- Keyword detail table ----
  doc.addPage();
  doc.fontSize(16).fillColor("#000").text("Keyword Detail", {
    underline: true
  });
  doc.moveDown(0.7);

  doc.fontSize(9).text("Keyword                           Prev  Curr  Change");
  doc.text("-------------------------------------------------------------");

  keywords.forEach((k) => {
    const prevStr =
      k.previous !== null ? String(k.previous).padStart(3) : " - ";
    const currStr =
      k.current !== null ? String(k.current).padStart(3) : " - ";
    let changeStr = "  - ";
    if (k.current !== null && k.previous !== null) {
      const diff = k.previous - k.current;
      if (diff > 0) changeStr = ` ↑${String(diff).padStart(2)}`;
      else if (diff < 0)
        changeStr = ` ↓${String(Math.abs(diff)).padStart(2)}`;
      else changeStr = "  0 ";
    }
    const kw =
      k.keyword.length > 30 ? k.keyword.slice(0, 27) + "..." : k.keyword;
    doc.text(`${kw.padEnd(30)}  ${prevStr}  ${currStr}  ${changeStr}`);

    if (doc.y > doc.page.height - doc.page.margins.bottom - 40) {
      doc.addPage();
      doc.fontSize(9).text(
        "Keyword                           Prev  Curr  Change"
      );
      doc.text(
        "-------------------------------------------------------------"
      );
    }
  });

  // ---- Backlinks overview ----
  if (backlinks && backlinks.sections && backlinks.sections.length) {
    doc.addPage();
    doc.fontSize(16).fillColor("#000").text("Backlinks Overview", {
      underline: true
    });
    doc.moveDown();

    doc.fontSize(12).text(
      `Total backlinks in this workbook: ${backlinks.totalBacklinks}`
    );
    doc.moveDown(0.5);

    backlinks.sections.forEach((section, idx) => {
      if (idx > 0) doc.moveDown(0.5);
      doc.fontSize(12).text(section.name);
      doc.fontSize(10).text(
        `Links in this category: ${section.total}`,
        { indent: 10 }
      );
    });

    // One page per category with sample links
    backlinks.sections.forEach((section) => {
      doc.addPage();
      doc.fontSize(16).fillColor("#000").text(section.name, {
        underline: true
      });
      doc.moveDown(0.5);
      doc.fontSize(11).text(
        `Total links: ${section.total}. Showing first ${
          section.rows.length > 10 ? 10 : section.rows.length
        } links:`
      );
      doc.moveDown(0.5);
      doc.fontSize(9);

      section.rows.slice(0, 10).forEach((row, i) => {
        const statusText = row.status ? ` [${row.status}]` : "";
        doc.text(
          `${i + 1}. ${row.backlink} → ${row.target}${statusText}`
        );
        doc.moveDown(0.2);
      });

      if (section.rows.length > 10) {
        doc.moveDown(0.5);
        doc.fontSize(10).text(
          `... plus ${section.rows.length - 10} more links in this category.`
        );
      }
    });
  }

  doc.end();
}

// ---------- Main handler ----------
export default async function handler(req, res) {
  try {
    if (req.method === "GET") {
      res.statusCode = 200;
      return res.json({ ok: true, message: "SEO reports API is alive" });
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
