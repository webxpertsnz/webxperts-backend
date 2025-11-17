// api/reports/seo.js
//
// Vercel/Next.js API route:
// - Accepts multipart/form-data with field "seo_file"
// - Reads an SEO ranking Excel file with exceljs
// - Auto-detects keyword + rank columns
// - Generates a PDF summary with pdfkit

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

// flexible date-ish parser (but we no longer *require* dates)
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
    /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(text) || // 2024-11-01, 01/11/2024
    /^\d{8}$/.test(text) || // 20241101
    (/\d/.test(text) && (text.includes("/") || text.includes("-")));

  if (!looksLikeDate) return null;

  const dt = new Date(text);
  return isNaN(dt) ? null : dt;
}

function cellToString(v) {
  if (!v && v !== 0) return "";
  if (typeof v === "object" && v !== null) {
    if (v.text) return String(v.text);
    if (v.result) return String(v.result);
    if (Array.isArray(v.richText)) {
      return v.richText.map((p) => p.text || "").join("");
    }
  }
  return String(v);
}

// ---------- Excel parsing ----------
async function parseSeoWorkbook(filePath) {
  const ExcelJS = await getExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Prefer a sheet called "Ranking", otherwise first sheet
  let sheet = workbook.getWorksheet("Ranking");
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("No worksheets found in uploaded file.");

  // Domain / location in first 10 rows (optional)
  let domain = "";
  let location = "";
  for (let r = 1; r <= Math.min(10, sheet.rowCount); r++) {
    const a = cellToString(sheet.getCell(r, 1)).trim().toLowerCase();
    const b = sheet.getCell(r, 2).value;
    if (a === "domain" || a === "website") domain = b ? String(b) : "";
    if (a === "location") location = b ? String(b) : "";
  }

  // ---------- find header row & keyword column ----------
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
      if (text === "keyword" || text === "keywords" || text.includes("keyword")) {
        headerRowNumber = r;
        keywordCol = c;
        break outer;
      }
    }
  }

  // Pass 2: fallback — first "busy row" with 3+ non-empty cells
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
        "Make sure your ranking sheet has a header row with a 'Keyword' column."
    );
  }

  const headerRow = sheet.getRow(headerRowNumber);

  // ---------- detect URL column ----------
  let urlCol = null;
  const columnCount = sheet.columnCount || headerRow.cellCount;
  for (let c = 1; c <= columnCount; c++) {
    const label = cellToString(headerRow.getCell(c).value)
      .trim()
      .toLowerCase();
    if (
      ["url", "landing page", "page", "target url", "address"].includes(
        label
      )
    ) {
      urlCol = c;
      break;
    }
  }

  // ---------- detect ranking columns (latest / previous) ----------
  const dateCandidates = [];
  const genericRankCandidates = [];

  for (let c = 1; c <= columnCount; c++) {
    const raw = headerRow.getCell(c).value;
    const label = cellToString(raw).trim();
    const labelLower = label.toLowerCase();

    if (!label) continue;

    const dt = parseHeaderDate(raw);
    if (dt) {
      // date-based candidate
      dateCandidates.push({ col: c, label, date: dt });
      continue;
    }

    // label-based candidates (no dates)
    const isRanky =
      labelLower.includes("rank") ||
      labelLower.includes("position") ||
      labelLower.includes("pos") ||
      labelLower.includes("google") ||
      labelLower.includes("bing") ||
      labelLower.includes("yahoo") ||
      labelLower.includes("current") ||
      labelLower.includes("previous") ||
      labelLower.includes("last") ||
      labelLower.includes("week") ||
      labelLower.includes("month");

    if (isRanky && c > keywordCol) {
      genericRankCandidates.push({ col: c, label, date: null });
    }
  }

  let latest = null;
  let previous = null;

  if (dateCandidates.length) {
    // Use real dates if present
    dateCandidates.sort((a, b) => a.date - b.date);
    latest = dateCandidates[dateCandidates.length - 1];
    if (dateCandidates.length >= 2) {
      previous = dateCandidates[dateCandidates.length - 2];
    }
  } else if (genericRankCandidates.length) {
    // Use label-based rank columns
    // Try to pick 'current' as latest, 'previous/last' as previous
    const cur = genericRankCandidates.find((c) =>
      c.label.toLowerCase().includes("current")
    );
    const prev = genericRankCandidates.find((c) => {
      const l = c.label.toLowerCase();
      return l.includes("prev") || l.includes("last");
    });

    latest = cur || genericRankCandidates[genericRankCandidates.length - 1];
    if (prev && prev.col !== latest.col) {
      previous = prev;
    } else if (genericRankCandidates.length >= 2) {
      // the column just before latest
      const idx = genericRankCandidates.findIndex(
        (c) => c.col === latest.col
      );
      if (idx > 0) previous = genericRankCandidates[idx - 1];
    }
  }

  // Fallback: scan numeric columns to the right of Keyword
  if (!latest) {
    const numericCols = new Set();
    const maxSampleRows = Math.min(sheet.rowCount, headerRowNumber + 50);

    for (let r = headerRowNumber + 1; r <= maxSampleRows; r++) {
      const row = sheet.getRow(r);
      for (let c = keywordCol + 1; c <= columnCount; c++) {
        const val = row.getCell(c).value;
        if (typeof val === "number") {
          numericCols.add(c);
        } else if (val && !isNaN(Number(val))) {
          numericCols.add(c);
        }
      }
    }

    const cols = Array.from(numericCols).sort((a, b) => a - b);
    if (cols.length) {
      const lastCol = cols[cols.length - 1];
      const prevCol = cols.length >= 2 ? cols[cols.length - 2] : null;

      latest = {
        col: lastCol,
        label: cellToString(headerRow.getCell(lastCol).value).trim() || `Col ${lastCol}`,
        date: null
      };

      if (prevCol) {
        previous = {
          col: prevCol,
          label: cellToString(headerRow.getCell(prevCol).value).trim() || `Col ${prevCol}`,
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

  // ---------- collect keyword rows ----------
  const keywords = [];
  for (let r = headerRowNumber + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const kwVal = row.getCell(keywordCol).value;
    if (!kwVal) continue;
    const kw = cellToString(kwVal).trim();
    if (!kw) continue;

    const urlVal = urlCol ? row.getCell(urlCol).value : "";
    const url = urlVal ? cellToString(urlVal) : "";

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
        "Check that your ranking sheet has data rows below the header."
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

  const fmtPeriod = (info) => {
    if (!info) return "-";
    if (info.date) return info.date.toISOString().slice(0, 10);
    if (info.label) return info.label;
    return `Column ${info.col}`;
  };

  const top10Pct = tracked > 0 ? Math.round((top10 / tracked) * 100) : 0;
  const top10PrevPct =
    tracked > 0 ? Math.round((top10Prev / tracked) * 100) : 0;

  // Cover
  doc.fontSize(18).text("SEO Ranking Report");
  doc.moveDown(0.5);
  doc.fontSize(11).fillColor("#555");
  if (domain) doc.text(`Domain: ${domain}`);
  if (location) doc.text(`Location: ${location}`);
  doc.text(`Current period: ${fmtPeriod(latest)}`);
  if (previous) doc.text(`Compared with: ${fmtPeriod(previous)}`);
  doc.moveDown();

  doc.fontSize(12).fillColor("#000");
  doc.text(`Tracked keywords: ${tracked}`);
  doc.text(`Average position (current): ${avgCurrent.toFixed(1)}`);
  if (hasPrevData) {
    doc.text(`Average position (previous): ${avgPrev.toFixed(1)}`);
    doc.text(
      `Top 10 visibility: ${top10Pct}% (prev ${top10PrevPct}%)`
    );
  } else {
    doc.text(`Top 10 visibility (current): ${top10Pct}%`);
  }
  doc.moveDown();

  doc.fontSize(11).fillColor("#555");
  if (hasPrevData) {
    doc.text(
      `This period we are tracking ${tracked} keywords. ` +
        `Average ranking changed from ${avgPrev.toFixed(
          1
        )} to ${avgCurrent.toFixed(
          1
        )}, with ${top10Pct}% of keywords currently in the top 10.`
    );
  } else {
    doc.text(
      `This period we are tracking ${tracked} keywords. ` +
        `Average ranking is ${avgCurrent.toFixed(
          1
        )}, with ${top10Pct}% of keywords in the top 10.`
    );
  }

  // Winners / losers (only if we have previous data)
  if (topWinners.length || topLosers.length) {
    doc.addPage();
    doc.fontSize(14).fillColor("#000").text("Top Winners", {
      underline: true
    });
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
      return res.json({ error: "Failed to generate SEO report: " + msg });
    }
  }
}
