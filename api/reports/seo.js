// api/reports/seo.js
//
// Vercel/Next.js API route:
//
// - Accepts multipart/form-data with field "seo_file"
// - Reads the SEO Excel workbook with exceljs
// - Extracts ranking data from the "Ranking" sheet
//   using the two right-most date columns
// - Extracts backlinks from the backlink sheets
// - Generates a branded, human-readable PDF with pdfkit

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

  // ranks must be positive and not insane
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

  // Normalise "Sept" to "Sep"
  text = text.replace(/\bSept\b/gi, "Sep");

  // Normalise "1 Sep-25" -> "1 Sep 2025"
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

// ---------- Ranking extraction tuned to your sheet ----------
function parseRankingSheet(workbook) {
  // Prefer "Ranking" sheet; fallback to first sheet
  let sheet = workbook.getWorksheet("Ranking");
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("No worksheets found in uploaded file.");

  const maxHeaderRows = Math.min(15, sheet.rowCount);

  // ---- Find date columns by scanning top rows ----
  const dateByCol = new Map(); // col -> { col, date, row }
  const rowFreq = new Map(); // row -> count of date cells

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

  // latest / previous by actual date
  dateCandidates.sort((a, b) => a.date - b.date);
  const latest = dateCandidates[dateCandidates.length - 1];
  const previous =
    dateCandidates.length >= 2
      ? dateCandidates[dateCandidates.length - 2]
      : null;

  // choose the row that most often has date headers as the header row
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

  // ---- Domain & location (optional) ----
  let domain = "";
  let location = "";
  const domainRegex = /[a-z0-9.-]+\.[a-z]{2,}/i;

  for (let r = 1; r <= maxHeaderRows; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= row.cellCount; c++) {
      const text = cellToString(row.getCell(c).value).trim();
      if (!text) continue;

      if (!domain && domainRegex.test(text)) {
        domain = text;
      }
      if (
        !location &&
        /new zealand|australia|united kingdom|united states|usa/i.test(text)
      ) {
        location = text;
      }
    }
  }

  // ---- Keywords: column A is the phrase ----
  const KEYWORD_COL = 1;

  const keywords = [];

  for (let r = firstDataRow; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);

    const curRank = normaliseRank(row.getCell(latest.col).value);
    const prevRank = previous
      ? normaliseRank(row.getCell(previous.col).value)
      : null;

    // Skip rows with no rank data
    if (curRank === null && prevRank === null) continue;

    const kwVal = row.getCell(KEYWORD_COL).value;
    let keyword = cellToString(kwVal).trim();

    if (!keyword) {
      // if for some reason col A is empty, fall back to any text on the row
      for (let c = 1; c < latest.col; c++) {
        const txt = cellToString(row.getCell(c).value).trim();
        if (txt && txt.length > keyword.length) keyword = txt;
      }
    }

    if (!keyword) keyword = "(keyword)";

    // URL: any http(s) cell on the row
    let url = "";
    for (let c = 1; c <= row.cellCount; c++) {
      const txt = cellToString(row.getCell(c).value).trim();
      if (
        txt &&
        (txt.startsWith("http://") || txt.startsWith("https://"))
      ) {
        url = txt;
        break;
      }
    }

    keywords.push({
      keyword,
      url,
      current: curRank,
      previous: prevRank
    });
  }

  if (!keywords.length) {
    throw new Error(
      "No usable keyword rows found. Check that your ranking sheet has rows with ranks in the last date columns."
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

  // “page 1” + position-1 / position-2 stats
  const page1Count = withCurrent.filter((k) => k.current <= 10).length;
  const pos1Count = withCurrent.filter((k) => k.current === 1).length;
  const pos2Count = withCurrent.filter((k) => k.current === 2).length;

  // medians give a more “typical” position
  const currentRanksSorted = withCurrent
    .map((k) => k.current)
    .sort((a, b) => a - b);
  let medianCurrent = 0;
  if (currentRanksSorted.length) {
    const mid = Math.floor(currentRanksSorted.length / 2);
    if (currentRanksSorted.length % 2 === 1) {
      medianCurrent = currentRanksSorted[mid];
    } else {
      medianCurrent =
        (currentRanksSorted[mid - 1] + currentRanksSorted[mid]) / 2;
    }
  }

  const prevRanksSorted = withPrev
    .map((k) => k.previous)
    .sort((a, b) => a - b);
  let medianPrev = 0;
  if (prevRanksSorted.length) {
    const mid = Math.floor(prevRanksSorted.length / 2);
    if (prevRanksSorted.length % 2 === 1) {
      medianPrev = prevRanksSorted[mid];
    } else {
      medianPrev =
        (prevRanksSorted[mid - 1] + prevRanksSorted[mid]) / 2;
    }
  }

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
    medianCurrent,
    medianPrev,
    top10,
    top10Prev,
    page1Count,
    pos1Count,
    pos2Count,
    hasPrevData,
    keywords,
    topWinners,
    topLosers
  };
}

// ---------- Backlink extraction ----------
function parseBacklinkSheet(sheet, sheetName) {
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

// ---------- Parse entire workbook ----------
async function parseSeoWorkbook(filePath) {
  const ExcelJS = await getExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const ranking = parseRankingSheet(workbook);
  const backlinks = parseBacklinks(workbook);

  return { ...ranking, backlinks };
}

// ---------- PDF generation (branded layout) ----------
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
    medianCurrent,
    medianPrev,
    top10,
    top10Prev,
    page1Count,
    pos1Count,
    pos2Count,
    hasPrevData,
    keywords,
    topWinners,
    topLosers,
    backlinks
  } = summary;

  const doc = new PDFKit({ size: "A4", margin: 40 });

  // --- HTTP headers ---
  res.statusCode = 200;
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="SEO-Report-${domain || "site"}.pdf"`
  );

  doc.pipe(res);

  // --- Brand colours ---
  const brandDark = "#222222";
  const brandBlue = "#1976d2";
  const brandRed = "#e53935";
  const brandGreen = "#43a047";

  const pageWidth = doc.page.width;
  const left = doc.page.margins.left;
  const right = pageWidth - doc.page.margins.right;
  const contentWidth = right - left;

  // Paths to hero + logo (adjust filenames if yours are different)
  const heroPath = path.join(process.cwd(), "public", "seo-hero.jpg");
  const logoPath = path.join(process.cwd(), "public", "webxperts-logo.png");

  const fmtPeriod = (info) => {
    if (!info) return "-";
    if (info.date) return info.date.toISOString().slice(0, 10);
    return `Col ${info.col}`;
  };

  const top10Pct = tracked > 0 ? Math.round((top10 / tracked) * 100) : 0;
  const top10PrevPct =
    tracked > 0 ? Math.round((top10Prev / tracked) * 100) : 0;

  // ======================================================
  // PAGE 1 – HERO + HIGH LEVEL SUMMARY
  // ======================================================

  const heroHeight = 180;

  if (fs.existsSync(heroPath)) {
    // hero image with dark overlay
    doc.image(heroPath, 0, 0, { width: pageWidth, height: heroHeight });
    doc
      .save()
      .rect(0, 0, pageWidth, heroHeight)
      .fillOpacity(0.5)
      .fill(brandDark)
      .fillOpacity(1)
      .restore();
  } else {
    // fallback: solid band
    doc
      .save()
      .rect(0, 0, pageWidth, heroHeight)
      .fill(brandDark)
      .restore();
  }

  doc
    .fillColor("#ffffff")
    .fontSize(22)
    .text("SEO Monthly Report", left, 60);

  doc.fontSize(14);
  if (domain) doc.text(domain, left, 95);
  if (location) doc.text(location, left, 115);

  const periodText = `Current period: ${fmtPeriod(
    latest
  )}   ·   Previous: ${fmtPeriod(previous)}`;
  doc.text(periodText, left, 135);

  // White body background
  doc
    .save()
    .rect(0, heroHeight, pageWidth, doc.page.height - heroHeight)
    .fill("#ffffff")
    .restore();

  doc.y = heroHeight + 30;

  // Metric cards (3 per row)
  const cardGap = 10;
  const cardsPerRow = 3;
  const cardWidth = (contentWidth - cardGap * (cardsPerRow - 1)) / cardsPerRow;
  const cardHeight = 60;

  const metricCards = [
    {
      label: "Tracked Keywords",
      value: tracked.toString(),
      color: brandBlue
    },
    {
      label: "Page 1 Keywords (1–10)",
      value: `${page1Count}/${tracked}`,
      color: brandGreen
    },
    {
      label: "Top 10 Visibility",
      value: `${top10Pct}%`,
      color: brandRed
    },
    {
      label: "Pos #1",
      value: pos1Count.toString(),
      color: brandBlue
    },
    {
      label: "Pos #2",
      value: pos2Count.toString(),
      color: brandGreen
    },
    {
      label: "Median Position",
      value: medianCurrent.toFixed(1),
      color: brandRed
    }
  ];

  if (backlinks && backlinks.totalBacklinks) {
    metricCards.push({
      label: "Backlinks in Workbook",
      value: backlinks.totalBacklinks.toString(),
      color: brandBlue
    });
  }

  let cardIndex = 0;
  metricCards.forEach((card) => {
    const row = Math.floor(cardIndex / cardsPerRow);
    const col = cardIndex % cardsPerRow;
    const x = left + col * (cardWidth + cardGap);
    const y = heroHeight + 30 + row * (cardHeight + cardGap);

    // coloured box
    doc
      .save()
      .rect(x, y, cardWidth, cardHeight)
      .fill(card.color)
      .restore();

    // text inside
    doc
      .fillColor("#ffffff")
      .fontSize(11)
      .text(card.label, x + 8, y + 8, {
        width: cardWidth - 16
      });

    doc
      .fontSize(20)
      .font("Helvetica-Bold")
      .text(card.value, x + 8, y + 26, {
        width: cardWidth - 16,
        align: "left"
      })
      .font("Helvetica");

    cardIndex++;
  });

  // Logo near bottom-right of page 1 (if present)
  if (fs.existsSync(logoPath)) {
    doc.image(logoPath, right - 140, doc.page.height - 100, { width: 120 });
  }

  // footer line on page 1
  doc
    .moveTo(left, doc.page.height - 50)
    .lineTo(right, doc.page.height - 50)
    .strokeColor("#dddddd")
    .lineWidth(1)
    .stroke();

  doc
    .fontSize(9)
    .fillColor("#777777")
    .text(
      "WebXperts SEO Report – generated automatically from your weekly ranking workbook.",
      left,
      doc.page.height - 45,
      { width: contentWidth }
    );

  // ======================================================
  // PAGE 2 – NARRATIVE OVERVIEW
  // ======================================================
  doc.addPage();

  doc
    .fontSize(16)
    .fillColor("#000000")
    .text("Overview", left, doc.y, { underline: true });

  doc.moveDown();

  const introLines = [];

  introLines.push(
    `We are currently tracking ${tracked} keywords for your website${
      domain ? " " + domain : ""
    }.`
  );

  if (hasPrevData) {
    introLines.push(
      `Average ranking moved from ${avgPrev.toFixed(
        1
      )} to ${avgCurrent.toFixed(1)}, with a typical (median) position of ${medianCurrent.toFixed(
        1
      )}.`
    );
  } else {
    introLines.push(
      `Average ranking this period is ${avgCurrent.toFixed(
        1
      )}, with a typical (median) position of ${medianCurrent.toFixed(1)}.`
    );
  }

  introLines.push(
    `${page1Count} of your ${tracked} keywords are currently on page 1 (positions 1–10).`
  );
  introLines.push(
    `${pos1Count} keywords are sitting in position 1 and ${pos2Count} are in position 2.`
  );

  if (hasPrevData) {
    introLines.push(
      `Top 10 visibility is now ${top10Pct}% (previously ${top10PrevPct}%).`
    );
  }

  if (backlinks && backlinks.totalBacklinks) {
    introLines.push(
      `This workbook also records ${backlinks.totalBacklinks} backlinks created across your various campaigns.`
    );
  }

  doc
    .fontSize(11)
    .fillColor("#333333")
    .text(introLines.join(" "), {
      width: contentWidth,
      align: "left"
    });

  // ======================================================
  // PAGE 3 – RANKING MOVEMENT
  // ======================================================
  if (topWinners.length || topLosers.length) {
    doc.addPage();

    doc
      .fontSize(16)
      .fillColor("#000000")
      .text("Ranking Movement", left, doc.y, { underline: true });
    doc.moveDown();

    // Winners
    doc
      .fontSize(13)
      .fillColor(brandGreen)
      .text("Top Gainers", left, doc.y);
    doc.moveDown(0.3);

    doc.fontSize(10).fillColor("#000000");
    if (!topWinners.length) {
      doc.text("No improving keywords this period.");
    } else {
      topWinners.forEach((k) => {
        doc.text(
          `• ${k.keyword}: ${k.previous ?? "-"} → ${
            k.current ?? "-"
          } (up ${k.change} places)`
        );
      });
    }

    doc.moveDown();

    // Losers
    doc
      .fontSize(13)
      .fillColor(brandRed)
      .text("Top Decliners", left, doc.y);
    doc.moveDown(0.3);

    doc.fontSize(10).fillColor("#000000");
    if (!topLosers.length) {
      doc.text("No dropping keywords this period.");
    } else {
      topLosers.forEach((k) => {
        doc.text(
          `• ${k.keyword}: ${k.previous ?? "-"} → ${
            k.current ?? "-"
          } (down ${Math.abs(k.change)} places)`
        );
      });
    }
  }

  // ======================================================
  // PAGE 4+ – KEYWORD TABLE
  // ======================================================
  doc.addPage();
  doc
    .fontSize(16)
    .fillColor("#000000")
    .text("Keyword Detail", left, doc.y, { underline: true });

  doc.moveDown(0.7);

  doc.fontSize(9);
  doc.fillColor("#000000");
  doc.text("Keyword                           Prev  Curr  Change");
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

  // ======================================================
  // BACKLINKS PAGES
  // ======================================================
  if (backlinks && backlinks.sections && backlinks.sections.length) {
    // Overview
    doc.addPage();
    doc
      .fontSize(16)
      .fillColor("#000000")
      .text("Backlinks Overview", left, doc.y, { underline: true });
    doc.moveDown();

    doc
      .fontSize(12)
      .fillColor("#333333")
      .text(
        `Total backlinks in this workbook: ${backlinks.totalBacklinks}`,
        { width: contentWidth }
      );
    doc.moveDown(0.5);

    backlinks.sections.forEach((section, idx) => {
      if (idx > 0) doc.moveDown(0.5);
      doc.fontSize(12).fillColor("#000000").text(section.name);
      doc
        .fontSize(10)
        .fillColor("#555555")
        .text(`Links in this category: ${section.total}`, {
          indent: 10
        });
    });

    // One page per backlink category
    backlinks.sections.forEach((section) => {
      doc.addPage();
      doc
        .fontSize(16)
        .fillColor("#000000")
        .text(section.name, left, doc.y, { underline: true });
      doc.moveDown(0.5);

      doc
        .fontSize(11)
        .fillColor("#333333")
        .text(
          `Total links: ${section.total}. Showing first ${
            section.rows.length > 10 ? 10 : section.rows.length
          } links:`,
          { width: contentWidth }
        );
      doc.moveDown(0.5);

      doc.fontSize(9).fillColor("#000000");
      section.rows.slice(0, 10).forEach((row, i) => {
        const statusText = row.status ? ` [${row.status}]` : "";
        doc.text(
          `${i + 1}. ${row.backlink} → ${row.target}${statusText}`
        );
        doc.moveDown(0.2);
      });

      if (section.rows.length > 10) {
        doc.moveDown(0.5);
        doc
          .fontSize(10)
          .fillColor("#555555")
          .text(
            `... plus ${section.rows.length - 10} more links in this category.`,
            { width: contentWidth }
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
