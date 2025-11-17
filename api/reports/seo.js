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
    const row = sheet.getRow(r);

    const curRank = normaliseRank(row.getCell(latest.col).value);
    const prevRank = previous
      ? normaliseRank(row.getCell(previous.col).value)
      : null;

    if (curRank === null && prevRank === null) continue;

    const kwVal = row.getCell(KEYWORD_COL).value;
    let keyword = cellToString(kwVal).trim();

    if (!keyword) {
      for (let c = 1; c < latest.col; c++) {
        const txt = cellToString(row.getCell(c).value).trim();
        if (txt && txt.length > keyword.length) keyword = txt;
      }
    }
    if (!keyword) keyword = "(keyword)";

    let url = "";
    for (let c = 1; c <= row.cellCount; c++) {
      const txt = cellToString(row.getCell(c).value).trim();
      if (txt && (txt.startsWith("http://") || txt.startsWith("https://"))) {
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

  const page1Count = withCurrent.filter((k) => k.current <= 10).length;
  const pos1Count = withCurrent.filter((k) => k.current === 1).length;
  const pos2Count = withCurrent.filter((k) => k.current === 2).length;
  const pos3Count = withCurrent.filter((k) => k.current === 3).length;

  const currentRanksSorted = withCurrent
    .map((k) => k.current)
    .sort((a, b) => a - b);
  let medianCurrent = 0;
  if (currentRanksSorted.length) {
    const mid = Math.floor(currentRanksSorted.length / 2);
    medianCurrent =
      currentRanksSorted.length % 2 === 1
        ? currentRanksSorted[mid]
        : (currentRanksSorted[mid - 1] + currentRanksSorted[mid]) / 2;
  }

  const prevRanksSorted = withPrev
    .map((k) => k.previous)
    .sort((a, b) => a - b);
  let medianPrev = 0;
  if (prevRanksSorted.length) {
    const mid = Math.floor(prevRanksSorted.length / 2);
    medianPrev =
      prevRanksSorted.length % 2 === 1
        ? prevRanksSorted[mid]
        : (prevRanksSorted[mid - 1] + prevRanksSorted[mid]) / 2;
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
    pos3Count,
    hasPrevData,
    keywords,
    topWinners,
    topLosers
  };
}

// ---------- Backlink extraction (more robust header detection) ----------
function parseBacklinkSheet(sheet, sheetName) {
  const maxHeaderRows = Math.min(10, sheet.rowCount);

  // First pass: detect which rows look like headers (have "Backlinks")
  const candidateRows = [];
  for (let r = 1; r <= maxHeaderRows; r++) {
    const row = sheet.getRow(r);
    let hasBacklink = false;
    let hasTarget = false;
    let hasStatus = false;

    for (let c = 1; c <= row.cellCount; c++) {
      const label = cellToString(row.getCell(c).value)
        .trim()
        .toLowerCase();
      if (!label) continue;

      if (label === "backlinks") hasBacklink = true;
      if (label.includes("target") || label.includes("terget")) hasTarget = true;
      if (label.includes("status")) hasStatus = true;
    }

    if (hasBacklink) {
      candidateRows.push({ row: r, hasTarget, hasStatus });
    }
  }

  if (!candidateRows.length) {
    return { name: sheetName, total: 0, rows: [] };
  }

  // Prefer a row that also has Target or Status, otherwise first with Backlinks
  candidateRows.sort((a, b) => {
    const scoreA = (a.hasTarget ? 1 : 0) + (a.hasStatus ? 1 : 0);
    const scoreB = (b.hasTarget ? 1 : 0) + (b.hasStatus ? 1 : 0);
    return scoreB - scoreA || a.row - b.row;
  });

  const headerRowNumber = candidateRows[0].row;
  const headerRow = sheet.getRow(headerRowNumber);

  let backlinkCol = null;
  let targetCol = null;
  let statusCol = null;

  headerRow.eachCell((cell, col) => {
    const label = cellToString(cell.value).trim().toLowerCase();
    if (!label) return;
    if (label === "backlinks") backlinkCol = col;
    if (
      label === "target url" ||
      label === "target" ||
      label === "terget url" ||
      label === "terget"
    ) {
      targetCol = col;
    }
    if (label === "status") statusCol = col;
  });

  if (!backlinkCol) {
    // fall back: any cell on that row containing "backlink"
    headerRow.eachCell((cell, col) => {
      const label = cellToString(cell.value).trim().toLowerCase();
      if (label.includes("backlink") && !backlinkCol) backlinkCol = col;
    });
  }

  if (!backlinkCol) {
    return { name: sheetName, total: 0, rows: [] };
  }

  const rows = [];
  for (let r = headerRowNumber + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const backlink = cellToString(row.getCell(backlinkCol).value).trim();
    const target = targetCol
      ? cellToString(row.getCell(targetCol).value).trim()
      : "";
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

// format as dd/mm/yyyy
function formatDateNZ(info) {
  if (!info || !info.date) return "-";
  const d = info.date;
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function describeBacklinkType(name) {
  const lower = name.toLowerCase();
  if (lower.includes("profile")) {
    return "Profile backlinks are links from profile pages on business listings or social platforms. They help build brand signals and basic authority.";
  }
  if (lower.includes("web 2.0")) {
    return "Web 2.0 backlinks come from content published on hosted blog platforms. They support topical relevance and can drive referral traffic.";
  }
  if (lower.includes("syndication")) {
    return "Syndication backlinks are created when your content is republished on other sites, spreading your brand and earning contextual links.";
  }
  if (lower.includes("article")) {
    return "Article submission backlinks are links gained from publishing articles on external sites, usually with contextual anchor text.";
  }
  if (lower.includes("social bookmarking")) {
    return "Social bookmarking backlinks are links from bookmarking sites where your content is saved and shared, helping with discovery and indexing.";
  }
  if (lower.includes("all backlinks")) {
    return "This section summarises all backlinks recorded in your workbook across every campaign type.";
  }
  return "These backlinks contribute to your overall authority and help search engines discover and trust your website.";
}

// ---------- PDF generation ----------
async function buildSeoPdf(res, summary) {
  const PDFKit = await getPdfKit();
  const {
    domain,
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
    pos3Count,
    hasPrevData,
    keywords,
    topWinners,
    topLosers,
    backlinks
  } = summary;

  const doc = new PDFKit({ size: "A4", margin: 40 });

  res.statusCode = 200;
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="SEO-Report-${domain || "site"}.pdf"`
  );

  doc.pipe(res);

  const brandDark = "#222222";
  const brandBlue = "#1976d2";
  const brandGreen = "#43a047";
  const brandOrange = "#fb8c00";

  const pageWidth = doc.page.width;
  const left = doc.page.margins.left;
  const right = pageWidth - doc.page.margins.right;
  const contentWidth = right - left;

  const heroPath = path.join(process.cwd(), "public", "IMG_0903.jpeg");

  const prevPage1 = hasPrevData ? top10Prev : null;
  const top3Count = pos1Count + pos2Count + pos3Count;

  let performanceTrend = "stable performance";
  let performanceDeltaPct = 0;
  let performanceDirection = "changed";

  if (hasPrevData && avgPrev > 0) {
    const diff = avgPrev - avgCurrent;
    performanceDeltaPct = Math.round(Math.abs((diff / avgPrev) * 100));

    if (diff > 0.5) {
      performanceTrend = "strong growth";
      performanceDirection = "improved";
    } else if (diff > 0) {
      performanceTrend = "slight improvement";
      performanceDirection = "improved";
    } else if (diff < -0.5) {
      performanceTrend = "a decline";
      performanceDirection = "declined";
    } else if (diff < 0) {
      performanceTrend = "a slight decline";
      performanceDirection = "declined";
    } else {
      performanceTrend = "no significant change";
      performanceDirection = "changed";
    }
  }

  const newTop10 =
    hasPrevData && top10 > top10Prev ? top10 - top10Prev : 0;

  // ======================================================
  // PAGE 1 – HERO + 3 BOXES + EXEC SUMMARY
  // ======================================================

  const heroHeight = 180;

  if (fs.existsSync(heroPath)) {
    doc.image(heroPath, 0, 0, { width: pageWidth, height: heroHeight });
    doc
      .save()
      .rect(0, 0, pageWidth, heroHeight)
      .fillOpacity(0.7)
      .fill(brandDark)
      .fillOpacity(1)
      .restore();
  } else {
    doc
      .save()
      .rect(0, 0, pageWidth, heroHeight)
      .fill(brandDark)
      .restore();
  }

  doc
    .fillColor("#ffffff")
    .fontSize(22)
    .text("SEO Report", left, 60);

  doc.fontSize(14);
  if (domain) doc.text(domain, left, 95);

  const dateText = `Report date: ${formatDateNZ(latest)}`;
  doc.text(dateText, left, 115);

  doc
    .save()
    .rect(0, heroHeight, pageWidth, doc.page.height - heroHeight)
    .fill("#ffffff")
    .restore();

  const cardGap = 10;
  const cardsPerRow = 3;
  const cardWidth = (contentWidth - cardGap * (cardsPerRow - 1)) / cardsPerRow;
  const cardHeight = 60;

  const metricCards = [
    {
      label: "Optimised keywords",
      value: tracked.toString(),
      color: brandBlue
    },
    {
      label: "Keywords on page 1",
      value: `${page1Count}/${tracked}`,
      color: brandGreen
    },
    {
      label: "Keywords in top position",
      value: pos1Count.toString(),
      color: brandOrange
    }
  ];

  let cardIndex = 0;
  const cardTopY = heroHeight + 30;

  metricCards.forEach((card) => {
    const col = cardIndex % cardsPerRow;
    const x = left + col * (cardWidth + cardGap);
    const y = cardTopY;

    doc
      .save()
      .rect(x, y, cardWidth, cardHeight)
      .fill(card.color)
      .restore();

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
        width: cardWidth - 16
      })
      .font("Helvetica");

    cardIndex++;
  });

  const yAfterCards = cardTopY + cardHeight + 25;

  // ---------- Executive Summary block with nicer spacing ----------
  doc.y = yAfterCards;

  doc
    .fontSize(14)
    .fillColor("#000000")
    .text("Executive Summary", left, doc.y, { underline: true });

  doc.moveDown(0.7);

  doc.fontSize(11).fillColor("#333333");

  if (hasPrevData && performanceDeltaPct > 0) {
    doc.text(
      `This month’s SEO performance shows ${performanceTrend} in key metrics. ` +
        `Performance ${performanceDirection} by approximately ${performanceDeltaPct}% compared to last week.`,
      { width: contentWidth }
    );
  } else {
    doc.text(
      "This month’s SEO performance establishes a baseline for ongoing tracking. Previous-week comparison data is not available in this workbook.",
      { width: contentWidth }
    );
  }

  doc.moveDown(0.5);

  const summaryLine = hasPrevData
    ? `We currently hold ${page1Count} positions on the first page of Google for targeted keywords, with ${newTop10} new top-10 rankings achieved since last week.`
    : `We currently hold ${page1Count} positions on the first page of Google for targeted keywords.`;

  doc.text(summaryLine, { width: contentWidth });

  // Keyword rankings sub-section
  doc.moveDown(1.0);
  doc.fontSize(12).fillColor("#000000").text("Keyword Rankings:");
  doc.moveDown(0.4);
  doc.fontSize(11).fillColor("#333333");

  const page1Change =
    hasPrevData && prevPage1 !== null
      ? page1Count - prevPage1
      : null;

  const krLines = [];
  krLines.push(`Total tracked keywords: ${tracked}.`);

  if (hasPrevData && prevPage1 !== null) {
    const dir =
      page1Change > 0
        ? "up"
        : page1Change < 0
        ? "down"
        : "the same as";
    krLines.push(
      `First-page positions: ${page1Count} (${dir} ${Math.abs(
        page1Change
      ) || ""} from ${prevPage1} last week).`
    );
  } else {
    krLines.push(`First-page positions: ${page1Count}.`);
  }

  krLines.push(`Top-3 positions: ${top3Count}.`);

  if (topWinners.length) {
    const improvements = topWinners
      .slice(0, 3)
      .map(
        (k) =>
          `'${k.keyword}' climbed from position ${k.previous ?? "-"} to ${
            k.current ?? "-"
          }`
      )
      .join("; ");
    krLines.push(`Notable improvements: ${improvements}.`);
  } else {
    krLines.push("Notable improvements: no major positive movements this period.");
  }

  if (topLosers.length) {
    const focus = topLosers
      .slice(0, 3)
      .map(
        (k) =>
          `'${k.keyword}' dropped from ${k.previous ?? "-"} to ${
            k.current ?? "-"
          }`
      )
      .join("; ");
    krLines.push(`Areas for focus: ${focus}.`);
  } else {
    krLines.push("Areas for focus: no significant ranking drops recorded.");
  }

  krLines.forEach((line) => {
    doc.text("• " + line, { width: contentWidth });
  });

  // Backlinks summary as a clearly separated section
  if (backlinks && backlinks.totalBacklinks) {
    doc.moveDown(1.2);
    doc.fontSize(12).fillColor("#000000").text("Backlinks and Authority:");
    doc.moveDown(0.4);
    doc.fontSize(11).fillColor("#333333");

    const parts = [];
    if (backlinks.sections && backlinks.sections.length) {
      backlinks.sections.forEach((s) => {
        parts.push(`${s.name} (${s.total})`);
      });
    }

    const line =
      parts.length > 0
        ? `Links recorded in this workbook: ${backlinks.totalBacklinks}. Key backlink categories: ${parts.join(
            ", "
          )}.`
        : `Links recorded in this workbook: ${backlinks.totalBacklinks}.`;

    doc.text(line, { width: contentWidth });
  }

  // NOTE: footer/logos removed on purpose to avoid blank extra page

  // ======================================================
  // KEYWORD LIST TABLE
  // ======================================================
  doc.addPage();

  doc
    .fontSize(16)
    .fillColor("#000000")
    .text("Keyword Rankings Overview", left, doc.y, {
      underline: true
    });

  doc.moveDown(0.7);

  const xCheck = left;
  const xKeyword = left + 18;
  const xCurrent = left + 310;
  const xPrev = left + 420;

  function drawKeywordHeader() {
    const headerY = doc.y;
    doc.fontSize(11).fillColor("#000000");
    doc.text("Keyword List", xKeyword, headerY);
    doc.text("Position", xCurrent, headerY);
    doc.text("Last week", xPrev, headerY);
    doc.moveDown(0.5);
  }

  drawKeywordHeader();

  let y = doc.y + 2;
  const rowHeight = 18;

  doc.fontSize(10).fillColor("#333333");

  keywords.forEach((k) => {
    if (y > doc.page.height - doc.page.margins.bottom - 30) {
      doc.addPage();
      drawKeywordHeader();
      y = doc.y + 2;
      doc.fontSize(10).fillColor("#333333");
    }

    doc.fontSize(11).fillColor(brandGreen).text("✓", xCheck, y + 2);

    doc.fontSize(10).fillColor("#333333");
    doc.text(k.keyword, xKeyword, y, {
      width: xCurrent - xKeyword - 10
    });

    const currentStr =
      k.current !== null && k.current !== undefined
        ? String(k.current)
        : "-";
    doc.text(currentStr, xCurrent, y, { width: 60 });

    const prevStr =
      k.previous !== null && k.previous !== undefined
        ? String(k.previous)
        : "-";
    doc.text(prevStr, xPrev, y, { width: 60 });

    y += rowHeight;
  });

  // ======================================================
  // BACKLINKS OVERVIEW + DETAILS (compact)
  // ======================================================
  if (backlinks && backlinks.sections && backlinks.sections.length) {
    doc.addPage();
    doc
      .fontSize(16)
      .fillColor("#000000")
      .text("Backlinks Overview", left, doc.y, { underline: true });
    doc.moveDown(0.7);

    doc
      .fontSize(11)
      .fillColor("#333333")
      .text(
        `Total backlinks in this workbook: ${backlinks.totalBacklinks}.`,
        { width: contentWidth }
      );

    if (backlinks.sections.length) {
      const summaryParts = backlinks.sections.map(
        (s) => `${s.name} (${s.total})`
      );
      doc.moveDown(0.3);
      doc.text(`By type: ${summaryParts.join(", ")}.`, {
        width: contentWidth
      });
    }

    doc.moveDown(1.0);

    // Now list each category, sharing pages, up to 5 links each
    backlinks.sections.forEach((section, idx) => {
      // space check
      if (doc.y > doc.page.height - doc.page.margins.bottom - 80) {
        doc.addPage();
      }

      doc
        .fontSize(13)
        .fillColor("#000000")
        .text(section.name, left, doc.y, { underline: true });
      doc.moveDown(0.3);

      const maxLinks = 5;
      const shown = Math.min(section.rows.length, maxLinks);

      doc
        .fontSize(10)
        .fillColor("#333333")
        .text(
          `Total links: ${section.total}. Showing first ${shown} links:`,
          { width: contentWidth }
        );
      doc.moveDown(0.3);

      doc.fontSize(9).fillColor("#000000");
      section.rows.slice(0, maxLinks).forEach((row, i) => {
        const url = row.backlink || row.target || "";
        const statusText = row.status ? ` [${row.status}]` : "";
        doc.text(`${i + 1}. ${url}${statusText}`, { width: contentWidth });
      });

      if (section.rows.length > maxLinks) {
        doc.moveDown(0.2);
        doc
          .fontSize(9)
          .fillColor("#555555")
          .text(
            `... plus ${section.rows.length - maxLinks} more links in this category.`,
            { width: contentWidth }
          );
      }

      doc.moveDown(0.4);
      doc
        .fontSize(9)
        .fillColor("#444444")
        .text(
          "What this means: " + describeBacklinkType(section.name),
          { width: contentWidth }
        );

      if (idx < backlinks.sections.length - 1) {
        doc.moveDown(0.9);
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
