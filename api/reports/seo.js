// api/reports/seo.js
//
// Accepts multipart/form-data with field "seo_file" (your weekly Excel/CSV)
// Reads Ranking + Backlinks sheets and returns a branded PDF.

import formidable from "formidable";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";

export const config = {
  api: {
    bodyParser: false, // we handle multipart ourselves
  },
};

function parseForm(req) {
  return new Promise((resolve, reject) => {
    const form = formidable({ multiples: false });
    form.parse(req, (err, fields, files) => {
      if (err) return reject(err);
      resolve({ fields, files });
    });
  });
}

async function parseSeoWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const rankingSheet = workbook.getWorksheet("Ranking");
  if (!rankingSheet) throw new Error("Missing 'Ranking' sheet");

  // --- Meta (domain/location) from first rows ---
  const domain = rankingSheet.getCell("A3").value || "";
  const location = rankingSheet.getCell("A4").value || "";

  // --- Dates row (C3 onwards) ---
  const dateRow = rankingSheet.getRow(3);
  const dates = [];
  dateRow.eachCell((cell, col) => {
    if (col >= 3 && cell.value) {
      const v = cell.value;
      const dt =
        v instanceof Date
          ? v
          : typeof v === "string"
          ? new Date(v)
          : v.text
          ? new Date(v.text)
          : null;
      if (dt) dates.push({ col, date: dt });
    }
  });
  if (!dates.length) throw new Error("No dates in Ranking sheet");

  // latest + previous
  dates.sort((a, b) => a.date - b.date);
  const latest = dates[dates.length - 1];
  const previous = dates[dates.length - 2]; // assumes at least 2 dates

  // --- Keywords rows ---
  const keywords = [];
  rankingSheet.eachRow((row, rowNumber) => {
    if (rowNumber < 5) return; // skip header rows

    const keyword = row.getCell(1).value || "";
    const url = row.getCell(2).value || "";
    if (!keyword) return;

    const cur = row.getCell(latest.col).value || 0;
    const prev = previous ? row.getCell(previous.col).value || 0 : 0;

    const curPos = typeof cur === "number" ? cur : Number(cur) || null;
    const prevPos = typeof prev === "number" ? prev : Number(prev) || null;

    keywords.push({
      keyword: String(keyword),
      url: String(url || ""),
      current: curPos,
      previous: prevPos,
    });
  });

  const tracked = keywords.length;
  const withCurrent = keywords.filter((k) => k.current && k.current > 0);

  const avgCurrent =
    withCurrent.reduce((sum, k) => sum + k.current, 0) /
    (withCurrent.length || 1);

  const prevWith = keywords.filter((k) => k.previous && k.previous > 0);
  const avgPrev =
    prevWith.reduce((sum, k) => sum + k.previous, 0) /
    (prevWith.length || 1);

  const top10 = withCurrent.filter((k) => k.current <= 10).length;
  const top10Prev = prevWith.filter((k) => k.previous <= 10).length;

  const winners = keywords.filter(
    (k) => k.current && k.previous && k.current < k.previous
  );
  const losers = keywords.filter(
    (k) => k.current && k.previous && k.current > k.previous
  );

  // sort top movers
  const movers = keywords
    .map((k) => ({
      ...k,
      change:
        k.current && k.previous ? k.previous - k.current : null, // +ve = improvement
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

  // --- Backlinks summary from All Backlinks sheet (if present) ---
  const allBacklinks = workbook.getWorksheet("All Backlinks");
  let backlinkSummary = null;
  if (allBacklinks) {
    // you can adapt this to exactly match how your sheet is structured
    // For example assume:
    // A2: "Profile backlinks", B2: 5 etc
    const rows = [];
    allBacklinks.eachRow((row, rowNumber) => {
      if (rowNumber < 2) return;
      const type = row.getCell(1).value;
      const count = row.getCell(2).value;
      if (type && count) {
        rows.push({
          type: String(type),
          count: Number(count),
        });
      }
    });
    backlinkSummary = rows;
  }

  return {
    domain: String(domain),
    location: String(location),
    dates,
    latest,
    previous,
    tracked,
    avgCurrent,
    avgPrev,
    top10,
    top10Prev,
    winners,
    losers,
    topWinners,
    topLosers,
    keywords,
    backlinkSummary,
  };
}

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
    topWinners,
    topLosers,
    keywords,
    backlinkSummary,
  } = summary;

  const doc = new PDFDocument({ margin: 40 });

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="SEO-Report-${domain || "site"}.pdf"`
  );

  doc.pipe(res);

  const fmtDate = (d) =>
    d ? d.date.toISOString().slice(0, 10) : "";

  // --- COVER / HEADER ---
  doc.fontSize(18).text("SEO Weekly Report", { bold: true });
  doc.moveDown(0.5);
  doc.fontSize(12).fillColor("#555");
  doc.text(`Domain: ${domain || "-"}`);
  doc.text(`Location: ${location || "-"}`);
  doc.text(`Week ending: ${fmtDate(latest)}`);
  if (previous) doc.text(`Compared with: ${fmtDate(previous)}`);
  doc.moveDown();

  // KPI summary
  doc.fontSize(12).fillColor("#000");
  doc.text(`Tracked keywords: ${tracked}`);
  doc.text(
    `Average position: ${avgCurrent.toFixed(1)} (prev ${avgPrev.toFixed(1)})`
  );
  const top10Pct =
    tracked > 0 ? Math.round((top10 / tracked) * 100) : 0;
  const top10PrevPct =
    tracked > 0 ? Math.round((top10Prev / tracked) * 100) : 0;
  doc.text(
    `Top 10 visibility: ${top10Pct}% (prev ${top10PrevPct}%)`
  );
  doc.moveDown();

  // Small summary sentence
  doc.fontSize(11).fillColor("#555");
  doc.text(
    `This week we are tracking ${tracked} keywords. ` +
      `Average ranking changed from ${avgPrev.toFixed(
        1
      )} to ${avgCurrent.toFixed(1)}, ` +
      `with ${top10Pct}% of keywords in the top 10.`
  );

  doc.addPage();

  // --- TOP WINNERS / LOSERS ---
  doc.fontSize(14).fillColor("#000").text("Top Winners", { underline: true });
  doc.moveDown(0.5);
  doc.fontSize(10);

  if (!topWinners.length) {
    doc.text("No improving keywords this period.");
  } else {
    topWinners.forEach((k) => {
      doc.text(
        `${k.keyword} — ${k.previous} → ${k.current} (↑ ${k.change})`
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
        `${k.keyword} — ${k.previous} → ${k.current} (${k.change})`
      );
    });
  }

  doc.addPage();

  // --- FULL KEYWORD TABLE (simple text version) ---
  doc.fontSize(14).fillColor("#000").text("Keyword detail", { underline: true });
  doc.moveDown(0.5);
  doc.fontSize(9);
  doc.text("Keyword                          Prev   Curr   Change");
  doc.text("------------------------------------------------------");

  keywords.forEach((k) => {
    const prevStr = k.previous ? String(k.previous).padStart(4) : "   -";
    const currStr = k.current ? String(k.current).padStart(4) : "   -";
    let changeStr = "   -";
    if (k.current && k.previous) {
      const diff = k.previous - k.current;
      if (diff > 0) changeStr = ` ↑${String(diff).padStart(2)}`;
      else if (diff < 0) changeStr = ` ↓${String(Math.abs(diff)).padStart(2)}`;
      else changeStr = "  0";
    }
    const kw = k.keyword.length > 30 ? k.keyword.slice(0,27) + "..." : k.keyword;
    doc.text(
      `${kw.padEnd(30)} ${prevStr}  ${currStr}  ${changeStr}`
    );
  });

  if (backlinkSummary && backlinkSummary.length) {
    doc.addPage();
    doc.fontSize(14).fillColor("#000").text("Backlinks summary", {
      underline: true,
    });
    doc.moveDown(0.5);
    doc.fontSize(10);
    backlinkSummary.forEach((row) => {
      doc.text(`${row.type}: ${row.count}`);
    });
  }

  doc.end();
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", ["POST"]);
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { files } = await parseForm(req);
    const file = files.seo_file;
    if (!file) {
      return res.status(400).json({ error: "Missing seo_file upload" });
    }

    const filePath = Array.isArray(file) ? file[0].filepath : file.filepath;

    const summary = await parseSeoWorkbook(filePath);

    // Stream PDF back to the client
    buildSeoPdf(res, summary);
  } catch (err) {
    console.error("SEO report error:", err);
    if (!res.headersSent) {
      res.status(500).json({
        error: "Failed to generate SEO report",
        detail: err.message,
      });
    }
  }
}
