// api/reports/seo.js
// Minimal test: return JSON on GET and a tiny PDF on POST.
// No Excel parsing yet – this is just to prove the pipeline works.

import PDFDocument from "pdfkit";

export default async function handler(req, res) {
  try {
    if (req.method === "GET") {
      // Sanity check endpoint
      return res
        .status(200)
        .json({ ok: true, message: "SEO reports API is alive (test mode)" });
    }

    if (req.method !== "POST") {
      res.setHeader("Allow", ["GET", "POST"]);
      return res.status(405).json({ error: "Method not allowed" });
    }

    // Ignore the uploaded file for now – just return a simple PDF
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="SEO-Report-test.pdf"'
    );

    const doc = new PDFDocument({ margin: 40 });
    doc.pipe(res);

    doc.fontSize(20).text("SEO Report (test)", { underline: true });
    doc.moveDown();
    doc.fontSize(12).text("If you see this PDF, the API + upload flow is working.");
    doc.moveDown();
    doc.text("Next step: plug the real Excel parsing back in.");

    doc.end();
  } catch (err) {
    console.error("SEO report test error:", err);
    if (!res.headersSent) {
      return res.status(500).json({
        error: "Failed in SEO test endpoint",
        detail: err.message || String(err),
      });
    }
  }
}
