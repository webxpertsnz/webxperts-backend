// api/reports/seo.js

const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const PDFDocument = require('pdfkit');

const router = express.Router();

// Multer config – keep file in memory
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 5 * 1024 * 1024, // 5MB
  },
});

// POST /api/reports/seo
router.post('/seo', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .json({ success: false, message: 'No file uploaded' });
    }

    // Parse CSV
    let records;
    try {
      records = parse(req.file.buffer.toString('utf-8'), {
        columns: true,
        skip_empty_lines: true,
        trim: true,
      });
    } catch (err) {
      console.error('CSV parse error:', err);
      return res
        .status(400)
        .json({ success: false, message: 'Invalid CSV format' });
    }

    if (!records || records.length === 0) {
      return res
        .status(400)
        .json({ success: false, message: 'CSV contains no data rows' });
    }

    // Group rows by "Tab" column if it exists
    const hasTabColumn = Object.prototype.hasOwnProperty.call(
      records[0],
      'Tab'
    );

    const tabMap = new Map();

    if (hasTabColumn) {
      for (const row of records) {
        const tabName = row.Tab && row.Tab.trim() ? row.Tab.trim() : 'Data';
        if (!tabMap.has(tabName)) tabMap.set(tabName, []);
        tabMap.get(tabName).push(row);
      }
    } else {
      // Fallback: everything in a single tab
      tabMap.set('SEO Data', records);
    }

    // Start PDF
    const doc = new PDFDocument({ margin: 40 });

    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => {
      const pdfBuffer = Buffer.concat(chunks);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader(
        'Content-Disposition',
        'attachment; filename="seo-report.pdf"'
      );
      res.send(pdfBuffer);
    });

    // Basic header info from first row (if exists in your CSV)
    const firstRow = records[0];
    const clientName =
      firstRow.Client ||
      firstRow['Client Name'] ||
      firstRow['Client'] ||
      '';
    const website =
      firstRow.Website || firstRow.Domain || firstRow['Website URL'] || '';
    const month =
      firstRow.Month || firstRow['Report Month'] || firstRow['Period'] || '';

    // Cover page
    doc.fontSize(22).text('SEO Performance Report', { align: 'center' });
    doc.moveDown();

    if (clientName) {
      doc.fontSize(14).text(`Client: ${clientName}`);
    }
    if (website) {
      doc.fontSize(14).text(`Website: ${website}`);
    }
    if (month) {
      doc.fontSize(14).text(`Reporting Period: ${month}`);
    }

    doc.moveDown(2);
    doc.fontSize(10).text(
      'This report has been generated automatically from your uploaded CSV data.',
      { align: 'left' }
    );

    // Each tab = its own section
    let isFirstTab = true;

    for (const [tabName, rows] of tabMap.entries()) {
      if (!isFirstTab) {
        doc.addPage();
      } else {
        doc.addPage(); // start first tab on a fresh page after cover
        isFirstTab = false;
      }

      doc.fontSize(18).text(tabName, { underline: true });
      doc.moveDown();

      if (!rows || rows.length === 0) {
        doc.fontSize(10).text('No data available for this section.');
        continue;
      }

      // Determine columns (exclude the Tab column itself)
      const allKeys = Object.keys(rows[0]);
      const columns = hasTabColumn
        ? allKeys.filter((k) => k !== 'Tab')
        : allKeys;

      // Table header
      doc.fontSize(11).text(columns.join('  |  '));
      doc.moveDown(0.5);

      // Divider
      doc
        .moveTo(doc.page.margins.left, doc.y)
        .lineTo(doc.page.width - doc.page.margins.right, doc.y)
        .stroke();
      doc.moveDown(0.5);

      // Table rows
      doc.fontSize(9);
      for (const row of rows) {
        const values = columns.map((key) => (row[key] ?? '').toString());
        doc.text(values.join('  |  '));

        // Simple pagination safety
        if (doc.y > doc.page.height - doc.page.margins.bottom - 50) {
          doc.addPage();
          doc.fontSize(11).text(columns.join('  |  '));
          doc.moveDown(0.5);
          doc
            .moveTo(doc.page.margins.left, doc.y)
            .lineTo(doc.page.width - doc.page.margins.right, doc.y)
            .stroke();
          doc.moveDown(0.5);
          doc.fontSize(9);
        }
      }
    }

    doc.end();
  } catch (err) {
    console.error('Unexpected error in SEO report API:', err);
    return res.status(500).json({
      success: false,
      message: 'Server error while generating SEO report',
    });
  }
});

module.exports = router;

/*
  HOW TO USE / MOUNT THIS ROUTER:

  In your main Express app (e.g. server.js or index.js):

  const express = require('express');
  const seoReportsRouter = require('./api/reports/seo');

  const app = express();
  app.use('/api/reports', seoReportsRouter);

  // Then your frontend should POST to: /api/reports/seo
  // with multipart/form-data where the file field name is "file".
*/
