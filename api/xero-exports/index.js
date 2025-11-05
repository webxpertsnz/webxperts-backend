// /api/xero-exports/index.js
import mysql from "mysql2/promise";

export default async function handler(req, res) {
  const db = await mysql.createConnection({
    host: "db.webxperts.co.nz",
    user: "u517327732_db_user",
    // password: "77Breebbnz#",
    database: "u517327732_db",
    port: 3306
  });

  try {
    if (req.method === "GET") {
      const [rows] = await db.query(`
        SELECT xe.*, COUNT(xt.id) AS item_count
        FROM xero_exports xe
        LEFT JOIN xero_export_items xt ON xe.id = xt.export_id
        GROUP BY xe.id
        ORDER BY xe.created_at DESC
      `);
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const { name, period_year, period_month } = req.body;
      await db.query(
        `INSERT INTO xero_exports (name, period_year, period_month, created_at)
         VALUES (?, ?, ?, NOW())`,
        [name, period_year, period_month]
      );
      return res.status(201).json({ message: "Export created" });
    }

    if (req.method === "PUT") {
      const { id, name, period_year, period_month } = req.body;
      await db.query(
        `UPDATE xero_exports
         SET name=?, period_year=?, period_month=?, updated_at=NOW()
         WHERE id=?`,
        [name, period_year, period_month, id]
      );
      return res.status(200).json({ message: "Export updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body;
      await db.query("DELETE FROM xero_exports WHERE id=?", [id]);
      await db.query("DELETE FROM xero_export_items WHERE export_id=?", [id]);
      return res.status(200).json({ message: "Export and related items deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Xero Exports API Error:", err);
    res.status(500).json({ error: "Internal Server Error" });
  } finally {
    await db.end();
  }
}
