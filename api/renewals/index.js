// /api/renewals/index.js
import mysql from "mysql2/promise";

export default async function handler(req, res) {
const db = await mysql.createConnection({
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASS,
  database: process.env.DB_NAME,
  port: process.env.DB_PORT
});

  try {
    if (req.method === "GET") {
      const [rows] = await db.query("SELECT * FROM renewals ORDER BY renewal_date DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const {
        client_id,
        service_type,
        item_label,
        renewal_date,
        cost,
        auto_renew,
        status
      } = req.body;

      await db.query(
        `INSERT INTO renewals
         (client_id, service_type, item_label, renewal_date, cost, auto_renew, status, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, NOW())`,
        [client_id, service_type, item_label, renewal_date, cost, auto_renew, status]
      );
      return res.status(201).json({ message: "Renewal added" });
    }

    if (req.method === "PUT") {
      const { id, renewal_date, cost, auto_renew, status } = req.body;

      await db.query(
        `UPDATE renewals
         SET renewal_date=?, cost=?, auto_renew=?, status=?, updated_at=NOW()
         WHERE id=?`,
        [renewal_date, cost, auto_renew, status, id]
      );
      return res.status(200).json({ message: "Renewal updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body;
      await db.query("DELETE FROM renewals WHERE id=?", [id]);
      return res.status(200).json({ message: "Renewal deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Renewals API Error:", err);
    res.status(500).json({ error: "Internal Server Error" });
  } finally {
    await db.end();
  }
}
