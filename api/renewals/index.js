// /api/renewals/index.js
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
