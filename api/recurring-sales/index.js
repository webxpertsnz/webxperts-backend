// /api/recurring-sales/index.js
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
      const [rows] = await db.query("SELECT * FROM recurring_sales ORDER BY id DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const {
        client_id,
        product,
        service_name,
        amount,
        quantity,
        unit_amount,
        description,
        start_date,
        end_date,
        notes
      } = req.body;

      await db.query(
        `INSERT INTO recurring_sales
         (client_id, product, service_name, amount, quantity, unit_amount, description, start_date, end_date, notes, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [client_id, product, service_name, amount, quantity, unit_amount, description, start_date, end_date, notes]
      );
      return res.status(201).json({ message: "Recurring sale created" });
    }

    if (req.method === "PUT") {
      const {
        id,
        product,
        service_name,
        amount,
        quantity,
        unit_amount,
        description,
        start_date,
        end_date,
        notes
      } = req.body;

      await db.query(
        `UPDATE recurring_sales
         SET product=?, service_name=?, amount=?, quantity=?, unit_amount=?, description=?, start_date=?, end_date=?, notes=?, updated_at=NOW()
         WHERE id=?`,
        [product, service_name, amount, quantity, unit_amount, description, start_date, end_date, notes, id]
      );
      return res.status(200).json({ message: "Recurring sale updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body;
      await db.query("DELETE FROM recurring_sales WHERE id=?", [id]);
      return res.status(200).json({ message: "Recurring sale deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Recurring Sales API Error:", err);
    res.status(500).json({ error: "Internal Server Error" });
  } finally {
    await db.end();
  }
}
