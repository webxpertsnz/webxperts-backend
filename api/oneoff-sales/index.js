// /api/oneoff-sales/index.js
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
      const [rows] = await db.query("SELECT * FROM oneoff_sales ORDER BY id DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const {
        client_id,
        description,
        amount,
        quantity,
        unit_amount,
        status,
        sale_date,
        notes
      } = req.body;

      await db.query(
        `INSERT INTO oneoff_sales
         (client_id, description, amount, quantity, unit_amount, status, sale_date, notes, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [client_id, description, amount, quantity, unit_amount, status, sale_date, notes]
      );
      return res.status(201).json({ message: "Sale created" });
    }

    if (req.method === "PUT") {
      const {
        id,
        description,
        amount,
        quantity,
        unit_amount,
        status,
        sale_date,
        notes
      } = req.body;

      await db.query(
        `UPDATE oneoff_sales
         SET description=?, amount=?, quantity=?, unit_amount=?, status=?, sale_date=?, notes=?, updated_at=NOW()
         WHERE id=?`,
        [description, amount, quantity, unit_amount, status, sale_date, notes, id]
      );
      return res.status(200).json({ message: "Sale updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body;
      await db.query("DELETE FROM oneoff_sales WHERE id=?", [id]);
      return res.status(200).json({ message: "Sale deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("OneOff Sales API Error:", err);
    res.status(500).json({ error: "Internal Server Error" });
  } finally {
    await db.end();
  }
}
