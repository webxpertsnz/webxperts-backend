// /api/clients/index.js
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
      const [rows] = await db.query("SELECT * FROM clients ORDER BY id DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const { name, company, contact_name, email, phone, address1, city } = req.body;
      await db.query(
        `INSERT INTO clients (name, company, contact_name, email, phone, address1, city, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, NOW())`,
        [name, company, contact_name, email, phone, address1, city]
      );
      return res.status(201).json({ message: "Client added" });
    }

    if (req.method === "PUT") {
      const { id, name, company, contact_name, email, phone, address1, city } = req.body;
      await db.query(
        `UPDATE clients SET name=?, company=?, contact_name=?, email=?, phone=?, address1=?, city=?, updated_at=NOW()
         WHERE id=?`,
        [name, company, contact_name, email, phone, address1, city, id]
      );
      return res.status(200).json({ message: "Client updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body;
      await db.query("DELETE FROM clients WHERE id=?", [id]);
      return res.status(200).json({ message: "Client deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Clients API Error:", err);
    res.status(500).json({ error: "Internal Server Error" });
  } finally {
    await db.end();
  }
}
