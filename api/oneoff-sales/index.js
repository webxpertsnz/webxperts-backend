import { getPool } from "../lib/db.js";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = getPool();

  try {
    if (req.method === "GET") {
      const [rows] = await db.query("SELECT * FROM oneoff_sales ORDER BY id DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const { client_id, description, amount, quantity, unit_amount, status="draft", sale_date, notes } = req.body || {};
      await db.query(
        `INSERT INTO oneoff_sales
         (client_id, description, amount, quantity, unit_amount, status, sale_date, notes, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [client_id, description ?? "", amount ?? null, quantity ?? 1, unit_amount ?? 0, status, sale_date ?? null, notes ?? null]
      );
      return res.status(201).json({ message: "Sale created" });
    }

    if (req.method === "PUT") {
      const { id, description, amount, quantity, unit_amount, status, sale_date, notes } = req.body || {};
      await db.query(
        `UPDATE oneoff_sales
         SET description=?, amount=?, quantity=?, unit_amount=?, status=?, sale_date=?, notes=?, updated_at=NOW()
         WHERE id=?`,
        [description ?? "", amount ?? null, quantity ?? 1, unit_amount ?? 0, status ?? "draft", sale_date ?? null, notes ?? null, id]
      );
      return res.status(200).json({ message: "Sale updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body || {};
      await db.query("DELETE FROM oneoff_sales WHERE id=?", [id]);
      return res.status(200).json({ message: "Sale deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("OneOff Sales API Error:", err);
    return res.status(500).json({ error: "Internal Server Error", detail: err.code || err.message });
  }
}
