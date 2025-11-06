import mysql from "mysql2/promise";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = await mysql.createConnection({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASS,
    database: process.env.DB_NAME,
    port: Number(process.env.DB_PORT || 3306),
    ssl: { rejectUnauthorized: false }
  });

  try {
    if (req.method === "GET") {
      const [rows] = await db.query("SELECT * FROM recurring_sales ORDER BY id DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const {
        client_id, product, service_name, amount, description,
        quantity, unit_amount, start_date, end_date, notes
      } = req.body;

      await db.query(
        `INSERT INTO recurring_sales
         (client_id, product, service_name, amount, description, quantity, unit_amount, start_date, end_date, notes, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [client_id, product ?? "", service_name ?? "", amount ?? 0, description ?? "", quantity ?? 1, unit_amount ?? 0, start_date ?? null, end_date ?? null, notes ?? null]
      );
      return res.status(201).json({ message: "Recurring sale created" });
    }

    if (req.method === "PUT") {
      const {
        id, product, service_name, amount, description,
        quantity, unit_amount, start_date, end_date, notes
      } = req.body;

      await db.query(
        `UPDATE recurring_sales
         SET product=?, service_name=?, amount=?, description=?, quantity=?, unit_amount=?, start_date=?, end_date=?, notes=?, updated_at=NOW()
         WHERE id=?`,
        [product ?? "", service_name ?? "", amount ?? 0, description ?? "", quantity ?? 1, unit_amount ?? 0, start_date ?? null, end_date ?? null, notes ?? null, id]
      );
      return res.status(200).json({ message: "Recurring sale updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body;
      await db.query("DELETE FROM recurring_sales WHERE id=?", [id]);
      return res.status(200).json({ message: "Recurring sale deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Recurring Sales API Error:", err);
    return res.status(500).json({ error: err.message || "Internal Server Error" });
  } finally {
    await db.end();
  }
}
