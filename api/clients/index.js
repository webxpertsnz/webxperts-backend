import { getPool } from "../_lib/db.js";

export default async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = getPool();

  try {
    if (req.method === "GET") {
      const [rows] = await db.query("SELECT * FROM clients ORDER BY id DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const { name, company, contact_name, email, phone, address1, city, status="active", website, notes } = req.body || {};
      await db.query(
        `INSERT INTO clients (name, company, contact_name, email, phone, address1, city, status, website, notes, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [name ?? company ?? "", company ?? "", contact_name ?? "", email ?? "", phone ?? "", address1 ?? "", city ?? "", status, website ?? "", notes ?? ""]
      );
      return res.status(201).json({ message: "Client added" });
    }

    if (req.method === "PUT") {
      const { id, name, company, contact_name, email, phone, address1, city, status, website, notes } = req.body || {};
      await db.query(
        `UPDATE clients
         SET name=?, company=?, contact_name=?, email=?, phone=?, address1=?, city=?, status=?, website=?, notes=?, updated_at=NOW()
         WHERE id=?`,
        [name ?? company ?? "", company ?? "", contact_name ?? "", email ?? "", phone ?? "", address1 ?? "", city ?? "", status ?? "active", website ?? "", notes ?? "", id]
      );
      return res.status(200).json({ message: "Client updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body || {};
      await db.query("DELETE FROM clients WHERE id=?", [id]);
      return res.status(200).json({ message: "Client deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Clients API Error:", err);
    return res.status(500).json({ error: "Internal Server Error", detail: err.code || err.message });
  }
}
