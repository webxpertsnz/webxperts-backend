// api/clients/index.js
import mysql from "mysql2/promise";

export default async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();

  let db;
  try {
    db = await mysql.createConnection({
      host: process.env.DB_HOST,
      user: process.env.DB_USER,
      password: process.env.DB_PASS,
      database: process.env.DB_NAME,
      port: Number(process.env.DB_PORT || 3306),
      ssl: { rejectUnauthorized: false }
    });

    if (req.method === "GET") {
      const [rows] = await db.query("SELECT * FROM clients ORDER BY id DESC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
      const {
        company = "",
        contact_name = "",
        email = "",
        phone = "",
        address = "", // from UI — map to address1
        city = null,
        status = "active",
        website = null,
        notes = null
      } = body;

      const [r] = await db.query(
        `INSERT INTO clients (name, company, contact_name, email, phone, address1, city, status, website, notes, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [company || "", company || "", contact_name, email, phone, address, city, status, website, notes]
      );

      return res.status(201).json({ id: r.insertId, message: "Client added" });
    }

    res.setHeader("Allow", "GET, POST");
    return res.status(405).end();
  } catch (err) {
    console.error("Clients index error:", err);
    return res.status(500).json({ error: "Internal Server Error", detail: err.code || err.message });
  } finally {
    if (db) try { await db.end(); } catch {}
  }
}
