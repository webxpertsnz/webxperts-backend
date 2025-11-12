import { getPool } from "../lib/db.js";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = getPool();

  try {
    // GET: list all renewals
    if (req.method === "GET") {
      const [rows] = await db.query(
        "SELECT * FROM renewals ORDER BY renewal_date DESC"
      );
      return res.status(200).json(rows);
    }

    // POST: create renewal
    if (req.method === "POST") {
      const {
        client_id,
        service_type,
        item_label,
        renewal_date,
        cost,
        auto_renew = 0,
        status = "active",
        provider,      // NEW
      } = req.body || {};

      await db.query(
        `INSERT INTO renewals
         (client_id, service_type, item_label, provider, renewal_date, cost, auto_renew, status, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [
          client_id,
          service_type ?? "other",
          item_label ?? "",
          provider ?? null,
          renewal_date ?? null,
          cost ?? 0,
          auto_renew ? 1 : 0,
          status,
        ]
      );
      return res.status(201).json({ message: "Renewal added" });
    }

    // PUT: update renewal
    if (req.method === "PUT") {
      const {
        id,
        renewal_date,
        cost,
        auto_renew,
        status,
        provider,
      } = req.body || {};

      await db.query(
        `UPDATE renewals
         SET renewal_date = ?,
             cost         = ?,
             auto_renew   = ?,
             status       = ?,
             provider     = ?,
             updated_at   = NOW()
         WHERE id = ?`,
        [
          renewal_date ?? null,
          cost ?? 0,
          auto_renew ? 1 : 0,
          status ?? "active",
          provider ?? null,
          id,
        ]
      );
      return res.status(200).json({ message: "Renewal updated" });
    }

    // DELETE: delete renewal
    if (req.method === "DELETE") {
      const { id } = req.body || {};
      await db.query("DELETE FROM renewals WHERE id = ?", [id]);
      return res.status(200).json({ message: "Renewal deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Renewals API Error:", err);
    return res
      .status(500)
      .json({ error: "Internal Server Error", detail: err.code || err.message });
  }
}
