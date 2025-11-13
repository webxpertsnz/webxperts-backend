import { getPool } from "../lib/db.js";

export default async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = getPool();

  try {
    // GET – list all renewals
    if (req.method === "GET") {
      const [rows] = await db.query(
        "SELECT * FROM renewals ORDER BY renewal_date ASC, item_label ASC"
      );
      return res.status(200).json(rows);
    }

    // POST – create renewal
    if (req.method === "POST") {
      const {
        client_id,
        service_type,
        item_label,
        provider,
        renewal_date,
        cost,
        auto_renew = 0,
        status = "active",
      } = req.body || {};

      await db.query(
        `INSERT INTO renewals
         (client_id, service_type, item_label, provider, renewal_date, cost, auto_renew, status, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [
          client_id || null,
          service_type || "domain",
          item_label || "",
          provider || null,
          renewal_date || null,
          cost ?? 0,
          auto_renew ? 1 : 0,
          status,
        ]
      );

      return res.status(201).json({ message: "Renewal added" });
    }

    // PUT – update renewal (INCLUDING client_id + provider)
    if (req.method === "PUT") {
      const {
        id,
        client_id,
        service_type,
        item_label,
        provider,
        renewal_date,
        cost,
        auto_renew,
        status,
      } = req.body || {};

      if (!id) {
        return res.status(400).json({ error: "Missing id for update" });
      }

      await db.query(
        `UPDATE renewals
         SET client_id = ?,
             service_type = ?,
             item_label   = ?,
             provider     = ?,
             renewal_date = ?,
             cost         = ?,
             auto_renew   = ?,
             status       = ?,
             updated_at   = NOW()
         WHERE id = ?`,
        [
          client_id || null,
          service_type || "domain",
          item_label || "",
          provider || null,
          renewal_date || null,
          cost ?? 0,
          auto_renew ? 1 : 0,
          status || "active",
          id,
        ]
      );

      return res.status(200).json({ message: "Renewal updated" });
    }

    // DELETE – remove renewal
    if (req.method === "DELETE") {
      const { id } = req.body || {};
      if (!id) {
        return res.status(400).json({ error: "Missing id for delete" });
      }

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
