import { getPool } from "../lib/db.js";

export default async function handler(req, res) {
  // --- CORS / preflight ---
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = getPool();

  try {
    // --------------------------------------------------
    // GET /api/renewals  -> list all renewals
    // --------------------------------------------------
    if (req.method === "GET") {
      const [rows] = await db.query(
        "SELECT * FROM renewals ORDER BY renewal_date ASC"
      );
      return res.status(200).json(rows);
    }

    // --------------------------------------------------
    // POST /api/renewals  -> add new renewal
    // body: { client_id?, service_type, item_label, renewal_date, cost?, auto_renew?, status? }
    // --------------------------------------------------
    if (req.method === "POST") {
      const {
        client_id,
        service_type,
        item_label,
        renewal_date,
        cost,
        auto_renew = 0,
        status = "active",
      } = req.body || {};

      // allow renewals with no client yet
      const client = client_id || null;

      await db.query(
        `INSERT INTO renewals
           (client_id, service_type, item_label, renewal_date, cost, auto_renew, status, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, NOW())`,
        [
          client,
          service_type ?? "other",
          item_label ?? "",
          renewal_date ?? null,
          cost ?? 0,
          auto_renew ? 1 : 0,
          status,
        ]
      );

      return res.status(201).json({ message: "Renewal added" });
    }

    // --------------------------------------------------
    // PUT /api/renewals  -> update existing renewal
    // body: { id, client_id?, service_type?, item_label?, renewal_date?, cost?, auto_renew?, status? }
    // --------------------------------------------------
    if (req.method === "PUT") {
      const {
        id,
        client_id,
        service_type,
        item_label,
        renewal_date,
        cost,
        auto_renew,
        status,
      } = req.body || {};

      if (!id) {
        return res.status(400).json({ error: "Missing id" });
      }

      const client = client_id || null;

      await db.query(
        `UPDATE renewals
           SET client_id   = ?,
               service_type = ?,
               item_label   = ?,
               renewal_date = ?,
               cost         = ?,
               auto_renew   = ?,
               status       = ?,
               updated_at   = NOW()
         WHERE id = ?`,
        [
          client,
          service_type ?? "other",
          item_label ?? "",
          renewal_date ?? null,
          cost ?? 0,
          auto_renew ? 1 : 0,
          status ?? "active",
          id,
        ]
      );

      return res.status(200).json({ message: "Renewal updated" });
    }

    // --------------------------------------------------
    // DELETE /api/renewals  -> delete renewal
    // body: { id }
    // --------------------------------------------------
    if (req.method === "DELETE") {
      const { id } = req.body || {};
      if (!id) {
        return res.status(400).json({ error: "Missing id" });
      }
      await db.query("DELETE FROM renewals WHERE id = ?", [id]);
      return res.status(200).json({ message: "Renewal deleted" });
    }

    // --------------------------------------------------
    // Method not allowed
    // --------------------------------------------------
    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE", "OPTIONS"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Renewals API Error:", err);
    return res
      .status(500)
      .json({ error: "Internal Server Error", detail: err.code || err.message });
  }
}
