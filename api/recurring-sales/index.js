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
      if (req.query.client_id) {
        // New: Support GET by client_id for frontend
        const [rows] = await db.query(
          "SELECT * FROM recurring_sales WHERE client_id = ? ORDER BY id DESC",
          [req.query.client_id]
        );
        return res.status(200).json(rows);
      } else {
        // Original: Global GET
        const [rows] = await db.query("SELECT * FROM recurring_sales ORDER BY id DESC");
        return res.status(200).json(rows);
      }
    }

    if (req.method === "POST") {
      const {
        client_id, product, service_name, amount, description,
        quantity, unit_amount, start_date, end_date, notes
      } = req.body || {};

      await db.query(
        `INSERT INTO recurring_sales
         (client_id, product, service_name, amount, description, quantity, unit_amount, start_date, end_date, notes, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
        [client_id, product ?? "", service_name ?? "", amount ?? 0, description ?? "", quantity ?? 1, unit_amount ?? 0, start_date ?? null, end_date ?? null, notes ?? null]
      );
      return res.status(201).json({ message: "Recurring sale created" });
    }

    if (req.method === "PUT") {
      const body = req.body || {};
      const id = body.id;  // FIXED: Get id from body (frontend sends it there)

      // Log for debugging (check Vercel logs if issues)
      console.log('PUT recurring-sales body:', body);
      if (!id) {
        return res.status(400).json({ error: 'ID required in body' });
      }

      // ONLY allow these fields to prevent ER_BAD_FIELD_ERROR
      const allowedFields = ['description', 'service_name', 'amount', 'start_date', 'end_date', 'product', 'quantity', 'unit_amount', 'notes'];
      const updateData = {};
      for (const [key, value] of Object.entries(body)) {
        if (allowedFields.includes(key)) {
          // FIX: Convert '' or undefined to null (esp for dates/notes)
          updateData[key] = (value === '' || value === undefined || value === null) ? null : value;
        }
      }

      // Build safe SQL
      let sql = 'UPDATE recurring_sales SET ';
      const updates = [];
      const values = [];
      for (const [key, value] of Object.entries(updateData)) {
        updates.push(`${key} = ?`);
        values.push(value);
      }
      if (updates.length === 0) {
        return res.status(400).json({ error: 'No valid fields to update' });
      }
      sql += updates.join(', ') + ' WHERE id = ?';
      values.push(id);

      console.log('SQL:', sql); // Log SQL
      console.log('Values:', values); // Log values (now null-safe)

      const [result] = await db.execute(sql, values);
      if (result.affectedRows === 0) {
        return res.status(404).json({ error: 'Record not found' });
      }

      // Return the updated record
      const [rows] = await db.query('SELECT * FROM recurring_sales WHERE id = ?', [id]);
      return res.status(200).json(rows[0]);
    }

    if (req.method === "DELETE") {
      const { id } = req.body || {};
      await db.query("DELETE FROM recurring_sales WHERE id=?", [id]);
      return res.status(200).json({ message: "Recurring sale deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Recurring Sales API Error:", err);
    return res.status(500).json({ error: "Internal Server Error", detail: err.code || err.message });
  }
}
