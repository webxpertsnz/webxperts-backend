// api/calendar/index.js
import { getPool } from "../lib/db.js";

export default async function handler(req, res) {
  // CORS (same as clients)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = getPool();

  try {
    // GET /api/calendar?month=YYYY-MM
    if (req.method === "GET") {
      const { month } = req.query || {};

      if (!month || !/^\d{4}-\d{2}$/.test(month)) {
        return res
          .status(400)
          .json({ ok: false, error: "Invalid or missing month (YYYY-MM)" });
      }

      const [yearStr, monStr] = month.split("-");
      const year = parseInt(yearStr, 10);
      const mon = parseInt(monStr, 10);

      if (isNaN(year) || isNaN(mon) || mon < 1 || mon > 12) {
        return res
          .status(400)
          .json({ ok: false, error: "Invalid month value" });
      }

      const startDate = `${year}-${String(mon).padStart(2, "0")}-01`;
      const nextMonth = mon === 12 ? 1 : mon + 1;
      const nextYear = mon === 12 ? year + 1 : year;
      const endDate = `${nextYear}-${String(nextMonth).padStart(
        2,
        "0"
      )}-01`;

      const [rows] = await db.query(
        `SELECT id, title, event_type, event_date, event_time, notes, created_at
         FROM calendar_events
         WHERE event_date >= ? AND event_date < ?
         ORDER BY event_date ASC, event_time ASC`,
        [startDate, endDate]
      );

      const events = rows.map((row) => ({
        id: row.id,
        title: row.title,
        type: row.event_type,
        date: row.event_date
          ? row.event_date.toISOString
            ? row.event_date.toISOString().slice(0, 10)
            : String(row.event_date).slice(0, 10)
          : null,
        time: row.event_time
          ? row.event_time.toISOString
            ? row.event_time.toISOString().slice(11, 19)
            : String(row.event_time).slice(0, 8)
          : null,
        notes: row.notes ?? null,
        created_at: row.created_at ?? null,
      }));

      return res.status(200).json({ ok: true, events });
    }

    // POST /api/calendar  (create)
    if (req.method === "POST") {
      const { title, type, date, time, notes } = req.body || {};

      if (!title || !date) {
        return res
          .status(400)
          .json({ ok: false, error: "Title and date required" });
      }

      const validTypes = ["meeting", "call", "followup", "personal"];
      const eventType = validTypes.includes(type) ? type : "meeting";

      let eventTime = null;
      if (time) {
        const parts = String(time).split(":");
        const hh = parseInt(parts[0] || "0", 10);
        const mm = parseInt(parts[1] || "0", 10);
        const ss = parseInt(parts[2] || "0", 10);
        eventTime = `${String(hh).padStart(2, "0")}:${String(mm).padStart(
          2,
          "0"
        )}:${String(ss).padStart(2, "0")}`;
      }

      const [result] = await db.query(
        `INSERT INTO calendar_events (title, event_type, event_date, event_time, notes)
         VALUES (?, ?, ?, ?, ?)`,
        [title.trim(), eventType, date, eventTime, notes ?? null]
      );

      return res.status(201).json({
        ok: true,
        event: {
          id: result.insertId,
          title: title.trim(),
          type: eventType,
          date,
          time: eventTime,
          notes: notes ?? null,
        },
      });
    }

    // PUT /api/calendar  (update)
    if (req.method === "PUT") {
      const { id, title, type, date, time, notes } = req.body || {};

      if (!id) {
        return res.status(400).json({ ok: false, error: "Missing id" });
      }
      if (!title || !date) {
        return res
          .status(400)
          .json({ ok: false, error: "Title and date required" });
      }

      const validTypes = ["meeting", "call", "followup", "personal"];
      const eventType = validTypes.includes(type) ? type : "meeting";

      let eventTime = null;
      if (time) {
        const parts = String(time).split(":");
        const hh = parseInt(parts[0] || "0", 10);
        const mm = parseInt(parts[1] || "0", 10);
        const ss = parseInt(parts[2] || "0", 10);
        eventTime = `${String(hh).padStart(2, "0")}:${String(mm).padStart(
          2,
          "0"
        )}:${String(ss).padStart(2, "0")}`;
      }

      await db.query(
        `UPDATE calendar_events
           SET title = ?, event_type = ?, event_date = ?, event_time = ?, notes = ?
         WHERE id = ?`,
        [title.trim(), eventType, date, eventTime, notes ?? null, id]
      );

      return res.status(200).json({ ok: true, message: "Event updated" });
    }

    // DELETE /api/calendar  (delete)
    if (req.method === "DELETE") {
      const { id } = req.body || {};
      if (!id) {
        return res.status(400).json({ ok: false, error: "Missing id" });
      }

      await db.query("DELETE FROM calendar_events WHERE id=?", [id]);

      return res.status(200).json({ ok: true, message: "Event deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Calendar API Error:", err);
    return res.status(500).json({
      ok: false,
      error: "Internal Server Error",
      detail: err.code || err.message,
    });
  }
}
