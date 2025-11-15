// api/calendar/index.js
// Calendar API built like your clients API

import mysql from 'mysql2/promise';

// IMPORTANT:
// Copy your existing DB config from api/clients/index.js.
// For example, if clients uses createPool({ host, user, ... }),
// use EXACTLY the same here.

const pool = mysql.createPool({
  host: process.env.DB_HOST,      // <- make this match your clients API
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
  waitForConnections: true,
  connectionLimit: 10,
});

export default async function handler(req, res) {
  try {
    if (req.method === 'GET') {
      // GET /api/calendar?month=YYYY-MM
      const { month } = req.query;

      if (!month || !/^\d{4}-\d{2}$/.test(month)) {
        return res.status(400).json({ ok: false, error: 'Invalid month' });
      }

      const [yearStr, monStr] = month.split('-');
      const year = parseInt(yearStr, 10);
      const mon = parseInt(monStr, 10);
      if (mon < 1 || mon > 12) {
        return res.status(400).json({ ok: false, error: 'Invalid month' });
      }

      const startDate = `${year}-${String(mon).padStart(2, '0')}-01`;
      const endDate = new Date(year, mon, 1); // JS month is 0-based
      const nextMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
      nextMonth.setMonth(nextMonth.getMonth() + 1);
      const endStr = `${nextMonth.getFullYear()}-${String(nextMonth.getMonth() + 1).padStart(2, '0')}-01`;

      const [rows] = await pool.query(
        `SELECT id, title, event_type, event_date, event_time, notes, created_at
         FROM calendar_events
         WHERE event_date >= ? AND event_date < ?
         ORDER BY event_date ASC, event_time ASC`,
        [startDate, endStr]
      );

      const events = rows.map((row) => ({
        id: row.id,
        title: row.title,
        type: row.event_type,
        date: row.event_date,          // 'YYYY-MM-DD'
        time: row.event_time,          // 'HH:MM:SS' or null
        notes: row.notes,
        created_at: row.created_at,
      }));

      return res.status(200).json({ ok: true, events });
    }

    if (req.method === 'POST') {
      // POST /api/calendar  -> create
      const { title, type, date, time, notes } = req.body || {};

      if (!title || !date) {
        return res.status(400).json({ ok: false, error: 'Title and date required' });
      }

      const validTypes = ['meeting', 'call', 'followup', 'personal'];
      const eventType = validTypes.includes(type) ? type : 'meeting';

      let eventTime = null;
      if (time && time !== '') {
        // normalise "HH:MM" or "HH:MM:SS"
        const parts = String(time).split(':');
        const hh = parseInt(parts[0] || '0', 10);
        const mm = parseInt(parts[1] || '0', 10);
        const ss = parseInt(parts[2] || '0', 10);
        eventTime = `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}`;
      }

      const [result] = await pool.query(
        `INSERT INTO calendar_events
           (title, event_type, event_date, event_time, notes)
         VALUES (?, ?, ?, ?, ?)`,
        [title.trim(), eventType, date, eventTime, notes || null]
      );

      return res.status(200).json({
        ok: true,
        event: {
          id: result.insertId,
          title: title.trim(),
          type: eventType,
          date,
          time: eventTime,
          notes: notes || null,
        },
      });
    }

    if (req.method === 'PUT') {
      // PUT /api/calendar?id=123 -> update
      const { id } = req.query;
      const { title, type, date, time, notes } = req.body || {};

      if (!id) {
        return res.status(400).json({ ok: false, error: 'Missing id' });
      }
      if (!title || !date) {
        return res.status(400).json({ ok: false, error: 'Title and date required' });
      }

      const validTypes = ['meeting', 'call', 'followup', 'personal'];
      const eventType = validTypes.includes(type) ? type : 'meeting';

      let eventTime = null;
      if (time && time !== '') {
        const parts = String(time).split(':');
        const hh = parseInt(parts[0] || '0', 10);
        const mm = parseInt(parts[1] || '0', 10);
        const ss = parseInt(parts[2] || '0', 10);
        eventTime = `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}`;
      }

      await pool.query(
        `UPDATE calendar_events
         SET title = ?, event_type = ?, event_date = ?, event_time = ?, notes = ?
         WHERE id = ?`,
        [title.trim(), eventType, date, eventTime, notes || null, id]
      );

      return res.status(200).json({ ok: true });
    }

    if (req.method === 'DELETE') {
      // DELETE /api/calendar?id=123
      const { id } = req.query;
      if (!id) {
        return res.status(400).json({ ok: false, error: 'Missing id' });
      }

      await pool.query('DELETE FROM calendar_events WHERE id = ?', [id]);

      return res.status(200).json({ ok: true });
    }

    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  } catch (err) {
    console.error('Calendar API error', err);
    return res.status(500).json({ ok: false, error: 'Server error' });
  }
}
