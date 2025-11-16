import { getPool } from "../lib/db.js";

const todayISO = (d = new Date()) => d.toISOString().slice(0, 10);

function toISODate(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  const s = String(value);
  return s.length >= 10 ? s.slice(0, 10) : s;
}

function toPlainDateTime(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return value.toISOString().slice(0, 19).replace("T", " ");
  }
  const s = String(value);
  return s.length >= 19 ? s.slice(0, 19).replace("T", " ") : s;
}

export default async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = getPool();

  try {
    /* ---------- GET: list projects ---------- */
    if (req.method === "GET") {
      const { status, month, client_id } = req.query || {};

      let sql = `
        SELECT p.*, c.company AS client_name
        FROM projects p
        LEFT JOIN clients c ON p.client_id = c.id
        WHERE 1=1
      `;
      const params = [];

      if (status === "active") {
        sql += " AND p.completed = 0";
      } else if (status === "completed") {
        sql += " AND p.completed = 1";
      }

      // filter by client id (for client details page)
      if (client_id) {
        sql += " AND p.client_id = ?";
        params.push(Number(client_id));
      }

      if (status === "completed" && typeof month === "string" && /^\d{4}-\d{2}$/.test(month)) {
        const year = parseInt(month.slice(0, 4), 10);
        const mon = parseInt(month.slice(5, 7), 10);
        const start = new Date(year, mon - 1, 1);
        const end = new Date(mon === 12 ? year + 1 : year, mon === 12 ? 0 : mon, 1);
        const pad = (n) => String(n).padStart(2, "0");
        const startStr = start.getFullYear() + "-" + pad(start.getMonth() + 1) + "-" + pad(1);
        const endStr = end.getFullYear() + "-" + pad(end.getMonth() + 1) + "-" + pad(1);
        sql += " AND p.completed_at >= ? AND p.completed_at < ?";
        params.push(startStr, endStr);
      }

      sql += `
        ORDER BY
          p.completed ASC,
          FIELD(p.priority,'urgent','high','normal','low'),
          p.eta_date IS NULL,
          p.eta_date ASC,
          p.id DESC
      `;

      const [rows] = await db.query(sql, params);
      const projects = rows.map((row) => ({
        id: row.id,
        client_id: row.client_id,
        client_name: row.client_name || null,
        client_name_manual: row.client_name_manual || null,
        title: row.title,
        notes: row.notes || null,
        allocated_to: row.allocated_to || null,
        cost: row.cost != null ? Number(row.cost) : null,
        start_date: toISODate(row.start_date),
        eta_date: toISODate(row.eta_date),
        completed: row.completed ? 1 : 0,
        completed_at: row.completed_at ? toPlainDateTime(row.completed_at) : null,
        priority: row.priority || "normal",
        stage: row.stage || "not_started",
        progress_percent: row.progress_percent != null ? Number(row.progress_percent) : 0,
      }));

      return res.status(200).json({ ok: true, projects });
    }

    /* ---------- POST: create project ---------- */
    if (req.method === "POST") {
      const {
        client_id,
        client_name_manual,
        title,
        notes,
        allocated_to,
        cost,
        start_date,
        eta_date,
        priority,
        stage,
        progress_percent,
      } = req.body || {};

      if (!title) {
        return res.status(400).json({ ok: false, error: "Title is required" });
      }

      const validPriorities = ["low", "normal", "high", "urgent"];
      const validStages = ["not_started", "discovery", "design", "build", "qa", "live", "blocked", "completed"];

      const prio = validPriorities.includes(priority) ? priority : "normal";
      const stageVal = validStages.includes(stage) ? stage : "not_started";

      const now = new Date();
      const startDate = start_date || todayISO(now);

      const costVal =
        cost === undefined || cost === null || cost === "" ? null : Number(cost);

      const progressValRaw =
        progress_percent === undefined || progress_percent === null || progress_percent === ""
          ? 0
          : parseInt(progress_percent, 10);
      const progressVal = Number.isFinite(progressValRaw)
        ? Math.max(0, Math.min(100, progressValRaw))
        : 0;

      let clientIdVal = client_id ? Number(client_id) : null;
      const manualName = client_name_manual ? String(client_name_manual).trim() : null;
      const alloc = (allocated_to && String(allocated_to).trim()) || "Unassigned";

      // 🔄 AUTO-CREATE CLIENT IF NEEDED
      if (!clientIdVal && manualName) {
        // Try find existing client with same company name
        const [existing] = await db.query(
          "SELECT id FROM clients WHERE company = ? LIMIT 1",
          [manualName]
        );
        if (existing.length) {
          clientIdVal = existing[0].id;
        } else {
          // Create a minimal client row
          const [clientResult] = await db.query(
            `INSERT INTO clients
               (name, company, contact_name, email, phone, address1, city, status, website, notes, created_at)
             VALUES (?, ?, '', '', '', '', '', 'active', '', '', NOW())`,
            [manualName, manualName]
          );
          clientIdVal = clientResult.insertId;
        }
      }

      const [result] = await db.query(
        `INSERT INTO projects
           (client_id, client_name_manual, title, notes, allocated_to, cost, start_date, eta_date, priority, stage, progress_percent)
         VALUES (?,?,?,?,?,?,?,?,?,?,?)`,
        [
          clientIdVal || null,
          manualName,
          String(title).trim(),
          notes || null,
          alloc,
          costVal,
          startDate,
          eta_date || null,
          prio,
          stageVal,
          progressVal,
        ]
      );

      return res.status(201).json({
        ok: true,
        project: {
          id: result.insertId,
          client_id: clientIdVal || null,
          client_name: null,
          client_name_manual: manualName,
          title: String(title).trim(),
          notes: notes || null,
          allocated_to: alloc,
          cost: costVal,
          start_date: startDate,
          eta_date: eta_date || null,
          completed: 0,
          completed_at: null,
          priority: prio,
          stage: stageVal,
          progress_percent: progressVal,
        },
      });
    }

    /* ---------- PUT: update project ---------- */
    if (req.method === "PUT") {
      const {
        id,
        client_id,
        client_name_manual,
        title,
        notes,
        allocated_to,
        cost,
        start_date,
        eta_date,
        priority,
        stage,
        progress_percent,
        completed,
      } = req.body || {};

      if (!id) {
        return res.status(400).json({ ok: false, error: "Missing id" });
      }
      if (!title) {
        return res.status(400).json({ ok: false, error: "Title is required" });
      }

      const validPriorities = ["low", "normal", "high", "urgent"];
      const validStages = ["not_started", "discovery", "design", "build", "qa", "live", "blocked", "completed"];

      const prio = validPriorities.includes(priority) ? priority : "normal";
      const stageVal = validStages.includes(stage) ? stage : "not_started";

      const costVal =
        cost === undefined || cost === null || cost === "" ? null : Number(cost);

      const progressValRaw =
        progress_percent === undefined || progress_percent === null || progress_percent === ""
          ? 0
          : parseInt(progress_percent, 10);
      const progressVal = Number.isFinite(progressValRaw)
        ? Math.max(0, Math.min(100, progressValRaw))
        : 0;

      const clientIdVal = client_id ? Number(client_id) : null;
      const manualName = client_name_manual ? String(client_name_manual).trim() : null;
      const alloc = (allocated_to && String(allocated_to).trim()) || "Unassigned";

      const completedFlag = completed ? 1 : 0;
      const completedAtVal = completedFlag ? new Date() : null;
      const completedAtStr = completedAtVal
        ? completedAtVal.toISOString().slice(0, 19).replace("T", " ")
        : null;

      const startDate = start_date || todayISO();

      await db.query(
        `UPDATE projects
           SET client_id = ?, client_name_manual = ?, title = ?, notes = ?, allocated_to = ?,
               cost = ?, start_date = ?, eta_date = ?, priority = ?, stage = ?, progress_percent = ?,
               completed = ?, completed_at = ?
         WHERE id = ?`,
        [
          clientIdVal || null,
          manualName,
          String(title).trim(),
          notes || null,
          alloc,
          costVal,
          startDate,
          eta_date || null,
          prio,
          stageVal,
          progressVal,
          completedFlag,
          completedAtStr,
          id,
        ]
      );

      return res.status(200).json({ ok: true, message: "Project updated" });
    }

    /* ---------- DELETE: remove project ---------- */
    if (req.method === "DELETE") {
      const { id } = req.body || {};
      if (!id) {
        return res.status(400).json({ ok: false, error: "Missing id" });
      }
      await db.query("DELETE FROM projects WHERE id = ?", [id]);
      return res.status(200).json({ ok: true, message: "Project deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE", "OPTIONS"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Projects API error:", err);
    return res.status(500).json({
      ok: false,
      error: "Internal Server Error",
      detail: err.code || err.message || String(err),
    });
  }
}
