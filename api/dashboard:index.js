// api/dashboard/index.js
import { getPool } from "../lib/db.js";

function ymd(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}
function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function lastOfMonth(d)  { return new Date(d.getFullYear(), d.getMonth() + 1, 0); }
function fyStartFor(d) { // FY starts April 1
  const y = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;
  return new Date(y, 3, 1);
}
function isActiveInMonth(rec, monthDate) {
  const start = rec.start_date ? new Date(rec.start_date) : null;
  const end   = rec.end_date ? new Date(rec.end_date) : null;
  const mStart = firstOfMonth(monthDate);
  const mEnd   = lastOfMonth(monthDate);
  const activeFlag = rec.active == null || rec.active == 1;
  const afterStart = !start || start <= mEnd;
  const beforeEnd  = !end   || end   >= mStart;
  const cycleOk = String(rec.billing_cycle || "monthly").toLowerCase() === "monthly";
  return activeFlag && cycleOk && afterStart && beforeEnd;
}
const rowAmount = (r) => (r.amount != null ? Number(r.amount) : Number(r.quantity || 0) * Number(r.unit_amount || 0)) || 0;

export default async function handler(req, res) {
  // CORS (works for your CRM domain or '*' while testing)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Method Not Allowed" });

  const pool = getPool();
  const conn = await pool.getConnection(); // reuse ONE connection for all queries in this request
  try {
    const now = new Date();
    const today = ymd(now);
    const fyStart = fyStartFor(now);
    const sMonth = firstOfMonth(now);
    const eMonth = lastOfMonth(now);

    // --- ONE-OFF sums (done in SQL so we move less data) ---
    const [[oneoffYTD]] = await conn.query(
      `SELECT COALESCE(SUM(CASE WHEN amount IS NOT NULL THEN amount ELSE COALESCE(quantity,0)*COALESCE(unit_amount,0) END),0) AS total
       FROM oneoff_sales
       WHERE LOWER(status) IN ('sent','paid') AND sale_date BETWEEN ? AND ?`,
      [ymd(fyStart), today]
    );

    const [[oneoffThisMonth]] = await conn.query(
      `SELECT COALESCE(SUM(CASE WHEN amount IS NOT NULL THEN amount ELSE COALESCE(quantity,0)*COALESCE(unit_amount,0) END),0) AS total
       FROM oneoff_sales
       WHERE LOWER(status) IN ('sent','paid') AND sale_date BETWEEN ? AND ?`,
      [ymd(sMonth), ymd(eMonth)]
    );

    const [[draftsTotal]] = await conn.query(
      `SELECT COALESCE(SUM(CASE WHEN amount IS NOT NULL THEN amount ELSE COALESCE(quantity,0)*COALESCE(unit_amount,0) END),0) AS total
       FROM oneoff_sales WHERE LOWER(status)='draft'`
    );

    // --- RECURRING rows (we calculate monthly inclusion in JS for accuracy) ---
    const [recurring] = await conn.query(
      `SELECT id, amount, quantity, unit_amount, start_date, end_date, active, billing_cycle
       FROM recurring_sales`
    );

    // Helper to sum recurring for a given month
    const sumRecurringForMonth = (monthDate) =>
      recurring.filter(r => isActiveInMonth(r, monthDate))
               .reduce((a, r) => a + rowAmount(r), 0);

    // Recurring YTD: month-by-month from FY start to now
    let recurringYTD = 0;
    for (let cursor = firstOfMonth(fyStart); cursor <= eMonth; cursor = new Date(cursor.getFullYear(), cursor.getMonth()+1, 1)) {
      recurringYTD += sumRecurringForMonth(cursor);
    }

    const recurringMonth = sumRecurringForMonth(now);

    // --- Active clients / Open tasks counts ---
    const [[activeClients]] = await conn.query(`SELECT COUNT(*) AS c FROM clients WHERE LOWER(status)='active'`);
    const [[openTasks]]    = await conn.query(`SELECT COUNT(*) AS c FROM tasks   WHERE LOWER(status)='open'`);

    // --- Final dashboard numbers ---
    const ytd = Number(oneoffYTD.total) + Number(recurringYTD);
    const incomeThisMonth = Number(oneoffThisMonth.total) + Number(recurringMonth);
    const gstThisMonth = incomeThisMonth * 0.15;

    // Expected Income to March (project recurring only for remaining months)
    const m = now.getMonth(); // 0..11
    const monthsLeft = (m <= 2) ? (2 - m) : (14 - m);
    let projection = 0;
    let projCursor = new Date(now.getFullYear(), m + 1, 1);
    for (let i = 0; i < monthsLeft; i++) {
      projection += sumRecurringForMonth(projCursor);
      projCursor = new Date(projCursor.getFullYear(), projCursor.getMonth()+1, 1);
    }
    const expectedIncome = ytd + projection;

    return res.status(200).json({
      ok: true,
      cards: {
        expectedIncome,
        ytd,
        gstThisMonth,
        draftsTotal: Number(draftsTotal.total),
        activeClients: activeClients.c,
        openTasks: openTasks.c
      }
    });
  } catch (err) {
    console.error("Dashboard API Error:", err);
    return res.status(500).json({ ok:false, error: "Internal Server Error", detail: err.code || err.message });
  } finally {
    try { conn.release(); } catch {}
  }
}
