import mysql from "mysql2/promise";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Max-Age", "86400");
  if (req.method === "OPTIONS") return res.status(200).end();

  const db = await mysql.createConnection({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASS,
    database: process.env.DB_NAME,
    port: Number(process.env.DB_PORT || 3306),
    ssl: { rejectUnauthorized: false }
  });

  try {
    if (req.method === "GET") {
      const [rows] = await db.query("SELECT * FROM tasks ORDER BY due_date ASC");
      return res.status(200).json(rows);
    }

    if (req.method === "POST") {
      const { client_id, title, details, due_date, status="open" } = req.body;
      await db.query(
        `INSERT INTO tasks (client_id, title, details, due_date, status, created_at)
         VALUES (?, ?, ?, ?, ?, NOW())`,
        [client_id, title ?? "", details ?? "", due_date ?? null, status]
      );
      return res.status(201).json({ message: "Task created" });
    }

    if (req.method === "PUT") {
      const { id, title, details, due_date, status } = req.body;
      await db.query(
        `UPDATE tasks
         SET title=?, details=?, due_date=?, status=?, updated_at=NOW()
         WHERE id=?`,
        [title ?? "", details ?? "", due_date ?? null, status ?? "open", id]
      );
      return res.status(200).json({ message: "Task updated" });
    }

    if (req.method === "DELETE") {
      const { id } = req.body;
      await db.query("DELETE FROM tasks WHERE id=?", [id]);
      return res.status(200).json({ message: "Task deleted" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} Not Allowed`);
  } catch (err) {
    console.error("Tasks API Error:", err);
    return res.status(500).json({ error: err.message || "Internal Server Error" });
  } finally {
    await db.end();
  }
}
