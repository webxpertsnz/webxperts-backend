// api/lib/db.js
import mysql from "mysql2/promise";

let pool;

/**
 * Reused pool per serverless instance.
 * Keep connections at 1 to avoid ER_USER_LIMIT_REACHED.
 */
export function getPool() {
  if (!pool) {
    pool = mysql.createPool({
      host: process.env.DB_HOST,
      user: process.env.DB_USER,
      password: process.env.DB_PASS,
      database: process.env.DB_NAME,
      port: Number(process.env.DB_PORT || 3306),
      ssl: { rejectUnauthorized: false },   // Hostinger + Vercel
      waitForConnections: true,
      connectionLimit: 1,
      maxIdle: 1,
      idleTimeout: 60_000,
      queueLimit: 0
    });
  }
  return pool;
}
