// api/lib/db.js
import mysql from "mysql2/promise";

let pool;

/**
 * Singleton pool reused inside each serverless function instance.
 * Keep the connection count at 1 to stay under Hostinger's user limit.
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
      connectionLimit: 1,   // <- the key change
      maxIdle: 1,
      idleTimeout: 60000,
      queueLimit: 0
    });
  }
  return pool;
}
