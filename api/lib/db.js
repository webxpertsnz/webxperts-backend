// api/lib/db.js
import mysql from "mysql2/promise";

let pool;

/**
 * Singleton pool reused across serverless invocations (per instance).
 * Keeps connections low and queues extra requests instead of throwing
 * ER_USER_LIMIT_REACHED on Hostinger.
 */
export function getPool() {
  if (!pool) {
    pool = mysql.createPool({
      host: process.env.DB_HOST,
      user: process.env.DB_USER,
      password: process.env.DB_PASS,
      database: process.env.DB_NAME,
      port: Number(process.env.DB_PORT || 3306),
      ssl: { rejectUnauthorized: false },  // Hostinger + Vercel
      waitForConnections: true,
      connectionLimit: 3,   // keep this small on shared hosting
      maxIdle: 3,
      idleTimeout: 60000,
      queueLimit: 0
    });
  }
  return pool;
}
