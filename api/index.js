// api/index.js
export default function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();
  return res.status(200).json({ ok: true, service: "webxperts-backend", ts: Date.now() });
}
