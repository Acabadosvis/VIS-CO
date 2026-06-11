// api/soportes.js — VIS-CO · Cartera
// Lee y escribe soportes de pago en Upstash KV (compartido entre Nathalia y Contabilidad)

const KEY = 'visco_sop2';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  const BASE  = process.env.KV_REST_API_URL   || process.env.UPSTASH_REDIS_REST_URL;
  const TOKEN = process.env.KV_REST_API_TOKEN  || process.env.UPSTASH_REDIS_REST_TOKEN;

  if (!BASE || !TOKEN) {
    res.status(500).json({ error: 'KV_REST_API_URL / KV_REST_API_TOKEN no configurados' });
    return;
  }

  const headers = { Authorization: `Bearer ${TOKEN}`, 'Content-Type': 'application/json' };

  if (req.method === 'GET') {
    const r = await fetch(`${BASE}/get/${KEY}`, { headers });
    const json = await r.json();
    const data = json.result ? JSON.parse(json.result) : {};
    res.status(200).json(data);

  } else if (req.method === 'POST') {
    // req.body viene parsed por Vercel si Content-Type: application/json
    const payload = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
    const r = await fetch(`${BASE}/pipeline`, {
      method: 'POST',
      headers,
      body: JSON.stringify([['SET', KEY, payload]])
    });
    const json = await r.json();
    res.status(200).json({ ok: true, result: json });

  } else {
    res.status(405).json({ error: 'Método no permitido' });
  }
};
