// api/reuniones.js — VIS-CO M3 · Sync de reuniones multi-dispositivo via Vercel KV
// GET  /api/reuniones?week=2026-W26        → devuelve JSON con datos de los 3 comerciales
// POST /api/reuniones?week=2026-W26&comercial=KM → guarda datos del comercial en esa semana

module.exports = async function handler(req, res) {
  // CORS — acceso desde cualquier dispositivo
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const KV_URL   = process.env.KV_REST_API_URL;
  const KV_TOKEN = process.env.KV_REST_API_TOKEN;

  if (!KV_URL || !KV_TOKEN) {
    return res.status(503).json({
      ok: false,
      error: 'Vercel KV no configurado. Ve a Storage → Create Database en tu dashboard de Vercel y conecta el proyecto VIS-CO.'
    });
  }

  const week = (req.query.week || '').replace(/[^0-9W\-]/g, '') || 'noweek';
  const key  = `vc3_reus_${week}`;

  // ─── Helper: ejecutar comando Redis via REST pipeline ───────────────────────
  async function kvPipeline(commands) {
    const r = await fetch(`${KV_URL}/pipeline`, {
      method:  'POST',
      headers: { Authorization: `Bearer ${KV_TOKEN}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify(commands)
    });
    return r.json();
  }

  async function kvGet(k) {
    const j = await kvPipeline([['GET', k]]);
    const v = j[0]?.result;
    return v ? JSON.parse(v) : null;
  }

  async function kvSet(k, val) {
    // TTL = 180 días (datos de semanas viejas se limpian solos)
    await kvPipeline([['SET', k, JSON.stringify(val), 'EX', 15552000]]);
  }

  // ─── GET — leer semana completa ─────────────────────────────────────────────
  if (req.method === 'GET') {
    const data = await kvGet(key) || {};
    return res.json({ ok: true, week, data });
  }

  // ─── POST — guardar datos de un comercial ───────────────────────────────────
  if (req.method === 'POST') {
    const comercial = (req.query.comercial || '').toUpperCase().replace(/[^A-Z]/g, '');
    if (!comercial) return res.status(400).json({ ok: false, error: 'Falta ?comercial=KM|AB|NM|ADMIN' });

    let body = req.body;
    if (typeof body === 'string') { try { body = JSON.parse(body); } catch { body = {}; } }

    // Leer datos existentes de la semana y mergear
    const existing = await kvGet(key) || {};
    Object.assign(existing, body);

    await kvSet(key, existing);
    return res.json({ ok: true, week, comercial, saved: Object.keys(body) });
  }

  res.status(405).json({ error: 'Método no permitido' });
};
