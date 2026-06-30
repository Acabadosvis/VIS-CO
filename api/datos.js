// api/datos.js — VIS-CO M3 · Sync de Redes y Leads multi-dispositivo via Vercel KV
// GET  /api/datos?tipo=redes&week=2026-W26  → datos de esa semana
// GET  /api/datos?tipo=leads&week=2026-W26  → datos de esa semana
// POST /api/datos?tipo=redes&week=2026-W26  → guarda/mergea datos
// POST /api/datos?tipo=leads&week=2026-W26  → guarda/mergea datos

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const KV_URL   = process.env.KV_REST_API_URL;
  const KV_TOKEN = process.env.KV_REST_API_TOKEN;

  if (!KV_URL || !KV_TOKEN) {
    return res.status(503).json({
      ok: false,
      error: 'Vercel KV no configurado. Ve a Storage → Create Database en tu dashboard de Vercel.'
    });
  }

  const tipo = (req.query.tipo || '').replace(/[^a-z]/g, '');
  const week = (req.query.week || '').replace(/[^0-9W\-]/g, '') || 'noweek';

  if (!['redes', 'leads'].includes(tipo)) {
    return res.status(400).json({ ok: false, error: 'tipo debe ser redes|leads' });
  }

  const key = `vc3_${tipo}_${week}`;

  async function kvGet(k) {
    const r = await fetch(`${KV_URL}/pipeline`, {
      method:  'POST',
      headers: { Authorization: `Bearer ${KV_TOKEN}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify([['GET', k]])
    });
    const j = await r.json();
    const v = j[0]?.result;
    return v ? JSON.parse(v) : null;
  }

  async function kvSet(k, val) {
    await fetch(`${KV_URL}/pipeline`, {
      method:  'POST',
      headers: { Authorization: `Bearer ${KV_TOKEN}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify([['SET', k, JSON.stringify(val), 'EX', 15552000]])
    });
  }

  if (req.method === 'GET') {
    const data = await kvGet(key) || {};
    return res.json({ ok: true, tipo, week, data });
  }

  if (req.method === 'POST') {
    let body = req.body;
    if (typeof body === 'string') { try { body = JSON.parse(body); } catch { body = {}; } }

    // Merge con datos existentes (no borrar lo que ya había)
    const existing = await kvGet(key) || {};
    Object.assign(existing, body);
    await kvSet(key, existing);
    return res.json({ ok: true, tipo, week, saved: Object.keys(body) });
  }

  res.status(405).json({ error: 'Método no permitido' });
};
