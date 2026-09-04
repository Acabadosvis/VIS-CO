// api/datos.js — VIS-CO M3 · Sync via GitHub file storage
// GET  /api/datos?tipo=redes|leads|reuniones&week=2026-09-01  → datos de esa semana
// POST /api/datos?tipo=redes|leads|reuniones&week=2026-09-01  → guarda/mergea datos

const OWNER = 'Acabadosvis';
const REPO   = 'VIS-CO';
const FILE   = 'data/visco-state.json';
const BRANCH = 'main';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const TOKEN = process.env.GH_TOKEN;
  if (!TOKEN) {
    return res.status(503).json({
      ok: false,
      error: 'GH_TOKEN no configurado. Ve a vercel.com → tu proyecto → Settings → Environment Variables y agrega GH_TOKEN.'
    });
  }

  const tipo = (req.query.tipo || '').replace(/[^a-z]/g, '');
  const week = (req.query.week || '').replace(/[^0-9\-]/g, '');

  if (!['redes', 'leads', 'reuniones'].includes(tipo)) {
    return res.status(400).json({ ok: false, error: 'tipo debe ser redes|leads|reuniones' });
  }

  const GH_URL = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${FILE}`;
  const GH_HDR = {
    Authorization: `Bearer ${TOKEN}`,
    Accept: 'application/vnd.github.v3+json',
    'User-Agent': 'VIS-CO-sync'
  };

  async function readState() {
    const r = await fetch(GH_URL, { headers: GH_HDR });
    if (r.status === 404) return { state: {}, sha: null };
    if (!r.ok) throw new Error('GitHub ' + r.status);
    const j = await r.json();
    const state = JSON.parse(Buffer.from(j.content.replace(/\n/g, ''), 'base64').toString('utf8'));
    return { state, sha: j.sha };
  }

  async function writeState(state, sha) {
    const content = Buffer.from(JSON.stringify(state, null, 0)).toString('base64');
    const body = { message: `sync: ${tipo}/${week || 'all'}`, content, branch: BRANCH };
    if (sha) body.sha = sha;
    const r = await fetch(GH_URL, {
      method: 'PUT',
      headers: { ...GH_HDR, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (!r.ok) {
      const txt = await r.text();
      // 409 = SHA conflict (otro dispositivo escribió al mismo tiempo) — reintentar
      if (r.status === 409) throw Object.assign(new Error('CONFLICT'), { conflict: true });
      throw new Error(txt.slice(0, 120));
    }
    const j = await r.json();
    return j.content?.sha;
  }

  /* ── GET: devuelve datos de un tipo + semana ── */
  if (req.method === 'GET') {
    try {
      const { state } = await readState();
      const tipoData = state[tipo] || {};
      const result   = week ? (tipoData[week] || {}) : tipoData;
      return res.json({ ok: true, tipo, week, data: result });
    } catch (e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  /* ── POST: mergea y guarda ── */
  if (req.method === 'POST') {
    let body = req.body;
    if (typeof body === 'string') { try { body = JSON.parse(body); } catch { body = {}; } }

    // Reintentos por conflicto de SHA (máx 3)
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const { state, sha } = await readState();
        if (!state[tipo]) state[tipo] = {};
        if (week) {
          state[tipo][week] = Object.assign({}, state[tipo][week] || {}, body);
        } else {
          Object.assign(state[tipo], body);
        }
        const newSha = await writeState(state, sha);
        return res.json({ ok: true, tipo, week, sha: newSha, saved: Object.keys(body) });
      } catch (e) {
        if (e.conflict && attempt < 2) continue; // reintentar
        return res.status(500).json({ ok: false, error: e.message });
      }
    }
  }

  return res.status(405).json({ error: 'Método no permitido' });
};
