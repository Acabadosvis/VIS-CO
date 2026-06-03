// api/generar.js  — Vercel Serverless Function
'use strict';

const fs    = require('fs');
const path  = require('path');
const JSZip = require('jszip');

// ── Helpers de texto ──────────────────────────────────────────────────────────

function fmtPesos(n) {
  return '$' + Number(n).toLocaleString('es-CO');
}

function numLetras(n) {
  n = Math.round(Number(n));
  if (!n) return 'CERO';
  const u = ['','UN','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE',
             'DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS',
             'DIECISIETE','DIECIOCHO','DIECINUEVE'];
  const d = ['','DIEZ','VEINTE','TREINTA','CUARENTA','CINCUENTA',
             'SESENTA','SETENTA','OCHENTA','NOVENTA'];
  const c = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
             'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS'];

  function m1000(x) {
    if (x === 0)   return '';
    if (x < 20)    return u[x];
    if (x < 100) {
      const [a, b] = [Math.floor(x / 10), x % 10];
      return d[a] + (b ? ' Y ' + u[b] : '');
    }
    if (x === 100) return 'CIEN';
    const [a, b] = [Math.floor(x / 100), x % 100];
    return c[a] + (b ? ' ' + m1000(b) : '');
  }

  const mil   = Math.floor(n / 1_000_000);
  const miles = Math.floor((n % 1_000_000) / 1_000);
  const uni   = n % 1_000;
  let r = '';
  if (mil)   r += (mil   === 1 ? 'UN MILLÓN'  : m1000(mil)   + ' MILLONES') + ' ';
  if (miles) r += (miles === 1 ? 'MIL'         : m1000(miles) + ' MIL')      + ' ';
  if (uni)   r += m1000(uni) + ' ';
  return r.trim() + ' PESOS M/CTE';
}

function pctTxt(p) {
  const t = {
    5:'CINCO', 10:'DIEZ', 13:'TRECE', 14:'CATORCE', 15:'QUINCE',
    17:'DIECISIETE', 20:'VEINTE', 25:'VEINTICINCO', 30:'TREINTA',
    40:'CUARENTA', 43:'CUARENTA Y TRES', 50:'CINCUENTA',
  };
  return t[p] || String(p);
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ── Motor de reemplazo con merge de runs ──────────────────────────────────────

function extractParaText(paraXml) {
  const matches = paraXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
  return matches.map(m => m.replace(/<[^>]*>/g, '')).join('');
}

function applyReplacements(text, sortedReps) {
  let result = text;
  for (const [k, v] of sortedReps) {
    if (result.includes(k)) result = result.split(k).join(String(v));
  }
  return result;
}

function replaceParagraphText(paraXml, sortedReps) {
  const original = extractParaText(paraXml);
  const updated  = applyReplacements(original, sortedReps);
  if (updated === original) return paraXml;
  let firstDone = false;
  return paraXml.replace(/<w:t([^>]*)>[^<]*<\/w:t>/g, (match, attrs) => {
    if (!firstDone) {
      firstDone = true;
      return `<w:t xml:space="preserve">${escapeXml(updated)}</w:t>`;
    }
    return `<w:t></w:t>`;
  });
}

async function replaceInDocx(docxBuffer, replacements) {
  const zip = await JSZip.loadAsync(docxBuffer);
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('word/document.xml no encontrado en el .docx');
  let xml = await xmlFile.async('string');
  const sorted = Object.entries(replacements).sort((a, b) => b[0].length - a[0].length);
  xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, para => replaceParagraphText(para, sorted));
  zip.file('word/document.xml', xml);
  return zip.generateAsync({
    type: 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

// ── Handler Vercel ────────────────────────────────────────────────────────────

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')   return res.status(405).end('Method not allowed');

  try {
    // Vercel parsea el body automáticamente si Content-Type es application/json
    const datos = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;

    const {
      tipo,
      nombre       = '',
      cedula       = '',
      lugar_exp    = '',
      direccion    = '',
      correo       = '',
      celular      = '',   // el form envía 'celular', no 'telefono'
      cotizacion   = '',
      valor_total,
      congelacion,
      fecha_cong   = '',
      dia          = '',
      mes          = '',
      torre        = '',
      apto         = '',
      proyecto     = '',
    } = datos;

    if (!tipo || !nombre || !valor_total) {
      return res.status(400).json({ error: 'Faltan campos obligatorios: tipo, nombre, valor_total' });
    }
    if (!fecha_cong) {
      return res.status(400).json({ error: 'Falta la fecha de vigencia de la cotización (fecha_cong). Por favor ingrésala para continuar.' });
    }

    const minutaNames = {
      acabadosvis : 'Minuta clientes Acabadosvis.docx',
      visualizza  : 'Contrato clientes visualizza.docx',
      gran_central: 'CON - XXXX - 2026 NOMBRE CLIENTE Gran central.docx',
    };

    if (!minutaNames[tipo]) {
      return res.status(400).json({ error: `Tipo desconocido: ${tipo}` });
    }

    const minutaPath = path.join(process.cwd(), minutaNames[tipo]);
    if (!fs.existsSync(minutaPath)) {
      return res.status(500).json({ error: `Minuta no encontrada en: ${minutaPath}` });
    }

    const docxBuffer = fs.readFileSync(minutaPath);

    // ── Cálculos ──────────────────────────────────────────────────────────────
    const vt   = Number(valor_total);
    const cong = Number(congelacion) || 5_000_000;

    const cotizNum = cotizacion
      .replace(/^[Cc][Oo][-\s]*/, '')
      .replace(/[-\s]*20\d\d$/, '')
      .replace(/-/g, '')
      .trim();

    const conNum   = `CON-${cotizNum}-2026`;
    const cotizFmt = `CO - ${cotizNum}`;  // el año ya está en la minuta, no duplicar

    let reps = {};

    // ── AcabadosVIS ───────────────────────────────────────────────────────────
    if (tipo === 'acabadosvis') {
      const resto   = vt - cong;
      const p2      = Math.round(resto * 0.50);
      const p3      = Math.round(resto * 0.20);
      const p4      = Math.round(resto * 0.20);
      const p5      = vt - cong - p2 - p3 - p4;
      const pctCong = Math.round((cong / vt) * 100);
      const pct2    = Math.round((p2   / vt) * 100);

      reps = {
        'No. CON - XXXXX'                                                          : `No. CON - ${cotizNum}`,
        'XXXXXXXXXXXXXXXXXXXX mayor de edad'                                        : `${nombre} mayor de edad`,
        'ciudadanía No. XXXXXXXXXXXXX expedida'                                    : `ciudadanía No. ${cedula} expedida`,
        'expedida en XXXXXXXXXX,'                                                  : `expedida en ${lugar_exp},`,
        'ubicado en XXXXXXXXXXXXXXXXXXXX Torre'                                    : `ubicado en ${proyecto} Torre`,
        'Torre XX Apartamento XXXX'                                                : `Torre ${torre} Apartamento ${apto}`,
        'CO - XXXXXXX'                                                             : cotizFmt,
        'CO – XXXXXX'                                                              : `CO – ${cotizNum}`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)'             : `${numLetras(vt)} (${fmtPesos(vt)})`,
        'XXXXXXXX por ciento XX%'                                                  : `${pctTxt(pct2)} por ciento`,
        'XXXXXX por ciento XX%'                                                    : `${pctTxt(pctCong)} por ciento`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)'               : `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)': `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)': `${numLetras(p4)} (${fmtPesos(p4)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)': `${numLetras(p5)} (${fmtPesos(p5)})`,
        'hasta el XXXXXXXXX y'                                                     : `hasta el ${fecha_cong} y`,
        'Dirección: xxxxxxxxxxxxxx'                                                : `Dirección: ${direccion}`,
        'Lugar: XXXXXXXXXX'                                                        : `Lugar: ${lugar_exp}`,
        'Teléfono: 3XXXXXXXXXXXXX'                                                 : `Teléfono: ${celular}`,
        'Email: XXXXXXXXXXXXXXXXXXXX'                                              : `Email: ${correo}`,
        'XX día del mes de xxxxx'                                                  : `${dia} día del mes de ${mes}`,
        'Nombre: XXXXXXXXXXXXXXXXX'                                                : `Nombre: ${nombre}`,
        'Email: XXXXXXXXXXXXXXXXXXXXX'                                             : `Email: ${correo}`,
      };

    // ── Visualizza ────────────────────────────────────────────────────────────
    } else if (tipo === 'visualizza') {
      // Congelación fija ($5M hardcodeada en la minuta) + 50/20/20/10 del resto
      const resto   = vt - cong;
      const p2      = Math.round(resto * 0.50);
      const p3      = Math.round(resto * 0.20);
      const p4      = Math.round(resto * 0.20);
      const p5      = vt - cong - p2 - p3 - p4;
      const pctCong = Math.round((cong / vt) * 100);
      const pct2    = Math.round((p2   / vt) * 100);

      reps = {
        // Número de contrato
        'CON-XXXXX-26'                                                                       : conNum,
        // Identificación cliente (19-X nombre, 11-X cédula, contexto para lugar_exp)
        'XXXXXXXXXXXXXXXXXXX'                                                                : nombre,
        'XXXXXXXXXXX'                                                                        : cedula,
        'expedida en XXXXXXXXXX,'                                                            : `expedida en ${lugar_exp},`,
        // Objeto PRIMERA (23-X proyecto+torre+apto)
        'XXXXXXXXXXXXXXXXXXXXXXX'                                                            : `${proyecto} Torre ${torre} Apartamento ${apto}`,
        // Cotización en cuerpo
        'CO-XXXXXX-25'                                                                       : `CO-${cotizNum}-2026`,
        'CO-XXXXX-25'                                                                        : `CO-${cotizNum}-2026`,
        'CO-XXXXXXXX'                                                                        : `CO-${cotizNum}-2026`,
        // Valor total (formato con apóstrofo en la minuta)
        "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX'XXX.XXX)"                           : `${numLetras(vt)} (${fmtPesos(vt)})`,
        // Porcentajes pagos 1 y 2 (3/4/5 son estáticos en la minuta: 20/20/10)
        'XXXXXX por ciento XX%'                                                              : `${pctTxt(pctCong)} por ciento ${pctCong}%`,
        'XXXXXXXX por ciento XX%'                                                            : `${pctTxt(pct2)} por ciento ${pct2}%`,
        // Valores pagos (el valor de congelación es estático $5M en la minuta)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)'                         : `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)': `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)': `${numLetras(p4)} (${fmtPesos(p4)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)': `${numLetras(p5)} (${fmtPesos(p5)})`,
        // Parágrafo primero (fecha vigencia valor)
        'MES DE XXXXXX DE 202X Y UN MES MAS'                                                : fecha_cong.toUpperCase() + ' Y UN MES MAS',
        // Notificaciones contratante (con contexto para evitar conflictos)
        'Dirección: XXXXXXXXXXXXXXXX'                                                        : `Dirección: ${direccion}`,
        'Lugar: XXXXXXX'                                                                     : `Lugar: ${lugar_exp}`,
        'Teléfono: XXXXXXXXXX'                                                               : `Teléfono: ${celular}`,
        'Email: XXXXXXXXXXXXXXX'                                                             : `Email: ${correo}`,
        // Fecha suscripción y firma
        '(23) día del mes de julio'                                                          : `(${dia}) día del mes de ${mes}`,
        'Nombre: XXXXXXXXXXXXXXXXXXXX'                                                       : `Nombre: ${nombre}`,
        'Email: XXXXXXXXXXXXXXXXXXXXX'                                                       : `Email: ${correo}`,
      };

    // ── Gran Central ──────────────────────────────────────────────────────────
    } else if (tipo === 'gran_central') {
      const resto   = vt - cong;
      const p2      = Math.round(resto * 0.50);
      const p3      = Math.round(resto * 0.20);
      const p4      = Math.round(resto * 0.20);
      const p5      = vt - cong - p2 - p3 - p4;
      const pctCong = Math.round((cong / vt) * 100);
      const pct2    = Math.round((p2   / vt) * 100);

      reps = {
        'CON-2702-2026'                                                                  : conNum,
        'XXXXXXXXXXXXXXXXXXXXXXXXX'                                                      : nombre,
        'XXXXXXXXXXXXXXXX'                                                               : cedula,
        'XXXXXXXXXX'                                                                     : lugar_exp,
        'Apartamento XXXX'                                                               : `Apartamento ${apto}`,
        'CO-XXXX-2026'                                                                   : cotizacion,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX M/CTE ($XX.XXX.XXX)'           : `${numLetras(vt)} (${fmtPesos(vt)})`,
        'XXXXXXX por ciento XX%'                                                         : `${pctTxt(pctCong)} por ciento ${pctCong}%`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX M/CTE ($X.XXX.XXX)'                          : `${numLetras(cong)} (${fmtPesos(cong)})`,
        'XXXXXXXXX por ciento XX%'                                                       : `${pctTxt(pct2)} por ciento ${pct2}%`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX.XXX.XXXX)'          : `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXX por ciento XX%'                                                          : 'VEINTE por ciento 20%',
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($X,XXX,XXX)'               : `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXPESOS ($X,XXX,XXX)'                  : `${numLetras(p4)} (${fmtPesos(p4)})`,
        'MES DE NOVIEMBRE DE 2026 Y UN MES MAS'                                          : fecha_cong.toUpperCase() + ' Y UN MES MAS',
        '(27) día del mes de abril'                                                       : `(${dia}) día del mes de ${mes}`,
        'Dirección: xxxxxxxxxxxxxx'                                                       : `Dirección: ${direccion}`,
        'Lugar: XXXXXXXXXX'                                                               : `Lugar: ${lugar_exp}`,
        'Teléfono: 3XXXXXXXXXXXXX'                                                        : `Teléfono: ${celular}`,
        'Email: XXXXXXXXXXXXXXXXXXXX'                                                      : `Email: ${correo}`,
        'Nombre: XXXXXXXXXXXXXXXXX'                                                        : `Nombre: ${nombre}`,
        'Email: XXXXXXXXXXXXXXXXXXXXX'                                                      : `Email: ${correo}`,
        'XX día del mes de xxxxx'                                                          : `${dia} día del mes de ${mes}`,
      };
    }

    const resultBuffer = await replaceInDocx(docxBuffer, reps);
    const conNumFinal  = `CON-${cotizNum}-2026`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${conNumFinal}.docx"`);
    return res.send(resultBuffer);

  } catch (error) {
    console.error('[generar.js] ERROR:', error);
    return res.status(500).json({ error: error.message, stack: error.stack });
  }
};
