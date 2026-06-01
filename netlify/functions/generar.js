// netlify/functions/generar.js
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip'); // ← DEBE estar aquí arriba para que Netlify lo empaquete
 
// ── Helpers de texto ──────────────────────────────────────────────────────────
 
function fmtPesos(n) {
  return '$' + n.toLocaleString('es-CO');
}
 
function numLetras(n) {
  if (!n) return 'CERO';
  const u = ['','UN','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE','DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS','DIECISIETE','DIECIOCHO','DIECINUEVE'];
  const d = ['','DIEZ','VEINTE','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA'];
  const c = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS','SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS'];
  function m1000(x) {
    if (x === 0) return '';
    if (x < 20) return u[x];
    if (x < 100) { const [a, b] = [Math.floor(x / 10), x % 10]; return d[a] + (b ? ' Y ' + u[b] : ''); }
    if (x === 100) return 'CIEN';
    const [a, b] = [Math.floor(x / 100), x % 100];
    return c[a] + (b ? ' ' + m1000(b) : '');
  }
  const mil = Math.floor(n / 1000000), miles = Math.floor((n % 1000000) / 1000), uni = n % 1000;
  let r = '';
  if (mil)   r += (mil   === 1 ? 'UN MILLÓN'  : m1000(mil)   + ' MILLONES') + ' ';
  if (miles) r += (miles === 1 ? 'MIL'         : m1000(miles) + ' MIL')      + ' ';
  if (uni)   r += m1000(uni) + ' ';
  return r.trim() + ' PESOS M/CTE';
}
 
function pctTxt(p) {
  const t = {10:'DIEZ',13:'TRECE',14:'CATORCE',15:'QUINCE',17:'DIECISIETE',20:'VEINTE',25:'VEINTICINCO',30:'TREINTA',40:'CUARENTA',43:'CUARENTA Y TRES',50:'CINCUENTA'};
  return t[p] || String(p);
}
 
// ── Reemplazar en DOCX ────────────────────────────────────────────────────────
 
async function replaceInDocx(docxBuffer, replacements) {
  const zip = await JSZip.loadAsync(docxBuffer);
 
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('word/document.xml no encontrado dentro del .docx');
 
  let xml = await xmlFile.async('string');
 
  // Ordenar por longitud descendente para evitar reemplazos parciales
  const items = Object.entries(replacements).sort((a, b) => b[0].length - a[0].length);
 
  for (const [old, nw] of items) {
    const escaped = String(nw)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    const escapedOld = old
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    xml = xml.split(escapedOld).join(escaped);
  }
 
  zip.file('word/document.xml', xml);
 
  return zip.generateAsync({
    type: 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}
 
// ── Handler principal ─────────────────────────────────────────────────────────
 
exports.handler = async (event) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
  };
 
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: corsHeaders, body: '' };
  }
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers: corsHeaders, body: 'Method not allowed' };
  }
 
  try {
    const datos = JSON.parse(event.body);
    const {
      tipo, nombre, cedula, lugar_exp, direccion,
      cotizacion, valor_total, congelacion, fecha_cong,
      dia, mes, torre, apto, proyecto,
    } = datos;
 
    // Validación básica
    if (!tipo || !nombre || !valor_total) {
      return { statusCode: 400, headers: corsHeaders, body: JSON.stringify({ error: 'Faltan campos obligatorios: tipo, nombre, valor_total' }) };
    }
 
    const minutaNames = {
      acabadosvis:  'Minuta clientes Acabadosvis.docx',
      visualizza:   'Contrato clientes visualizza.docx',
      gran_central: 'CON - XXXX - 2026 NOMBRE CLIENTE Gran central.docx',
    };
 
    if (!minutaNames[tipo]) {
      return { statusCode: 400, headers: corsHeaders, body: JSON.stringify({ error: `Tipo desconocido: ${tipo}` }) };
    }
 
    // En Netlify los archivos del repo están en /var/task/
    // process.cwd() en Netlify Functions apunta a /var/task (raíz del bundle)
    const minutaPath = path.join(process.cwd(), minutaNames[tipo]);
 
    if (!fs.existsSync(minutaPath)) {
      return { statusCode: 500, headers: corsHeaders, body: JSON.stringify({ error: `Minuta no encontrada en: ${minutaPath}` }) };
    }
 
    const docxBuffer = fs.readFileSync(minutaPath);
 
    const vt   = Number(valor_total);
    const cong = Number(congelacion) || 5000000;
    const cotizNum = cotizacion.replace('CO-', '').replace('-2026', '').replace('-', '').trim();
    const conNum   = cotizacion.startsWith('CO-') ? cotizacion.replace('CO-', 'CON-') : cotizacion;
 
    let reps = {};
 
    // ── AcabadosVIS ──
    if (tipo === 'acabadosvis') {
      const resto    = vt - cong;
      const p2       = Math.round(resto * 0.50);
      const p3       = Math.round(resto * 0.20);
      const p4       = Math.round(resto * 0.20);
      const p5       = vt - cong - p2 - p3 - p4;
      const pctCong  = Math.round((cong / vt) * 100);
      const pct2     = Math.round((p2   / vt) * 100);
 
      reps = {
        'XXXXX':                                                           cotizNum,
        'XXXXXXXXXXXXXXXXXXXX':                                            nombre,
        'XXXXXXXXXXXXX':                                                   cedula,
        'XXXXXXXXXX':                                                      lugar_exp,
        'XXXXXXX':                                                         cotizNum,
        'Torre XX':                                                        `Torre ${torre}`,
        'Apartamento XXXX':                                                `Apartamento ${apto}`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE':                   numLetras(vt),
        '$XX.XXX.XXX':                                                     fmtPesos(vt),
        'XXXXXX por ciento XX%':                                           `${pctTxt(pctCong)} por ciento ${pctCong}%`,
        'XXXXXXXX por ciento XX%':                                         `${pctTxt(pct2)} por ciento ${pct2}%`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)':       `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)':  `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)': `${numLetras(p4)} (${fmtPesos(p4)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)':  `${numLetras(p5)} (${fmtPesos(p5)})`,
        'XXXXXXXXX ':                                                      fecha_cong + ' ',
        'XX d&#237;a del mes de xxxxx':                                    `${dia} d&#237;a del mes de ${mes}`,
        'xxxxx':                                                           mes,
        'CO &#8211; XXXXXX':                                              `CO &#8211; ${cotizNum}`,
        'CO - XXXXXXX - 2026':                                             cotizacion,
      };
 
    // ── Visualizza ──
    } else if (tipo === 'visualizza') {
      const p1 = Math.round(vt * 0.50);
      const p2 = Math.round(vt * 0.20);
      const p3 = Math.round(vt * 0.20);
      const p4 = vt - p1 - p2 - p3;
 
      reps = {
        'CON-XXXXX-26':                                                     conNum,
        'XXXXXXXXXXXXXXXXXXX':                                              nombre,
        'XXXXXXXXXXX':                                                      cedula,
        'XXXXXXXXXX':                                                       lugar_exp,
        'XXXXXXXXXXXXXXXXXXXXXXX':                                          `${proyecto} ${torre} Apartamento ${apto}`,
        'CO-XXXXXX-25':                                                     cotizacion,
        'CO-XXXXX-25':                                                      cotizacion,
        'CO-XXXXXXXX':                                                      cotizacion,
        "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX'XXX.XXX)":         `${numLetras(vt)} (${fmtPesos(vt)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)':               `${numLetras(p1)} (${fmtPesos(p1)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)':   `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)':         `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($X,XXX,XXX)':                   `${numLetras(p4)} (${fmtPesos(p4)})`,
        'MES DE XXXXXX DE 202X Y UN MES MAS':                              fecha_cong.toUpperCase() + ' Y UN MES MAS',
        '(23) d&#237;a del mes de julio':                                  `(${dia}) d&#237;a del mes de ${mes}`,
      };
 
    // ── Gran Central ──
    } else if (tipo === 'gran_central') {
      const resto   = vt - cong;
      const p2      = Math.round(resto * 0.50);
      const p3      = Math.round(resto * 0.20);
      const p4      = Math.round(resto * 0.20);
      const p5      = vt - cong - p2 - p3 - p4;
      const pctCong = Math.round((cong / vt) * 100);
      const pct2    = Math.round((p2   / vt) * 100);
 
      reps = {
        'CON-2702-2026':                                                            conNum,
        'XXXXXXXXXXXXXXXXXXXXXXXXX':                                                nombre,
        'XXXXXXXXXXXXXXXX':                                                         cedula,
        'XXXXXXXXXX':                                                               lugar_exp,
        'Apartamento XXXX':                                                         `Apartamento ${apto}`,
        'CO-XXXX-2026':                                                             cotizacion,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX M/CTE ($XX.XXX.XXX)':     `${numLetras(vt)} (${fmtPesos(vt)})`,
        'XXXXXXX por ciento XX%':                                                   `${pctTxt(pctCong)} por ciento ${pctCong}%`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX M/CTE ($X.XXX.XXX)':                   `${numLetras(cong)} (${fmtPesos(cong)})`,
        'XXXXXXXXX por ciento XX%':                                                 `${pctTxt(pct2)} por ciento ${pct2}%`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX.XXX.XXXX)':    `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXX por ciento XX%':                                                    'VEINTE por ciento 20%',
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($X,XXX,XXX)':         `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXPESOS ($X,XXX,XXX)':            `${numLetras(p4)} (${fmtPesos(p4)})`,
        'MES DE NOVIEMBRE DE 2026 Y UN MES MAS':                                   fecha_cong.toUpperCase() + ' Y UN MES MAS',
        '(27) d&#237;a del mes de abril':                                           `(${dia}) d&#237;a del mes de ${mes}`,
      };
    }
 
    const resultBuffer = await replaceInDocx(docxBuffer, reps);
    const base64 = resultBuffer.toString('base64');
 
    return {
      statusCode: 200,
      headers: {
        ...corsHeaders,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${conNum}.docx"`,
      },
      body: base64,
      isBase64Encoded: true,
    };
 
  } catch (error) {
    console.error('[generar.js] ERROR:', error);
    return {
      statusCode: 500,
      headers: { 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ error: error.message, stack: error.stack }),
    };
  }
};
 
