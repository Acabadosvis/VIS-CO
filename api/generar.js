// api/generar.js — Vercel Serverless Function — v2.2
'use strict';

const fs   = require('fs');
const path = require('path');
const JSZip = require('jszip');

// ── Helpers ───────────────────────────────────────────────────────────────────

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
    if (x === 0) return '';
    if (x < 20)  return u[x];
    if (x < 100) { const [a,b]=[Math.floor(x/10),x%10]; return d[a]+(b?' Y '+u[b]:''); }
    if (x === 100) return 'CIEN';
    const [a,b]=[Math.floor(x/100),x%100]; return c[a]+(b?' '+m1000(b):'');
  }
  const mil=Math.floor(n/1_000_000), miles=Math.floor((n%1_000_000)/1_000), uni=n%1_000;
  let r='';
  if (mil)   r+=(mil===1?'UN MILLÓN':m1000(mil)+' MILLONES')+' ';
  if (miles) r+=(miles===1?'MIL':m1000(miles)+' MIL')+' ';
  if (uni)   r+=m1000(uni)+' ';
  // En español: "CINCO MILLONES DE PESOS" (DE cuando sólo hay millones exactos)
  const soloPorMilOnesPuros = mil > 0 && miles === 0 && uni === 0;
  return r.trim()+(soloPorMilOnesPuros?' DE':'')+' PESOS M/CTE';
}

// Devuelve "veinte por ciento 20%" — minúsculas, sin paréntesis (igual que el docx hardcoded)
function pctFmt(p) {
  const t = {
    5:'cinco',10:'diez',13:'trece',14:'catorce',15:'quince',
    17:'diecisiete',20:'veinte',25:'veinticinco',30:'treinta',
    40:'cuarenta',43:'cuarenta y tres',50:'cincuenta',
  };
  return (t[p]||String(p))+' por ciento '+p+'%';
}

function escapeXml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── Motor de reemplazo por párrafo ────────────────────────────────────────────

function extractParaText(paraXml) {
  const matches = paraXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
  return matches.map(m => m.replace(/<[^>]*>/g,'')).join('');
}

function applyReplacements(text, sortedReps) {
  let result = text;
  for (const [k,v] of sortedReps) {
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
    if (!firstDone) { firstDone=true; return `<w:t xml:space="preserve">${escapeXml(updated)}</w:t>`; }
    return `<w:t></w:t>`;
  });
}

async function replaceInDocx(docxBuffer, replacements) {
  const zip = await JSZip.loadAsync(docxBuffer);
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('word/document.xml no encontrado en el .docx');
  let xml = await xmlFile.async('string');
  const sorted = Object.entries(replacements).sort((a,b) => b[0].length - a[0].length);
  xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, para => replaceParagraphText(para, sorted));
  zip.file('word/document.xml', xml);
  return zip.generateAsync({
    type:'nodebuffer',
    mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

// ── Handler Vercel ────────────────────────────────────────────────────────────

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin','*');
  res.setHeader('Access-Control-Allow-Headers','Content-Type');
  res.setHeader('Access-Control-Allow-Methods','POST, OPTIONS');

  if (req.method==='OPTIONS') { res.status(200).end(); return; }
  if (req.method!=='POST')    { res.status(405).json({error:'Method not allowed'}); return; }

  try {
    const datos = typeof req.body==='string' ? JSON.parse(req.body) : req.body;
    const {
      tipo, nombre='', cedula='', lugar_exp='', direccion='', correo='',
      telefono='', cotizacion='', valor_total, congelacion,
      fecha_cong='', dia='', mes='', torre='', apto='', proyecto='',
    } = datos;

    if (!tipo || !nombre || !valor_total) {
      res.status(400).json({error:'Faltan campos: tipo, nombre, valor_total'}); return;
    }
    if (!fecha_cong) {
      res.status(400).json({error:'Falta la fecha de vigencia (fecha_cong).'}); return;
    }

    const minutaNames = {
      acabadosvis           : 'Minuta clientes Acabadosvis.docx',
      visualizza            : 'Contrato clientes visualizza.docx',
      gran_central_sin_bono : 'CON - XXXX - 2026 NOMBRE CLIENTE Gran central sin bono.docx',
      gran_central_con_bono : 'CON - XXXX - 2026 NOMBRE CLIENTE Gran central con bono.docx',
    };
    if (!minutaNames[tipo]) { res.status(400).json({error:`Tipo desconocido: ${tipo}`}); return; }

    const minutaPath = path.join(process.cwd(), minutaNames[tipo]);
    if (!fs.existsSync(minutaPath)) {
      res.status(500).json({error:`Minuta no encontrada: ${minutaPath}`}); return;
    }
    const docxBuffer = fs.readFileSync(minutaPath);

    const vt   = Number(valor_total);
    const cong = Number(congelacion) || 5_000_000;

    const cotizNum = cotizacion
      .replace(/^[Cc][Oo][-\s]*/,'').replace(/[-\s]*20\d\d$/,'').replace(/-/g,'').trim();
    const conNum   = `CON-${cotizNum}-2026`;
    const cotizFmt = `CO - ${cotizNum} - 2026`;

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
        // TÍTULO: "No. CON - XXXXX\xa0-\xa02026"  (con espacios duros)
        'No. CON - XXXXX\xa0-\xa02026'                                        : `No. CON - ${cotizNum} - 2026`,

        // IDENTIFICACIÓN: nombre (20 X) + cédula (13 X) con contexto + lugar (10 X)
        'XXXXXXXXXXXXXXXXXXXX mayor de edad'                                   : `${nombre} mayor de edad`,
        'ciudadanía No. XXXXXXXXXXXXX expedida'                               : `ciudadanía No. ${cedula} expedida`,
        'expedida en XXXXXXXXXX,'                                              : `expedida en ${lugar_exp},`,

        // PRIMERA — objeto: proyecto, torre, apto
        'ubicado en XXXXXXXXXXXXXXXXXXXX Torre'                               : `ubicado en ${proyecto} Torre`,
        'Torre XX Apartamento XXXX'                                            : `Torre ${torre} Apartamento ${apto}`,

        // PRIMERA — cotización "CO - XXXXXXX\xa0-\xa02026"
        'CO - XXXXXXX\xa0-\xa02026'                                           : cotizFmt,

        // QUINTA — cotización "CO – XXXXXX\xa0-\xa02026" (guión largo + espacio duro)
        'CO – XXXXXX\xa0-\xa02026'                                       : `CO – ${cotizNum} - 2026`,

        // QUINTA — valor total (33 X + M/CTE)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)'         : `${numLetras(vt)} (${fmtPesos(vt)})`,

        // PAGO 1: valor congelación — texto HARDCODEADO en el docx, se reemplaza dinámicamente
        'CINCO MILLONES DE PESOS M/CTE ($5.000.000)'                          : `${numLetras(cong)} (${fmtPesos(cong)})`,

        // PAGOS 1 y 2: porcentajes (8 X primero, luego 6 X — por longitud)
        'XXXXXXXX por ciento XX%'                                              : pctFmt(pct2),
        'XXXXXX por ciento XX%'                                                : pctFmt(pctCong),

        // PAGO 2: valor anticipo (31 X)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)'           : `${numLetras(p2)} (${fmtPesos(p2)})`,

        // PAGO 3: valor (59 X)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)' : `${numLetras(p3)} (${fmtPesos(p3)})`,

        // PAGO 4: valor (62 X)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)' : `${numLetras(p4)} (${fmtPesos(p4)})`,

        // PAGO 5: valor (61 X)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)' : `${numLetras(p5)} (${fmtPesos(p5)})`,

        // PARÁGRAFO 1: vigencia
        'hasta el XXXXXXXXX y'                                                : `hasta el ${fecha_cong} y`,

        // NOTIFICACIONES
        'Dirección: xxxxxxxxxxxxxx'                                            : `Dirección: ${direccion}`,
        'Lugar: XXXXXXXXXX'                                                    : `Lugar: ${lugar_exp}`,
        'Teléfono: 3XXXXXXXXXXXXX'                                             : `Teléfono: ${telefono}`,
        'Email: XXXXXXXXXXXXXXXXXXXX'                                          : `Email: ${correo}`,

        // FECHA FIRMA
        'XX día del mes de xxxxx'                                              : `${dia} día del mes de ${mes}`,

        // FIRMA
        'Nombre: XXXXXXXXXXXXXXXXX'                                            : `Nombre: ${nombre}`,
        'Email: XXXXXXXXXXXXXXXXXXXXX'                                         : `Email: ${correo}`,
      };

    // ── Visualizza ────────────────────────────────────────────────────────────
    } else if (tipo === 'visualizza') {
      const p1 = Math.round(vt * 0.50);
      const p2 = Math.round(vt * 0.20);
      const p3 = Math.round(vt * 0.20);
      const p4 = vt - p1 - p2 - p3;

      reps = {
        'CON-XXXXX-26'                                                         : conNum,
        'XXXXXXXXXXXXXXXXXXX'                                                  : nombre,
        'XXXXXXXXXXX'                                                          : cedula,
        'XXXXXXXXXX'                                                           : lugar_exp,
        'XXXXXXXXXXXXXXXXXXXXXXX'                                              : `${proyecto} ${torre} Apartamento ${apto}`,
        'CO-XXXXXX-25'                                                         : cotizacion,
        'CO-XXXXX-25'                                                          : cotizacion,
        'CO-XXXXXXXX'                                                          : cotizacion,
        "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX'XXX.XXX)"             : `${numLetras(vt)} (${fmtPesos(vt)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)'                   : `${numLetras(p1)} (${fmtPesos(p1)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)'       : `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)'             : `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($X,XXX,XXX)'                       : `${numLetras(p4)} (${fmtPesos(p4)})`,
        'MES DE XXXXXX DE 202X Y UN MES MAS'                                  : fecha_cong.toUpperCase()+' Y UN MES MAS',
        '(23) día del mes de julio'                                            : `(${dia}) día del mes de ${mes}`,
        'Dirección: xxxxxxxxxxxxxx'                                            : `Dirección: ${direccion}`,
        'Lugar: XXXXXXXXXX'                                                    : `Lugar: ${lugar_exp}`,
        'Teléfono: 3XXXXXXXXXXXXX'                                             : `Teléfono: ${telefono}`,
        'Email: XXXXXXXXXXXXXXXXXXXX'                                          : `Email: ${correo}`,
        'Nombre: XXXXXXXXXXXXXXXXX'                                            : `Nombre: ${nombre}`,
        'Email: XXXXXXXXXXXXXXXXXXXXX'                                         : `Email: ${correo}`,
        'XX día del mes de xxxxx'                                              : `${dia} día del mes de ${mes}`,
      };

    // ── Gran Central Sin Bono ─────────────────────────────────────────────────
    } else if (tipo === 'gran_central_sin_bono') {
      const resto   = vt - cong;
      const p2      = Math.round(resto * 0.50);
      const p3      = Math.round(resto * 0.20);
      const p4      = Math.round(resto * 0.20);
      const p5      = vt - cong - p2 - p3 - p4;
      const pctCong = Math.round((cong / vt) * 100);
      const pct2    = Math.round((p2   / vt) * 100);

      reps = {
        // TÍTULO
        'CON-xxxxxxxx'                                                         : conNum,

        // IDENTIFICACIÓN — nombre (25 X), cédula con contexto (16 X), lugar (10 X)
        'XXXXXXXXXXXXXXXXXXXXXXXXX, mayor de edad'                             : `${nombre}, mayor de edad`,
        'ciudadanía No. XXXXXXXXXXXXXXXX expedida'                            : `ciudadanía No. ${cedula} expedida`,
        'expedida en XXXXXXXXXX,'                                              : `expedida en ${lugar_exp},`,

        // PRIMERA — apto
        'Apartamento XXXX,'                                                    : `Apartamento ${apto},`,

        // COTIZACIÓN
        'CO-XXXX-2026'                                                         : cotizacion,

        // VALOR TOTAL (48 X)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX M/CTE ($XX.XXX.XXX)' : `${numLetras(vt)} (${fmtPesos(vt)})`,

        // PAGO 1: porcentaje (7 X) y valor congelación (34 X)
        'XXXXXXX por ciento XX%'                                               : pctFmt(pctCong),
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX M/CTE ($X.XXX.XXX)'               : `${numLetras(cong)} (${fmtPesos(cong)})`,

        // PAGO 2: porcentaje (9 X) y valor anticipo (48 X)
        'XXXXXXXXX por ciento XX%'                                             : pctFmt(pct2),
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX.XXX.XXXX)': `${numLetras(p2)} (${fmtPesos(p2)})`,

        // PAGOS 3-5: valores (6 X ya hardcoded como "veinte/diez")
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($X,XXX,XXX)'    : `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXPESOS ($X,XXX,XXX)'      : `${numLetras(p4)} (${fmtPesos(p4)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($X,XXX,XXX)'        : `${numLetras(p5)} (${fmtPesos(p5)})`,

        // VIGENCIA
        'MES DE NOVIEMBRE DE 2026 Y UN MES MAS'                               : fecha_cong.toUpperCase()+' Y UN MES MAS',

        // NOTIFICACIONES
        'Dirección: xxxxxxxx'                                                  : `Dirección: ${direccion}`,
        'Lugar: xxxxxxx'                                                       : `Lugar: ${lugar_exp}`,
        'Teléfono: xxxxxxx'                                                    : `Teléfono: ${telefono}`,
        'Email: xxxxxxxxxxx'                                                   : `Email: ${correo}`,

        // FECHA FIRMA
        'xx día del mes de xxxx'                                               : `${dia} día del mes de ${mes}`,

        // FIRMA
        'Nombre: xxxxxxxxxxx'                                                  : `Nombre: ${nombre}`,
        'Email: xxxxxxxxxxxxxx'                                                 : `Email: ${correo}`,
      };

    // ── Gran Central Con Bono ─────────────────────────────────────────────────
    } else if (tipo === 'gran_central_con_bono') {
      const BONO      = 40_247_135;
      const excedente = vt - BONO;

      reps = {
        // TÍTULO
        'CON-xxxxxxxx'                                                         : conNum,

        // IDENTIFICACIÓN — nombre (25 X), cédula con contexto (16 X), lugar (10 X)
        'XXXXXXXXXXXXXXXXXXXXXXXXX, mayor de edad'                             : `${nombre}, mayor de edad`,
        'ciudadanía No. XXXXXXXXXXXXXXXX expedida'                            : `ciudadanía No. ${cedula} expedida`,
        'expedida en XXXXXXXXXX,'                                              : `expedida en ${lugar_exp},`,

        // PRIMERA — apto
        'Apartamento XXXX,'                                                    : `Apartamento ${apto},`,

        // COTIZACIÓN
        'CO-XXXX-2026'                                                         : cotizacion,

        // VALOR EXCEDENTE (44 X — aparece 2 veces, en 2 párrafos distintos)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)' : `${numLetras(excedente)} (${fmtPesos(excedente)})`,

        // VIGENCIA
        'MES DE NOVIEMBRE DE 2026 Y UN MES MAS'                               : fecha_cong.toUpperCase()+' Y UN MES MAS',

        // NOTIFICACIONES
        'Dirección: xxxxxxxx'                                                  : `Dirección: ${direccion}`,
        'Lugar: xxxxxxx'                                                       : `Lugar: ${lugar_exp}`,
        'Teléfono: xxxxxxx'                                                    : `Teléfono: ${telefono}`,
        'Email: xxxxxxxxxxx'                                                   : `Email: ${correo}`,

        // FECHA FIRMA
        'xx día del mes de xxxx'                                               : `${dia} día del mes de ${mes}`,

        // FIRMA
        'Nombre: xxxxxxxxxxx'                                                  : `Nombre: ${nombre}`,
        'Email: xxxxxxxxxxxxxx'                                                 : `Email: ${correo}`,
      };
    }

    const resultBuffer = await replaceInDocx(docxBuffer, reps);

    // Nombre de archivo: "CON-0001-2026 JUAN CARLOS PÉREZ.docx"
    const filename = `${conNum} ${nombre.toUpperCase()}.docx`;

    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.status(200).send(resultBuffer);

  } catch (error) {
    console.error('[api/generar.js] ERROR:', error);
    res.status(500).json({error: error.message, stack: error.stack});
  }
};
