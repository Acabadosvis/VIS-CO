// netlify/functions/generar.js
'use strict';

const fs   = require('fs');
const path = require('path');
const JSZip = require('jszip'); // DEBE estar aquí arriba para que Netlify lo empaquete

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
// Word parte los placeholders largos en múltiples <w:r> con distinto rsid.
// Solución: para cada <w:p> extraemos el texto completo, aplicamos los reemplazos
// y ponemos todo el texto en el primer <w:t>, vaciando los demás.

function extractParaText(paraXml) {
  // Captura el contenido de TODOS los <w:t>...</w:t> del párrafo
  const matches = paraXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
  return matches.map(m => m.replace(/<[^>]*>/g, '')).join('');
}

function applyReplacements(text, sortedReps) {
  let result = text;
  for (const [k, v] of sortedReps) {
    if (result.includes(k)) {
      result = result.split(k).join(String(v));
    }
  }
  return result;
}

function replaceParagraphText(paraXml, sortedReps) {
  const original = extractParaText(paraXml);
  const updated  = applyReplacements(original, sortedReps);

  if (updated === original) return paraXml; // nada cambió

  // Poner todo el texto en el PRIMER <w:t>, vaciar el resto
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

  // Ordenar por longitud de clave descendente (claves largas primero = más específicas)
  const sorted = Object.entries(replacements)
    .sort((a, b) => b[0].length - a[0].length);

  // Procesar párrafo a párrafo con merge de runs
  xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, para => replaceParagraphText(para, sorted));

  zip.file('word/document.xml', xml);
  return zip.generateAsync({
    type: 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

// ── Handler principal ─────────────────────────────────────────────────────────

exports.handler = async (event) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin' : '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers: corsHeaders, body: '' };
  if (event.httpMethod !== 'POST')   return { statusCode: 405, headers: corsHeaders, body: 'Method not allowed' };

  try {
    const datos = JSON.parse(event.body);

    const {
      tipo,
      nombre       = '',
      cedula       = '',
      lugar_exp    = '',
      direccion    = '',
      correo       = '',
      telefono     = '',
      cotizacion   = '',
      valor_total,
      congelacion,
      fecha_cong   = '',   // Parágrafo 1: fecha vigencia cotización
      dia          = '',
      mes          = '',
      torre        = '',
      apto         = '',
      proyecto     = '',
    } = datos;

    // Validación básica
    if (!tipo || !nombre || !valor_total) {
      return {
        statusCode: 400,
        headers: corsHeaders,
        body: JSON.stringify({ error: 'Faltan campos obligatorios: tipo, nombre, valor_total' }),
      };
    }
    if (!fecha_cong) {
      return {
        statusCode: 400,
        headers: corsHeaders,
        body: JSON.stringify({ error: 'Falta la fecha de vigencia de la cotización (fecha_cong). Por favor ingrésala para continuar.' }),
      };
    }

    const minutaNames = {
      acabadosvis : 'Minuta clientes Acabadosvis.docx',
      visualizza  : 'Contrato clientes visualizza.docx',
      gran_central: 'CON - XXXX - 2026 NOMBRE CLIENTE Gran central.docx',
    };

    if (!minutaNames[tipo]) {
      return { statusCode: 400, headers: corsHeaders, body: JSON.stringify({ error: `Tipo desconocido: ${tipo}` }) };
    }

    const minutaPath = path.join(process.cwd(), minutaNames[tipo]);
    if (!fs.existsSync(minutaPath)) {
      return { statusCode: 500, headers: corsHeaders, body: JSON.stringify({ error: `Minuta no encontrada en: ${minutaPath}` }) };
    }

    const docxBuffer = fs.readFileSync(minutaPath);

    // ── Cálculos de valores ────────────────────────────────────────────────────
    const vt   = Number(valor_total);
    const cong = Number(congelacion) || 5_000_000;

    // Número limpio de cotización: "co-2502-2026" → "2502"
    const cotizNum = cotizacion
      .replace(/^[Cc][Oo][-\s]*/,  '')  // quita el prefijo CO-
      .replace(/[-\s]*20\d\d$/,     '')  // quita el año al final
      .replace(/-/g, '')                  // quita guiones internos
      .trim();

    // Número de contrato: "CON-2502-2026"
    const conNum = `CON-${cotizNum}-2026`;

    // Cotización formateada para el cuerpo del contrato: "CO - 2502 - 2026"
    const cotizFmt = `CO - ${cotizNum} - 2026`;

    let reps = {};

    // ── AcabadosVIS ──────────────────────────────────────────────────────────
    if (tipo === 'acabadosvis') {
      const resto   = vt - cong;
      const p2      = Math.round(resto * 0.50);
      const p3      = Math.round(resto * 0.20);
      const p4      = Math.round(resto * 0.20);
      const p5      = vt - cong - p2 - p3 - p4;
      const pctCong = Math.round((cong / vt) * 100);
      const pct2    = Math.round((p2   / vt) * 100);

      reps = {
        // ── Número de contrato (título) ─────────────────────────────
        // El 5-X aparece en "CON - XXXXX - 2026"
        'No. CON - XXXXX': `No. CON - ${cotizNum}`,

        // ── Identificación del cliente ──────────────────────────────
        // 20-X CONTEXTO 1: nombre del cliente (párrafo identificación)
        'XXXXXXXXXXXXXXXXXXXX mayor de edad': `${nombre} mayor de edad`,

        // 13-X con contexto para NO chocar con Teléfono "3XXXXXXXXXXXXX"
        'ciudadanía No. XXXXXXXXXXXXX expedida': `ciudadanía No. ${cedula} expedida`,

        // 10-X lugar expedición
        'expedida en XXXXXXXXXX,': `expedida en ${lugar_exp},`,

        // ── Objeto del contrato (PRIMERA) ───────────────────────────
        // 20-X CONTEXTO 2: ubicación del inmueble → proyecto, NO nombre cliente
        'ubicado en XXXXXXXXXXXXXXXXXXXX Torre': `ubicado en ${proyecto} Torre`,

        // Torre y Apartamento
        'Torre XX Apartamento XXXX': `Torre ${torre} Apartamento ${apto}`,

        // ── Referencias de cotización en cuerpo ─────────────────────
        // 7-X → número cotización (en PRIMERA, SEGUNDA)
        'CO - XXXXXXX': cotizFmt,

        // 6-X → número cotización (en QUINTA con guión largo "CO – XXXXXX")
        'CO – XXXXXX': `CO – ${cotizNum}`,  // – = U+2013

        // ── Valor total (33-X intacto en un solo run) ───────────────
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)':
          `${numLetras(vt)} (${fmtPesos(vt)})`,

        // ── Pagos: porcentajes ──────────────────────────────────────
        // 8-X pago 2 (DEBE ir antes que 6-X)
        'XXXXXXXX por ciento XX%': `${pctTxt(pct2)} por ciento ${pct2}%`,

        // 6-X pago 1
        'XXXXXX por ciento XX%': `${pctTxt(pctCong)} por ciento ${pctCong}%`,

        // ── Pagos: valores en letras y números ──────────────────────
        // 31-X pago 2 (anticipo)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX.XXX.XXX)':
          `${numLetras(p2)} (${fmtPesos(p2)})`,

        // 59-X pago 3 (avance obra civil)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)':
          `${numLetras(p3)} (${fmtPesos(p3)})`,

        // 62-X pago 4 (70% carpintería)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)':
          `${numLetras(p4)} (${fmtPesos(p4)})`,

        // 61-X pago 5 (final)
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS MCTE ($X.XXX.XXX)':
          `${numLetras(p5)} (${fmtPesos(p5)})`,

        // ── Parágrafo primero: vigencia cotización ──────────────────
        // 9-X → fecha de vigencia ingresada por el usuario en el form
        'hasta el XXXXXXXXX y': `hasta el ${fecha_cong} y`,

        // ── Datos del cliente (sección notificaciones) ──────────────
        'Dirección: xxxxxxxxxxxxxx'   : `Dirección: ${direccion}`,
        'Lugar: XXXXXXXXXX'           : `Lugar: ${lugar_exp}`,
        'Teléfono: 3XXXXXXXXXXXXX'    : `Teléfono: ${telefono}`,

        // 20-X CONTEXTO 3: email en sección notificaciones
        'Email: XXXXXXXXXXXXXXXXXXXX' : `Email: ${correo}`,

        // ── Fecha de suscripción ────────────────────────────────────
        // 'día' usa la í literal (NO &#237;)
        'XX día del mes de xxxxx': `${dia} día del mes de ${mes}`,

        // ── Bloque de firmas ────────────────────────────────────────
        // 17-X nombre en firma contratante
        'Nombre: XXXXXXXXXXXXXXXXX'  : `Nombre: ${nombre}`,
        // 21-X email en firma contratante
        'Email: XXXXXXXXXXXXXXXXXXXXX': `Email: ${correo}`,
      };

    // ── Visualizza ───────────────────────────────────────────────────────────
    } else if (tipo === 'visualizza') {
      const p1 = Math.round(vt * 0.50);
      const p2 = Math.round(vt * 0.20);
      const p3 = Math.round(vt * 0.20);
      const p4 = vt - p1 - p2 - p3;

      reps = {
        'CON-XXXXX-26'                                                          : conNum,
        'XXXXXXXXXXXXXXXXXXX'                                                   : nombre,
        'XXXXXXXXXXX'                                                           : cedula,
        'XXXXXXXXXX'                                                            : lugar_exp,
        'XXXXXXXXXXXXXXXXXXXXXXX'                                               : `${proyecto} ${torre} Apartamento ${apto}`,
        'CO-XXXXXX-25'                                                          : cotizacion,
        'CO-XXXXX-25'                                                           : cotizacion,
        'CO-XXXXXXXX'                                                           : cotizacion,
        "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS M/CTE ($XX'XXX.XXX)"              : `${numLetras(vt)} (${fmtPesos(vt)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)'                    : `${numLetras(p1)} (${fmtPesos(p1)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)'        : `${numLetras(p2)} (${fmtPesos(p2)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($XX,XXX,XXX)'              : `${numLetras(p3)} (${fmtPesos(p3)})`,
        'XXXXXXXXXXXXXXXXXXXXXXXXXX PESOS ($X,XXX,XXX)'                        : `${numLetras(p4)} (${fmtPesos(p4)})`,
        'MES DE XXXXXX DE 202X Y UN MES MAS'                                   : fecha_cong.toUpperCase() + ' Y UN MES MAS',
        '(23) día del mes de julio'                                             : `(${dia}) día del mes de ${mes}`,
        'Dirección: xxxxxxxxxxxxxx'                                             : `Dirección: ${direccion}`,
        'Lugar: XXXXXXXXXX'                                                     : `Lugar: ${lugar_exp}`,
        'Teléfono: 3XXXXXXXXXXXXX'                                              : `Teléfono: ${telefono}`,
        'Email: XXXXXXXXXXXXXXXXXXXX'                                            : `Email: ${correo}`,
        'Nombre: XXXXXXXXXXXXXXXXX'                                             : `Nombre: ${nombre}`,
        'Email: XXXXXXXXXXXXXXXXXXXXX'                                           : `Email: ${correo}`,
        'XX día del mes de xxxxx'                                               : `${dia} día del mes de ${mes}`,
      };

    // ── Gran Central ─────────────────────────────────────────────────────────
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
        'Teléfono: 3XXXXXXXXXXXXX'                                                        : `Teléfono: ${telefono}`,
        'Email: XXXXXXXXXXXXXXXXXXXX'                                                      : `Email: ${correo}`,
        'Nombre: XXXXXXXXXXXXXXXXX'                                                        : `Nombre: ${nombre}`,
        'Email: XXXXXXXXXXXXXXXXXXXXX'                                                      : `Email: ${correo}`,
        'XX día del mes de xxxxx'                                                          : `${dia} día del mes de ${mes}`,
      };
    }

    const resultBuffer = await replaceInDocx(docxBuffer, reps);
    const base64 = resultBuffer.toString('base64');

    return {
      statusCode: 200,
      headers: {
        ...corsHeaders,
        'Content-Type'       : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${conNum}.docx"`,
      },
      body           : base64,
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
