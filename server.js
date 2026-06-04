/* ══════════════════════════════════════════════════════════════════════════
   MICROSERVICIO PDF — Railway (con LibreOffice)
   Recibe los datos del presupuesto, genera el mapa satélite, modifica el xlsx
   y lo convierte a PDF con LibreOffice. NO usa ConvertAPI.
   La web principal (Hostinger) reenvía aquí la petición y devuelve el PDF.
   ══════════════════════════════════════════════════════════════════════════ */

const express  = require('express');
const path     = require('path');
const fs       = require('fs');
const os       = require('os');
const https    = require('https');
const http     = require('http');
const { spawnSync } = require('child_process');
const JSZip    = require('jszip');
const Jimp     = require('jimp');
const app      = express();
const PORT     = process.env.PORT || 3000;

/* ── MAPA SATÉLITE — ESRI World Imagery (gratuito) ───────────────────────── */

function fetchBuffer(url) {
  return new Promise((resolve, reject) => {
    const client = url.startsWith('https') ? https : http;
    const req = client.get(url, {
      headers: { 'User-Agent': 'AutoconsumoEconomico/1.0 (presupuesto solar)' }
    }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchBuffer(res.headers.location).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) return reject(new Error(`HTTP ${res.statusCode} para ${url}`));
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    });
    req.on('error', reject);
    req.setTimeout(5000, () => { req.destroy(); reject(new Error('Timeout tile')); });
  });
}

function latLngToTile(lat, lng, zoom) {
  const n = Math.pow(2, zoom);
  const x = Math.floor((lng + 180) / 360 * n);
  const latRad = lat * Math.PI / 180;
  const y = Math.floor((1 - Math.log(Math.tan(latRad) + 1 / Math.cos(latRad)) / Math.PI) / 2 * n);
  return { x, y };
}

function fetchEsriTile(z, tileX, tileY) {
  const url = `https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/${z}/${tileY}/${tileX}`;
  return fetchBuffer(url);
}

async function buildSatelliteImage(lat, lng) {
  const ZOOM = 18, TSIZE = 256, GRID = 3, TOTAL = TSIZE * GRID;
  const center = latLngToTile(lat, lng, ZOOM);
  const tilePromises = [];
  for (let dy = -1; dy <= 1; dy++)
    for (let dx = -1; dx <= 1; dx++)
      tilePromises.push(fetchEsriTile(ZOOM, center.x + dx, center.y + dy));
  const tileBuffers = await Promise.all(tilePromises);
  const canvas = await Jimp.create(TOTAL, TOTAL, 0x000000ff);
  let idx = 0;
  for (let row = 0; row < GRID; row++)
    for (let col = 0; col < GRID; col++) {
      const tile = await Jimp.read(tileBuffers[idx++]);
      canvas.composite(tile, col * TSIZE, row * TSIZE);
    }
  const cx = TOTAL / 2, cy = TOTAL / 2;
  canvas.scan(cx - 16, cy - 16, 33, 33, function (px, py) {
    const d = Math.sqrt((px - cx) ** 2 + (py - cy) ** 2);
    if (d <= 14)
      this.setPixelColor(d <= 10 ? Jimp.rgbaToInt(220,50,50,255) : Jimp.rgbaToInt(255,255,255,255), px, py);
  });
  canvas.resize(300, 300);
  return canvas.getBufferAsync(Jimp.MIME_PNG);
}

async function geocodeAddress(address) {
  try {
    const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(address)}&format=json&limit=1`;
    const buf = await fetchBuffer(url);
    const results = JSON.parse(buf.toString());
    if (!results.length) return null;
    return { lat: parseFloat(results[0].lat), lng: parseFloat(results[0].lon) };
  } catch { return null; }
}

/* ── CONVERSIÓN XLSX → PDF con LibreOffice ───────────────────────────────── */

// Crea un perfil de LibreOffice con recálculo automático al cargar (OOXML/ODF
// RecalcMode=0 = "Siempre"). Sin esto, LibreOffice headless NO recalcula las
// fórmulas y el PDF saldría con los valores cacheados (título/precio viejos).
function makeLoProfile() {
  const base = path.join(os.tmpdir(), 'lo_profile_' + Date.now() + '_' + Math.random().toString(36).slice(2));
  const userDir = path.join(base, 'user');
  fs.mkdirSync(userDir, { recursive: true });
  const xcu = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>`;
  fs.writeFileSync(path.join(userDir, 'registrymodifications.xcu'), xcu);
  return base;
}

function xlsxToPdf(xlsxBuffer, tmpXlsx, tmpPdf) {
  fs.writeFileSync(tmpXlsx, xlsxBuffer);
  const profileBase = makeLoProfile();
  const profileUrl  = 'file://' + profileBase.replace(/\\/g, '/').replace(/^([A-Za-z]:)/, '/$1');
  let loResult;
  if (process.platform === 'win32') {
    const soffice = 'C:\\Program Files\\LibreOffice\\program\\soffice.exe';
    loResult = spawnSync('powershell.exe',
      ['-Command', `& '${soffice}' --headless --norestore '-env:UserInstallation=${profileUrl}' --convert-to pdf --outdir '${os.tmpdir()}' '${tmpXlsx}'`],
      { timeout: 60000, windowsHide: true });
  } else {
    loResult = spawnSync('soffice',
      ['--headless','--norestore',`-env:UserInstallation=${profileUrl}`,
       '--convert-to','pdf','--outdir', os.tmpdir(), tmpXlsx],
      { timeout: 60000, env: { ...process.env, HOME: os.tmpdir() } });
  }
  if (loResult.error)
    throw new Error(`LibreOffice no se pudo ejecutar: ${loResult.error.message}`);
  if (loResult.status !== 0)
    throw new Error(`LibreOffice status ${loResult.status}: stderr=${(loResult.stderr||'').toString()} stdout=${(loResult.stdout||'').toString()}`);
  return fs.readFileSync(tmpPdf);
}

// Aplana los rellenos de degradado de los gráficos a color sólido (solo para
// LibreOffice, que los pinta negros). Devuelve un nuevo buffer xlsx.
async function degradeChartsToSolid(xlsxBuffer) {
  const z = await JSZip.loadAsync(xlsxBuffer);
  const charts = Object.keys(z.files).filter(f => /^xl\/charts\/chart\d+\.xml$/.test(f));
  for (const cf of charts) {
    let cxml = await z.file(cf).async('string');
    cxml = cxml.replace(/<a:gradFill[\s\S]*?<\/a:gradFill>/g, (g) => {
      const gs = g.match(/<a:gs\b[^>]*>([\s\S]*?)<\/a:gs>/);
      return `<a:solidFill>${gs ? gs[1] : '<a:srgbClr val="4472C4"/>'}</a:solidFill>`;
    });
    z.file(cf, cxml);
  }
  return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/* ── CONVERSIÓN XLSX → PDF con Aspose.Cells Cloud (alta fidelidad) ──────────
   Aspose NO recalcula en la conversión directa, así que usamos el flujo con
   almacenamiento: subir → CalculateFormula → convertir a PDF → borrar.        */

function asposeRequest(method, urlStr, token, bodyBuffer, contentType) {
  return new Promise((resolve, reject) => {
    const u = new URL(urlStr);
    const headers = {};
    if (token)       headers['Authorization']  = 'Bearer ' + token;
    if (contentType) headers['Content-Type']   = contentType;
    headers['Content-Length'] = bodyBuffer ? bodyBuffer.length : 0;
    const req = https.request(
      { hostname: u.hostname, path: u.pathname + u.search, method, headers },
      (r) => {
        const chunks = [];
        r.on('data', c => chunks.push(c));
        r.on('end', () => resolve({
          status: r.statusCode,
          body: Buffer.concat(chunks),
          contentType: r.headers['content-type'] || ''
        }));
        r.on('error', reject);
      }
    );
    req.on('error', reject);
    req.setTimeout(60000, () => { req.destroy(); reject(new Error('Timeout Aspose ' + method)); });
    if (bodyBuffer) req.write(bodyBuffer);
    req.end();
  });
}

let asposeTokenCache = { token: null, exp: 0 };
async function asposeGetToken(id, secret) {
  if (asposeTokenCache.token && Date.now() < asposeTokenCache.exp) return asposeTokenCache.token;
  const body = Buffer.from(`grant_type=client_credentials&client_id=${id}&client_secret=${secret}`);
  const res = await asposeRequest('POST', 'https://api.aspose.cloud/connect/token', null, body, 'application/x-www-form-urlencoded');
  if (res.status !== 200) throw new Error(`Aspose token ${res.status}: ${res.body.toString().slice(0,200)}`);
  const data = JSON.parse(res.body.toString());
  asposeTokenCache = { token: data.access_token, exp: Date.now() + ((data.expires_in || 3600) - 60) * 1000 };
  return data.access_token;
}

async function convertWithAspose(xlsxBuffer, id, secret) {
  const token = await asposeGetToken(id, secret);
  const name  = `pres_${Date.now()}_${Math.random().toString(36).slice(2)}.xlsx`;
  const out   = name.replace(/\.xlsx$/, '.pdf');
  const base  = 'https://api.aspose.cloud/v3.0/cells';
  try {
    let r = await asposeRequest('PUT', `${base}/storage/file/${name}`, token, xlsxBuffer, 'application/octet-stream');
    if (r.status !== 200) throw new Error(`Aspose upload ${r.status}: ${r.body.toString().slice(0,200)}`);
    r = await asposeRequest('POST', `${base}/${name}/calculateformula`, token, null);
    if (r.status !== 200) throw new Error(`Aspose calculate ${r.status}: ${r.body.toString().slice(0,200)}`);
    // Salto de página horizontal ANTES de la cabecera "Detalles técnicos" del mapa
    // (celda A123) usando la API nativa de Aspose: los saltos del archivo .xlsx los
    // IGNORA, pero este sí lo respeta → pág. 3 = cabecera + mapa juntos (no mapa huérfano).
    r = await asposeRequest('PUT', `${base}/${name}/worksheets/Factura/horizontalpagebreaks?cellname=A123`, token, null);
    if (r.status !== 200) console.warn(`[ASPOSE] salto de página A123 no aplicado (${r.status})`);
    // SaveAs con PdfSaveOptions: OnePagePerSheet=false para que Aspose pagine por el
    // page setup del Excel (scale 73% → 3 páginas A4) en vez de comprimir todo en 1.
    const pdfOpts = Buffer.from(JSON.stringify({ SaveFormat: 'pdf', OnePagePerSheet: false }));
    r = await asposeRequest('POST', `${base}/${name}/SaveAs?newfilename=${out}`, token, pdfOpts, 'application/json');
    if (r.status !== 200) throw new Error(`Aspose saveas ${r.status}: ${r.body.toString().slice(0,200)}`);
    r = await asposeRequest('GET', `${base}/storage/file/${out}`, token, null);
    if (r.status !== 200 || r.body.slice(0, 4).toString() !== '%PDF')
      throw new Error(`Aspose download ${r.status}: ${r.body.toString().slice(0,200)}`);
    return r.body;
  } finally {
    asposeRequest('DELETE', `${base}/storage/file/${name}`, token, null).catch(() => {});
    asposeRequest('DELETE', `${base}/storage/file/${out}`,  token, null).catch(() => {});
  }
}

/* ── INYECCIÓN DE DATOS EN LA HOJA FACTURA Y LECTURA DE RESULTADOS ─────────── */

const _cellRe   = ref => new RegExp(`<c r="${ref}"((?:[^>]*?))(?:/>|>[\\s\\S]*?</c>)`);
const _styleOf  = attrs => ((/ s="\d+"/.exec(attrs) || [''])[0]);
const _xmlEsc   = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const _colNum   = c => { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
const _setNum   = (x, ref, val) => x.replace(_cellRe(ref), (m, a) => `<c r="${ref}"${_styleOf(a)}><v>${val}</v></c>`);
const _setStr   = (x, ref, str) => x.replace(_cellRe(ref), (m, a) => `<c r="${ref}"${_styleOf(a)} t="inlineStr"><is><t xml:space="preserve">${_xmlEsc(str)}</t></is></c>`);
// Como _setNum pero inserta la celda si no existe (P9 y M16 vienen vacías).
const _setCellNum = (x, ref, val) => {
  if (_cellRe(ref).test(x)) return _setNum(x, ref, val);
  const [, col, row] = ref.match(/^([A-Z]+)(\d+)$/);
  const rowRe = new RegExp(`(<row r="${row}"[^>]*>)([\\s\\S]*?)(</row>)`);
  return x.replace(rowRe, (m, open, body, close) => {
    const newCell = `<c r="${ref}"><v>${val}</v></c>`;
    let insertAt = body.length;
    for (const cm of body.matchAll(/<c r="([A-Z]+)\d+"[\s\S]*?(?:\/>|<\/c>)/g)) {
      if (_colNum(cm[1]) > _colNum(col)) { insertAt = cm.index; break; }
    }
    return open + body.slice(0, insertAt) + newCell + body.slice(insertAt) + close;
  });
};

// Normaliza los datos del formulario a los valores que espera el Excel.
function normalizeInputs(body) {
  const { kWp, numPaneles, consumo, cliente, tarifa, fase, bateria, estructura } = body;
  return {
    kWp, numPaneles,
    consumoVal:     (consumo && consumo > 0) ? consumo : kWp * 1000,
    bateriaVal:     bateria ? 1 : 0,
    excelDate:      Math.round(Date.now() / 86400000 + 25569),
    faseExcel:      (fase && /tri/i.test(fase)) ? 'trifásico' : 'monofásico',
    clienteExcel:   (cliente && /empresa/i.test(cliente)) ? 'empresa' : 'particular',
    tarifaExcel:    tarifa || '2.0 TD',
    estructuraExcel:(estructura === 'inclinada/solarblocs') ? 'inclinada/solarblocs' : 'coplanar'
  };
}

// Escribe los datos en la hoja Factura (sheet1) de un zip ya cargado.
async function injectInputs(zip, v) {
  let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  xml = _setNum(xml, 'N8',  v.consumoVal);       // Consumo kWh/año
  xml = _setCellNum(xml, 'M16', v.numPaneles);   // Nº paneles → título kWp + precio
  xml = _setCellNum(xml, 'P9',  v.kWp);          // Potencia inversor kW
  xml = _setStr(xml, 'O9',  v.faseExcel);        // monofásico / trifásico
  xml = _setStr(xml, 'O12', v.estructuraExcel);  // coplanar / inclinada-solarblocs
  xml = _setNum(xml, 'I31', v.bateriaVal);       // Batería (0/1)
  xml = _setNum(xml, 'C9',  v.excelDate);        // Fecha
  xml = _setStr(xml, 'N6',  v.clienteExcel);     // Particular/empresa → IVA
  xml = _setStr(xml, 'N7',  v.tarifaExcel);      // Tarifa
  zip.file('xl/worksheets/sheet1.xml', xml);
}

// Recalcula un xlsx con LibreOffice (convierte a xlsx con el perfil de recálculo)
// y devuelve el buffer recalculado. Gratis y sin Aspose.
function recalcXlsx(xlsxBuffer, tmpIn, outDir) {
  fs.writeFileSync(tmpIn, xlsxBuffer);
  const profileBase = makeLoProfile();
  const profileUrl  = 'file://' + profileBase.replace(/\\/g, '/').replace(/^([A-Za-z]:)/, '/$1');
  let r;
  if (process.platform === 'win32') {
    const soffice = 'C:\\Program Files\\LibreOffice\\program\\soffice.exe';
    r = spawnSync('powershell.exe',
      ['-Command', `& '${soffice}' --headless --norestore '-env:UserInstallation=${profileUrl}' --convert-to xlsx --outdir '${outDir}' '${tmpIn}'`],
      { timeout: 60000, windowsHide: true });
  } else {
    r = spawnSync('soffice',
      ['--headless','--norestore',`-env:UserInstallation=${profileUrl}`,'--convert-to','xlsx','--outdir', outDir, tmpIn],
      { timeout: 60000, env: { ...process.env, HOME: os.tmpdir() } });
  }
  if (r.error)      throw new Error(`LibreOffice (recalc) no se pudo ejecutar: ${r.error.message}`);
  if (r.status !== 0) throw new Error(`LibreOffice (recalc) status ${r.status}: ${(r.stderr||'').toString()}`);
  const outFile = path.join(outDir, path.basename(tmpIn));
  return fs.readFileSync(outFile);
}

// Lee valores de celdas (resueltos) de un xlsx recalculado.
async function readCells(xlsxBuffer, refs) {
  const zip = await JSZip.loadAsync(xlsxBuffer);
  const sst = [];
  const ssx = zip.file('xl/sharedStrings.xml');
  if (ssx) { const t = await ssx.async('string');
    for (const m of t.matchAll(/<si>([\s\S]*?)<\/si>/g))
      sst.push([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join('')); }
  const cache = {};
  const result = {};
  for (const { key, sheet, ref } of refs) {
    if (!cache[sheet]) cache[sheet] = await zip.file(sheet).async('string');
    const m = cache[sheet].match(new RegExp(`<c r="${ref}"([^>]*)>(?:<f[^>]*>[\\s\\S]*?<\\/f>)?(?:<v>([\\s\\S]*?)<\\/v>|<is>([\\s\\S]*?)<\\/is>)<\\/c>`));
    let val = null;
    if (m) {
      if (/t="s"/.test(m[1])) val = sst[+m[2]];
      else if (m[3] !== undefined) val = [...m[3].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join('');
      else val = parseFloat(m[2]);
    }
    result[key] = val;
  }
  return result;
}

// Nombre de columna a partir de su número (1=A, 13=M, 27=AA...).
const _colName = n => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

// Lee un valor numérico de una celda concreta de un XML de hoja ya cargado.
function _readNumFromSheet(sheetXml, ref) {
  const m = sheetXml.match(new RegExp(`<c r="${ref}"[^>]*>(?:<f[^>]*>[\\s\\S]*?<\\/f>)?<v>([\\s\\S]*?)<\\/v>`));
  return m ? parseFloat(m[1]) : NaN;
}

// Congela como CONSTANTES las celdas de la curva (filas 118-130 × cols M:AJ de
// "Datos partida") usando los valores ya calculados por LibreOffice. Al quedar sin
// fórmula, el calculateformula de Aspose no las puede recalcular a 0 → la curva se dibuja.
const CURVA_FILAS = [118, 119, 120, 123, 124, 127, 130];
const CURVA_COLS  = (() => { const a = []; for (let n = 13; n <= 36; n++) a.push(_colName(n)); return a; })(); // M..AJ
async function congelarCurva(zip, recalcedBuffer) {
  const recZip = await JSZip.loadAsync(recalcedBuffer);
  const s2rec  = await recZip.file('xl/worksheets/sheet2.xml').async('string'); // Datos partida
  let s2 = await zip.file('xl/worksheets/sheet2.xml').async('string');
  let n = 0;
  for (const row of CURVA_FILAS) {
    for (const col of CURVA_COLS) {
      const ref = `${col}${row}`;
      const val = _readNumFromSheet(s2rec, ref);
      if (!isFinite(val)) continue;
      s2 = _setCellNum(s2, ref, val);
      n++;
    }
  }
  zip.file('xl/worksheets/sheet2.xml', s2);
  return n;
}

// Localiza la imagen del MAPA dentro del xlsx (la PNG más grande del dibujo de la
// hoja Factura). El número de imagen cambia según la plantilla, así que NO se puede
// asumir image1.png; hay que detectarla.
async function findMapImagePath(zip) {
  try {
    const sheetRels = zip.file('xl/worksheets/_rels/sheet1.xml.rels');
    if (!sheetRels) return null;
    const rx = await sheetRels.async('string');
    const dm = rx.match(/Target="([^"]*drawings\/drawing\d+\.xml)"/);
    if (!dm) return null;
    const drawingName = dm[1].split('/').pop();
    const drawRels = zip.file(`xl/drawings/_rels/${drawingName}.rels`);
    if (!drawRels) return null;
    const drx = await drawRels.async('string');
    const pngs = [...drx.matchAll(/Target="([^"]+\.png)"/gi)].map(m => m[1].replace(/^\.\.\//, 'xl/'));
    let best = null, bestSize = -1;
    for (const p of pngs) {
      const f = zip.file(p);
      if (!f) continue;
      const b = await f.async('nodebuffer');
      if (b.length > bestSize) { bestSize = b.length; best = p; }
    }
    return best;
  } catch { return null; }
}

/* ── EXPRESS ─────────────────────────────────────────────────────────────── */

app.use(express.json({ limit: '50mb' }));
app.disable('x-powered-by');

app.get('/health', (req, res) => res.json({ ok: true }));

app.post('/api/presupuesto-excel', async (req, res) => {
  const ts      = Date.now();
  const tmpDir  = os.tmpdir();
  const tmpXlsx = path.join(tmpDir, `presupuesto_${ts}.xlsx`);
  const tmpPdf  = path.join(tmpDir, `presupuesto_${ts}.pdf`);

  try {
    const { kWp, address, lat, lng } = req.body;
    const v = normalizeInputs(req.body);

    // Mapa satélite en paralelo con la preparación del xlsx
    const latNum = (lat !== null && lat !== undefined && lat !== '') ? parseFloat(lat) : NaN;
    const lngNum = (lng !== null && lng !== undefined && lng !== '') ? parseFloat(lng) : NaN;
    const mapPromise = (async () => {
      try {
        let coords = null;
        if (!isNaN(latNum) && !isNaN(lngNum)) coords = { lat: latNum, lng: lngNum };
        else if (address && address.trim()) coords = await geocodeAddress(address.trim());
        if (coords) {
          console.log('[MAPA] Generando satelite para', coords);
          const buf = await buildSatelliteImage(coords.lat, coords.lng);
          console.log('[MAPA] PNG generado:', buf.length, 'bytes');
          return buf;
        }
      } catch (mapErr) { console.warn('[MAPA] Error:', mapErr.message); }
      return null;
    })();

    const srcData = fs.readFileSync(path.join(__dirname, 'exelhome.xlsx'));
    const zip = await JSZip.loadAsync(srcData);

    // Inyectar los datos del formulario en la hoja Factura (celdas de Toño).
    await injectInputs(zip, v);

    // Recalcular con LibreOffice (calcula bien la curva) y CONGELAR sus celdas como
    // constantes. Así Aspose, al hacer calculateformula, no puede recalcularlas a 0
    // → la curva CONSUMO/PRODUCCIÓN se dibuja (Aspose ignora la numCache y recomputa).
    try {
      let wbInj = await zip.file('xl/workbook.xml').async('string');
      wbInj = wbInj.replace(/<calcPr([^>]*)\/>/, (m, a) =>
        /fullCalcOnLoad/.test(a) ? m : `<calcPr${a} fullCalcOnLoad="1"/>`);
      zip.file('xl/workbook.xml', wbInj);
      const injectedBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      const recalcDir = path.join(tmpDir, `recalc_${ts}`);
      fs.mkdirSync(recalcDir, { recursive: true });
      const recalcIn = path.join(tmpDir, `recalc_${ts}.xlsx`);
      const recalcedBuf = recalcXlsx(injectedBuf, recalcIn, recalcDir);
      const nCong = await congelarCurva(zip, recalcedBuf);
      console.log('[CURVA] celdas congeladas:', nCong);
      fs.unlink(recalcIn, () => {});
      fs.rm(recalcDir, { recursive: true, force: true }, () => {});
    } catch (curvaErr) {
      console.warn('[CURVA] No se pudo congelar la curva (sigue sin ella):', curvaErr.message);
    }

    // Aspose interpreta fitToWidth="0" como "encajar TODO en 1 página" y comprime las
    // 3 hojas en una. Excel/LibreOffice lo ignoran (no hay fitToPage). Lo quitamos para
    // que Aspose pagine por el scale (73%) → 3 páginas A4 como el diseño de Toño.
    {
      let s1 = await zip.file('xl/worksheets/sheet1.xml').async('string');
      s1 = s1.replace(/<pageSetup\b[^>]*\/>/, (tag) =>
        tag.replace(/\s+fitToWidth="\d+"/g, '').replace(/\s+fitToHeight="\d+"/g, ''));
      zip.file('xl/worksheets/sheet1.xml', s1);

      // Recortar el área de impresión (J194 → J166): la imagen del mapa termina en la
      // fila 166; las filas 167-194 son vacías y crean el hueco blanco bajo el mapa.
      // (NO recortar por debajo de 166 o se aplasta/corta el mapa.)
      let wbA = await zip.file('xl/workbook.xml').async('string');
      wbA = wbA.replace(/Factura!\$A\$1:\$J\$194/, () => 'Factura!$A$1:$J$166');
      zip.file('xl/workbook.xml', wbA);
    }

    // Ocultar hojas auxiliares (se necesitan para los cálculos, pero no deben
    // imprimirse). Robusto: con o sin atributo state previo.
    let wbXml = await zip.file('xl/workbook.xml').async('string');
    for (const sheetName of ['Datos partida', 'Calculos', 'Hoja1']) {
      wbXml = wbXml.replace(new RegExp(`<sheet name="${sheetName}"[^>]*?/>`), (tag) =>
        /state="/.test(tag)
          ? tag.replace(/state="[^"]*"/, 'state="veryHidden"')
          : tag.replace(/\/>\s*$/, ' state="veryHidden"/>')
      );
    }
    // Forzar recálculo completo al abrir (refuerzo del perfil de LibreOffice)
    wbXml = wbXml.replace(/<calcPr([^>]*)\/>/, (m, a) =>
      /fullCalcOnLoad/.test(a) ? m : `<calcPr${a} fullCalcOnLoad="1"/>`);
    zip.file('xl/workbook.xml', wbXml);

    const mapPngBuffer = await mapPromise;
    if (mapPngBuffer) {
      const mapPath = (await findMapImagePath(zip)) || 'xl/media/image2.png';
      zip.file(mapPath, mapPngBuffer);
      console.log('[MAPA]', mapPath, 'sustituida en xlsx');
    }

    // Los gráficos toman sus datos de "Datos partida", que ocultamos para que el
    // PDF salga en 3 páginas. Con plotVisOnly=1 no dibujarían (datos en hoja oculta)
    // → curvas planas. Lo ponemos a 0 para que se dibujen igualmente.
    const allCharts = Object.keys(zip.files).filter(f => /^xl\/charts\/chart\d+\.xml$/.test(f));
    for (const cf of allCharts) {
      let cxml = await zip.file(cf).async('string');
      cxml = cxml.replace(/<c:plotVisOnly val="1"\/>/g, '<c:plotVisOnly val="0"/>');
      zip.file(cf, cxml);
    }

    // xlsx con los gráficos INTACTOS (degradados incluidos). Aspose los dibuja bien;
    // para LibreOffice se aplanan a sólido justo antes de convertir.
    const xlsxBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

    const useAspose = !!(process.env.ASPOSE_CLIENT_ID && process.env.ASPOSE_CLIENT_SECRET);
    let pdfBytes;
    if (useAspose) {
      console.log('[PDF] Convirtiendo con Aspose (recálculo + gráficos fieles)...');
      // Si Aspose está configurado y falla, NO servimos el PDF feo de LibreOffice (mala
      // imagen para el cliente): devolvemos error para que reintente. El respaldo
      // LibreOffice solo se usa cuando Aspose NO está configurado (entorno local/dev).
      try {
        pdfBytes = await convertWithAspose(xlsxBuf, process.env.ASPOSE_CLIENT_ID, process.env.ASPOSE_CLIENT_SECRET);
      } catch (aspErr) {
        console.error('[PDF] Aspose falló:', aspErr.message);
        if (!res.headersSent) res.status(502).json({ error: 'El generador de PDF no está disponible ahora mismo. Inténtalo de nuevo en unos minutos.' });
        return;
      }
    } else {
      console.log('[PDF] Convirtiendo con LibreOffice...');
      pdfBytes = xlsxToPdf(await degradeChartsToSolid(xlsxBuf), tmpXlsx, tmpPdf);
    }
    console.log('[PDF] PDF generado:', pdfBytes.length, 'bytes');

    const nombre = `Presupuesto_Fotovoltaico_${kWp}kWp.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${nombre}"`);
    res.setHeader('Content-Length', pdfBytes.length);
    res.end(pdfBytes);

  } catch (err) {
    console.error('Error generando PDF:', err);
    if (!res.headersSent) res.status(500).json({ error: 'No se pudo generar el presupuesto.' });
  } finally {
    fs.unlink(tmpXlsx, () => {});
    fs.unlink(tmpPdf,  () => {});
  }
});

/* ── POST /api/calcular — números REALES del Excel para el panel de la web ───
   Inyecta los datos, recalcula con LibreOffice (gratis) y devuelve los valores
   exactos (precio, IRPF, neto, amortización, ahorro). Así la web = el PDF.      */
app.post('/api/calcular', async (req, res) => {
  const ts      = Date.now();
  const tmpXlsx = path.join(os.tmpdir(), `calc_${ts}.xlsx`);
  const outDir  = path.join(os.tmpdir(), `calcout_${ts}`);
  try {
    if (!req.body || !req.body.kWp || !req.body.numPaneles)
      return res.status(400).json({ error: 'Faltan datos' });

    const v = normalizeInputs(req.body);
    const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, 'exelhome.xlsx')));
    await injectInputs(zip, v);
    let wbXml = await zip.file('xl/workbook.xml').async('string');
    wbXml = wbXml.replace(/<calcPr([^>]*)\/>/, (m, a) =>
      /fullCalcOnLoad/.test(a) ? m : `<calcPr${a} fullCalcOnLoad="1"/>`);
    zip.file('xl/workbook.xml', wbXml);
    const xlsxBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

    fs.mkdirSync(outDir, { recursive: true });
    const recalced = recalcXlsx(xlsxBuf, tmpXlsx, outDir);

    const F = 'xl/worksheets/sheet1.xml';     // Factura
    const D = 'xl/worksheets/sheet2.xml';     // Datos partida
    const c = await readCells(recalced, [
      { key: 'precioConIva',  sheet: F, ref: 'I36' },
      { key: 'base',          sheet: F, ref: 'J33' },
      { key: 'iva',           sheet: F, ref: 'J34' },
      { key: 'deduccionIRPF', sheet: F, ref: 'F37' },
      { key: 'precioNeto',    sheet: F, ref: 'I40' },
      { key: 'amortizacion',  sheet: D, ref: 'J13' },
      { key: 'ahorroAnual',   sheet: D, ref: 'W87' }
    ]);

    const round = n => (typeof n === 'number' && isFinite(n)) ? Math.round(n) : null;
    res.json({
      precioConIva:  round(c.precioConIva),
      base:          round(c.base),
      iva:           round(c.iva),
      deduccionIRPF: round(c.deduccionIRPF),
      precioNeto:    round(c.precioNeto),
      ahorroAnual:   round(c.ahorroAnual),
      ahorroMensual: round((c.ahorroAnual || 0) / 12),
      amortizacion:  (typeof c.amortizacion === 'number' && isFinite(c.amortizacion))
                       ? Math.round(c.amortizacion) : null
    });
  } catch (err) {
    console.error('Error en /api/calcular:', err);
    if (!res.headersSent) res.status(500).json({ error: 'No se pudo calcular.' });
  } finally {
    fs.unlink(tmpXlsx, () => {});
    fs.rm(outDir, { recursive: true, force: true }, () => {});
  }
});

if (require.main === module) {
  app.listen(PORT, () => console.log(`Microservicio PDF en puerto ${PORT}`));
}

module.exports = { recalcXlsx, congelarCurva, _readNumFromSheet, CURVA_FILAS, CURVA_COLS, normalizeInputs, injectInputs };
