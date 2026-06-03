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
    const { kWp, numPaneles, consumo, cliente, tarifa, fase, bateria,
            address, lat, lng } = req.body;

    const consumoVal   = (consumo && consumo > 0) ? consumo : kWp * 1000;
    const bateriaVal   = bateria ? 1 : 0;
    const excelDate    = Math.round(Date.now() / 86400000 + 25569);
    const faseExcel    = (fase && /tri/i.test(fase)) ? 'trifásico' : 'monofásico';
    const clienteExcel = (cliente && /empresa/i.test(cliente)) ? 'empresa' : 'particular';
    const tarifaExcel  = tarifa || '2.0 TD';

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

    // Escribe un valor en una celda de la hoja Factura, FORZÁNDOLO literal
    // (elimina cualquier fórmula previa para que LibreOffice no la recalcule).
    const cellRe = ref => new RegExp(`<c r="${ref}"((?:[^>]*?))(?:/>|>[\\s\\S]*?</c>)`);
    const styleOf = attrs => ((/ s="\d+"/.exec(attrs) || [''])[0]);
    const xmlEsc  = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const colNum  = c => { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
    const setNum  = (x, ref, val) => x.replace(cellRe(ref),
      (m, attrs) => `<c r="${ref}"${styleOf(attrs)}><v>${val}</v></c>`);
    const setStr  = (x, ref, str) => x.replace(cellRe(ref),
      (m, attrs) => `<c r="${ref}"${styleOf(attrs)} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(str)}</t></is></c>`);
    // Como setNum, pero si la celda NO existe en el XML, la INSERTA en su fila en orden
    // de columna (necesario para P9 y M16, que vienen vacías en la plantilla).
    const setCellNum = (x, ref, val) => {
      if (cellRe(ref).test(x)) return setNum(x, ref, val);
      const [, col, row] = ref.match(/^([A-Z]+)(\d+)$/);
      const rowRe = new RegExp(`(<row r="${row}"[^>]*>)([\\s\\S]*?)(</row>)`);
      return x.replace(rowRe, (m, open, body, close) => {
        const newCell = `<c r="${ref}"><v>${val}</v></c>`;
        let insertAt = body.length;
        for (const cm of body.matchAll(/<c r="([A-Z]+)\d+"[\s\S]*?(?:\/>|<\/c>)/g)) {
          if (colNum(cm[1]) > colNum(col)) { insertAt = cm.index; break; }
        }
        return open + body.slice(0, insertAt) + newCell + body.slice(insertAt) + close;
      });
    };

    // Hoja Factura (sheet1) — celdas de entrada confirmadas con Toño:
    let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    xml = setNum(xml, 'N8',  consumoVal);    // Consumo kWh/año
    xml = setCellNum(xml, 'M16', numPaneles);// Nº paneles (override manual) → título kWp + precio
    xml = setCellNum(xml, 'P9',  kWp);       // Potencia inversor kW (override manual)
    xml = setStr(xml, 'O9',  faseExcel);     // monofásico / trifásico
    xml = setNum(xml, 'I31', bateriaVal);    // Batería (0/1)
    xml = setNum(xml, 'C9',  excelDate);     // Fecha
    xml = setStr(xml, 'N6',  clienteExcel);  // Particular/empresa → IVA
    xml = setStr(xml, 'N7',  tarifaExcel);   // Tarifa
    // NOTA: se respeta el área de impresión y la escala del Excel de Toño tal cual
    // (3 páginas). Las 3 páginas limpias dependen de tener la fuente Calibri/Carlito
    // instalada en el contenedor (ver Dockerfile), no de modificar el pageSetup.
    zip.file('xl/worksheets/sheet1.xml', xml);

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

    // Convertir los rellenos de DEGRADADO de los gráficos a color sólido (su color
    // dominante). LibreOffice headless dibuja los degradados como cuadros negros;
    // con sólido se ven bien y se respeta el color elegido en el Excel.
    const chartFiles = Object.keys(zip.files).filter(f => /^xl\/charts\/chart\d+\.xml$/.test(f));
    for (const cf of chartFiles) {
      let cxml = await zip.file(cf).async('string');
      cxml = cxml.replace(/<a:gradFill[\s\S]*?<\/a:gradFill>/g, (g) => {
        const gs = g.match(/<a:gs\b[^>]*>([\s\S]*?)<\/a:gs>/); // primer punto del degradado
        const color = gs ? gs[1] : '<a:srgbClr val="4472C4"/>';
        return `<a:solidFill>${color}</a:solidFill>`;
      });
      zip.file(cf, cxml);
    }

    const mapPngBuffer = await mapPromise;
    if (mapPngBuffer) {
      zip.file('xl/media/image1.png', mapPngBuffer);
      console.log('[MAPA] image1.png sustituida en xlsx');
    }

    const xlsxBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

    console.log('[PDF] Convirtiendo con LibreOffice...');
    const pdfBytes = xlsxToPdf(xlsxBuf, tmpXlsx, tmpPdf);
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

app.listen(PORT, () => console.log(`Microservicio PDF en puerto ${PORT}`));
