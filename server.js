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

function xlsxToPdf(xlsxBuffer, tmpXlsx, tmpPdf) {
  fs.writeFileSync(tmpXlsx, xlsxBuffer);
  let loResult;
  if (process.platform === 'win32') {
    const soffice = 'C:\\Program Files\\LibreOffice\\program\\soffice.exe';
    loResult = spawnSync('powershell.exe',
      ['-Command', `& '${soffice}' --headless --norestore --convert-to pdf --outdir '${os.tmpdir()}' '${tmpXlsx}'`],
      { timeout: 50000, windowsHide: true });
  } else {
    loResult = spawnSync('soffice',
      ['--headless','--norestore','--convert-to','pdf','--outdir', os.tmpdir(), tmpXlsx],
      { timeout: 50000 });
  }
  if (loResult.status !== 0)
    throw new Error(`LibreOffice status ${loResult.status}: ${(loResult.stderr||'').toString()}`);
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

    const SS = { 'particular':16, 'empresa':37, 'no':7, '2.0 TD':134, '3.0 TD':22, '6.1 TD':140 };
    const consumoVal = (consumo && consumo > 0) ? consumo : kWp * 1000;
    const clienteIdx = cliente === 'empresa' ? SS['empresa'] : SS['particular'];
    const tarifaIdx  = SS[tarifa] ?? SS['2.0 TD'];
    const bateriaVal = bateria ? 1 : 0;
    const excelDate  = Math.round(Date.now() / 86400000 + 25569);

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

    const setNum = (x, ref, val) =>
      x.replace(new RegExp(`(<c r="${ref}"[^>]*>(?:<f[^<]*</f>)?<v>)[^<]*(</v>)`), `$1${val}$2`);
    const setSS  = (x, ref, idx) =>
      x.replace(new RegExp(`(<c r="${ref}"[^>]*t="s"[^>]*>(?:<f[^<]*</f>)?<v>)[^<]*(</v>)`), `$1${idx}$2`);

    let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    xml = setNum(xml, 'N8',  consumoVal);
    xml = setNum(xml, 'I23', numPaneles);
    xml = setNum(xml, 'I31', bateriaVal);
    xml = setNum(xml, 'C9',  excelDate);
    xml = setSS (xml, 'N6',  clienteIdx);
    xml = setSS (xml, 'N7',  tarifaIdx);
    zip.file('xl/worksheets/sheet1.xml', xml);

    let wbXml = await zip.file('xl/workbook.xml').async('string');
    for (const sheetName of ['Datos partida', 'Calculos']) {
      wbXml = wbXml.replace(
        new RegExp(`(<sheet name="${sheetName}"[^>]*?)state="visible"`),
        '$1state="veryHidden"'
      );
    }
    zip.file('xl/workbook.xml', wbXml);

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
