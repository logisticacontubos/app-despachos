/**
 * =========================================================================
 * api/action.js — Vercel serverless replacement for Code_final.gs
 * =========================================================================
 * Talks directly to the Google Sheets API (googleapis) instead of going
 * through an Apps Script Web App. Same Google Sheet as source of truth,
 * same sheet names (Cola, Historial, CONFIG, Usuarios, RequestLog), same
 * column layout and same request contract as Code_final.gs, so the
 * frontend only needs to change SCRIPT_URL to '/api/action'.
 *
 * Env vars required (Vercel → Settings → Environment Variables):
 *   GOOGLE_SERVICE_ACCOUNT_EMAIL
 *   GOOGLE_PRIVATE_KEY   (paste with real newlines; this file also
 *                         supports the literal "\n"-escaped form)
 *   GOOGLE_SHEET_ID
 *
 * IMPORTANT: this file is intentionally self-contained (no imports from
 * other project files) — splitting it into api/action.js + lib/sheets.js
 * broke Vercel's build for the PT app ("Cannot find module '../lib/sheets'"),
 * so everything lives in this one file on purpose.
 * =========================================================================
 */

const { google } = require('googleapis');

const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const COLA_BASE_LEN = 26;   // indices 0-25 (status at 25) — trace block appended at 26-33
const HIST_BASE_LEN = 27;   // indices 0-26 (comentario at 25, status at 26) — trace appended at 27-34

// ============================================================
// AUTH (cached across warm invocations)
// ============================================================
let sheetsClientPromise = null;
function getSheetsClient() {
  if (!sheetsClientPromise) {
    const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
    const auth = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      null,
      privateKey,
      ['https://www.googleapis.com/auth/spreadsheets']
    );
    sheetsClientPromise = auth.authorize().then(() => google.sheets({ version: 'v4', auth }));
  }
  return sheetsClientPromise;
}

// ============================================================
// FORMAT HELPERS (equivalent to Code_final.gs formatVal/formatDate/formatTime,
// but reading UNFORMATTED_VALUE from the Sheets API: legacy rows where a
// cell got auto-converted to a real Date arrive as a numeric serial instead
// of a JS Date object, so we convert serials here instead).
// ============================================================
function pad2(n) { return String(n).padStart(2, '0'); }
function serialToDate(serial) {
  // Sheets/Excel epoch is Dec 30, 1899 — 25569 days before the Unix epoch.
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  return new Date(ms);
}
function formatDate(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') {
    const d = serialToDate(val);
    return pad2(d.getUTCDate()) + '/' + pad2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  }
  const s = String(val).trim();
  if (s.includes('GMT') || s.length > 15) {
    const d = new Date(s);
    if (!isNaN(d)) return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear();
  }
  return s;
}
function formatTime(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') {
    const fraction = val - Math.floor(val);
    const totalMinutes = Math.round(fraction * 24 * 60);
    const h = Math.floor(totalMinutes / 60) % 24;
    const m = totalMinutes % 60;
    return pad2(h) + ':' + pad2(m);
  }
  const s = String(val).trim();
  if (s.includes('1899') || s.includes('GMT') || s.includes('Dec 30')) {
    const match = s.match(/(\d{2}):(\d{2})/);
    if (match) return match[1] + ':' + match[2];
  }
  return s;
}
const formatVal = formatTime;

// ============================================================
// SHEET METADATA HELPERS
// ============================================================
let metaCache = null;
async function getMeta(sheets) {
  if (!metaCache) {
    const resp = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties' });
    metaCache = resp.data.sheets;
  }
  return metaCache;
}
async function refreshMeta(sheets) {
  const resp = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties' });
  metaCache = resp.data.sheets;
  return metaCache;
}
async function getFirstSheetTitle(sheets) {
  const meta = await getMeta(sheets);
  const list = meta.slice().sort((a, b) => a.properties.index - b.properties.index);
  return list[0].properties.title;
}
async function sheetExists(sheets, title) {
  const meta = await getMeta(sheets);
  return meta.some(s => s.properties.title === title);
}
async function getSheetId(sheets, title) {
  const meta = await refreshMeta(sheets);
  const found = meta.find(s => s.properties.title === title);
  return found ? found.properties.sheetId : null;
}
async function ensureSheetExists(sheets, title, headerRow) {
  const exists = await sheetExists(sheets, title);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title } } }] }
    });
    await refreshMeta(sheets);
    if (headerRow) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `'${title}'!A1`,
        valueInputOption: 'RAW',
        requestBody: { values: [headerRow] }
      });
    }
  }
}

// ============================================================
// REQUEST LOG (idempotency / dedup)
// ============================================================
async function findRequestLog(sheets, requestId) {
  await ensureSheetExists(sheets, 'RequestLog', ['RequestId', 'Consecutivo', 'Timestamp', 'Action']);
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'RequestLog'!A:D`, valueRenderOption: 'UNFORMATTED_VALUE'
  });
  const rows = resp.data.values || [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === String(requestId).trim()) {
      return { consecutivo: String(rows[i][1] || ''), action: String(rows[i][3] || '') };
    }
  }
  return null;
}
async function logRequest(sheets, requestId, consecutivo, action) {
  await ensureSheetExists(sheets, 'RequestLog', ['RequestId', 'Consecutivo', 'Timestamp', 'Action']);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'RequestLog'!A:D`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [[requestId, consecutivo || '', new Date().toISOString(), action || '']] }
  });
}

// ============================================================
// CONSECUTIVO
// ============================================================
async function getConsecutivoValue(sheets) {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'CONFIG'!B1`, valueRenderOption: 'UNFORMATTED_VALUE'
  });
  const v = resp.data.values && resp.data.values[0] && resp.data.values[0][0];
  return parseInt(v, 10) || 1;
}
async function incrementConsecutivo(sheets) {
  const val = await getConsecutivoValue(sheets);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: `'CONFIG'!B1`, valueInputOption: 'RAW',
    requestBody: { values: [[val + 1]] }
  });
  return val;
}

// ============================================================
// USUARIOS
// ============================================================
async function getUsers(sheets) {
  const exists = await sheetExists(sheets, 'Usuarios');
  if (!exists) return { ok: true, users: [] };
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'Usuarios'!A:D`, valueRenderOption: 'UNFORMATTED_VALUE'
  });
  const all = resp.data.values || [];
  const rows = all.filter((r, i) => {
    if (i === 0 && !/^\d+$/.test(String(r[0] || '').trim())) return false; // header
    return String(r[0] || '').trim() !== '';
  });
  const users = rows.map(r => ({
    pin: String(r[0]).trim(),
    nombre: String(r[1] || '').trim(),
    rol: String(r[2] || '').trim(),
    activo: String(r[3] || '').trim().toUpperCase()
  })).filter(u => u.activo === 'SI');
  return { ok: true, users };
}

// ============================================================
// TRACE / ROW MAPPING (mirrors Code_final.gs mapRow/extractTrace)
// ============================================================
function extractTrace(r, base) {
  return {
    registradoPor: String(r[base] || ''),
    registradoEn: String(r[base + 1] || ''),
    iniciadoPor: String(r[base + 2] || ''),
    iniciadoEn: String(r[base + 3] || ''),
    finalizadoPor: String(r[base + 4] || ''),
    finalizadoEn: String(r[base + 5] || ''),
    requestId: String(r[base + 6] || ''),
    refsExtraJSON: String(r[base + 7] || '')
  };
}
function mapRow(r, sheetBase) {
  const base = sheetBase;
  let refsExtra = [];
  try {
    const raw = r[base + 7] || '';
    if (raw) refsExtra = JSON.parse(raw);
  } catch (e) { refsExtra = []; }
  return {
    consecutivo: String(r[0] || '').trim(),
    tipo: String(r[1] || ''), oc: String(r[2] || ''), costo: String(r[3] || ''), empresa: String(r[4] || ''),
    material: String(r[5] || ''), cantidad: String(r[6] || ''), embalaje: String(r[7] || ''),
    placa: String(r[8] || ''), transportador: String(r[9] || ''), vehiculo: String(r[10] || ''), operario: String(r[11] || ''),
    fechaIngreso: formatDate(r[12]), horaIngreso: formatVal(r[13]),
    horaMuelle: formatVal(r[14]), horaSalida: formatVal(r[15]),
    tiempoEspera: formatTime(r[16]), tiempoDescarga: formatTime(r[17]), tiempoTotal: formatTime(r[18]),
    material2: String(r[19] || ''), cantidad2: String(r[20] || ''), embalaje2: String(r[21] || ''),
    material3: String(r[22] || ''), cantidad3: String(r[23] || ''), embalaje3: String(r[24] || ''),
    status: sheetBase === HIST_BASE_LEN ? 'done' : String(r[25] || 'waiting'),
    comentario: sheetBase === HIST_BASE_LEN ? String(r[25] || '') : '',
    registradoPor: String(r[base] || ''),
    registradoEn: String(r[base + 1] || ''),
    iniciadoPor: String(r[base + 2] || ''),
    iniciadoEn: String(r[base + 3] || ''),
    finalizadoPor: String(r[base + 4] || ''),
    finalizadoEn: String(r[base + 5] || ''),
    requestId: String(r[base + 6] || ''),
    refsExtra: refsExtra
  };
}
function skipHeaderRow(all) {
  if (!all.length) return all;
  const firstCell = String(all[0][0] || '').trim();
  if (firstCell !== '' && !/^\d+$/.test(firstCell)) return all.slice(1);
  return all;
}

// ============================================================
// GET — QUEUE / HISTORY
// ============================================================
async function getQueue(sheets) {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'Cola'`, valueRenderOption: 'UNFORMATTED_VALUE'
  });
  const all = resp.data.values || [];
  if (!all.length) return { ok: true, queue: [] };
  const rows = skipHeaderRow(all);
  const queue = rows.map(r => mapRow(r, COLA_BASE_LEN)).filter(r => r.consecutivo !== '' && r.status !== 'done');
  return { ok: true, queue };
}
async function getHistory(sheets) {
  const exists = await sheetExists(sheets, 'Historial');
  if (!exists) return { ok: true, history: [] };
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'Historial'`, valueRenderOption: 'UNFORMATTED_VALUE'
  });
  const all = resp.data.values || [];
  if (!all.length) return { ok: true, history: [] };
  const rows = skipHeaderRow(all);
  const history = rows.map(r => mapRow(r, HIST_BASE_LEN)).filter(r => r.consecutivo !== '');
  return { ok: true, history };
}

// ============================================================
// WRITES — append / update / removeQueue (mirrors Code_final.gs exactly)
// ============================================================
async function appendRecord(sheets, data) {
  const consec = await incrementConsecutivo(sheets);
  const consecStr = String(consec).padStart(6, '0');
  const row = (data.row || []).slice();
  row[0] = consecStr;

  const mainTitle = await getFirstSheetTitle(sheets);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${mainTitle}'!A:AH`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });

  const trace = extractTrace(row, 26); // MAIN_BASE_LEN
  const colaRow = [
    row[0], row[1], row[2], row[3], row[4],
    row[5], row[6], row[7], row[8], row[9],
    row[10], row[11], row[12], row[13],
    '', '', '', '', '',
    row[19] || '', row[20] || '', row[21] || '',
    row[22] || '', row[23] || '', row[24] || '',
    'waiting',
    trace.registradoPor, trace.registradoEn,
    trace.iniciadoPor, trace.iniciadoEn,
    trace.finalizadoPor, trace.finalizadoEn,
    data.requestId || '', trace.refsExtraJSON
  ];
  await ensureSheetExists(sheets, 'Cola');
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'Cola'!A:AH`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [colaRow] }
  });

  return { consecutivo: consecStr };
}

async function findRowIndexByConsecutivo(sheets, sheetTitle, consecutivo) {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${sheetTitle}'!A:A`, valueRenderOption: 'UNFORMATTED_VALUE'
  });
  const vals = resp.data.values || [];
  for (let i = 0; i < vals.length; i++) {
    if (String((vals[i] && vals[i][0]) || '').trim() === String(consecutivo).trim()) return i; // 0-indexed
  }
  return -1;
}

async function updateMainSheet(sheets, consecutivo, row) {
  if (!row) return;
  const mainTitle = await getFirstSheetTitle(sheets);
  const idx = await findRowIndexByConsecutivo(sheets, mainTitle, consecutivo);
  if (idx === -1) return;
  const rowNum = idx + 1;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: `'${mainTitle}'!A${rowNum}`, valueInputOption: 'RAW',
    requestBody: { values: [row] }
  });
}

async function updateQueueStatus(sheets, consecutivo, row, status) {
  if (!row) return;
  const idx = await findRowIndexByConsecutivo(sheets, 'Cola', consecutivo);
  if (idx === -1) return;
  const rowNum = idx + 1;
  const data = [
    { range: `'Cola'!L${rowNum}`, values: [[row[11] || '']] },   // operario
    { range: `'Cola'!O${rowNum}`, values: [[row[14] || '']] },   // horaMuelle
    { range: `'Cola'!P${rowNum}`, values: [[row[15] || '']] },   // horaSalida
    { range: `'Cola'!Q${rowNum}`, values: [[row[16] || '']] },   // tiempoEspera
    { range: `'Cola'!R${rowNum}`, values: [[row[17] || '']] },   // tiempoDescarga
    { range: `'Cola'!S${rowNum}`, values: [[row[18] || '']] },   // tiempoTotal
    { range: `'Cola'!Z${rowNum}`, values: [[status || 'waiting']] }, // status
    { range: `'Cola'!AA${rowNum}`, values: [[row[26] || '']] },  // registradoPor
    { range: `'Cola'!AB${rowNum}`, values: [[row[27] || '']] },  // registradoEn
    { range: `'Cola'!AC${rowNum}`, values: [[row[28] || '']] },  // iniciadoPor
    { range: `'Cola'!AD${rowNum}`, values: [[row[29] || '']] },  // iniciadoEn
    { range: `'Cola'!AE${rowNum}`, values: [[row[30] || '']] },  // finalizadoPor
    { range: `'Cola'!AF${rowNum}`, values: [[row[31] || '']] },  // finalizadoEn
    { range: `'Cola'!AG${rowNum}`, values: [[row[32] || '']] },  // requestId
    { range: `'Cola'!AH${rowNum}`, values: [[row[33] || '']] }   // refsExtraJSON
  ];
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { valueInputOption: 'RAW', data }
  });
}

async function removeFromQueue(sheets, consecutivo) {
  const idx = await findRowIndexByConsecutivo(sheets, 'Cola', consecutivo);
  if (idx === -1) return;
  const colaSheetId = await getSheetId(sheets, 'Cola');
  if (colaSheetId === null) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId: colaSheetId, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 }
        }
      }]
    }
  });
}

async function addToHistory(sheets, record) {
  if (!record) return;
  await ensureSheetExists(sheets, 'Historial');
  const row = [
    record.consecutivo, record.tipo, record.oc, record.costo, record.empresa,
    record.material, record.cantidad, record.embalaje, record.placa,
    record.transportador, record.vehiculo, record.operario,
    record.fechaIngreso, record.horaIngreso,
    record.horaMuelle || '', record.horaSalida || '',
    record.tiempoEspera || '', record.tiempoDescarga || '', record.tiempoTotal || '',
    record.material2 || '', record.cantidad2 || '', record.embalaje2 || '',
    record.material3 || '', record.cantidad3 || '', record.embalaje3 || '',
    record.comentario || '', 'done',
    record.registradoPor || '', record.registradoEn || '',
    record.iniciadoPor || '', record.iniciadoEn || '',
    record.finalizadoPor || '', record.finalizadoEn || '',
    record.requestId || '',
    record.refsExtraJSON || (record.refsExtra ? JSON.stringify(record.refsExtra) : '')
  ];
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'Historial'!A:AI`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });
}

// ============================================================
// HTTP HANDLER
// ============================================================
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }

  if (!SHEET_ID || !process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
    res.status(500).json({ ok: false, error: 'Missing GOOGLE_SHEET_ID / GOOGLE_SERVICE_ACCOUNT_EMAIL / GOOGLE_PRIVATE_KEY env vars' });
    return;
  }

  try {
    const sheets = await getSheetsClient();

    if (req.method === 'GET') {
      const action = (req.query && req.query.action) || 'getQueue';
      let result;
      if (action === 'getHistory') result = await getHistory(sheets);
      else if (action === 'getConsecutivo') result = { ok: true, consecutivo: await getConsecutivoValue(sheets) };
      else if (action === 'getUsers') result = await getUsers(sheets);
      else result = await getQueue(sheets);
      res.status(200).json(result);
      return;
    }

    if (req.method === 'POST') {
      let data = req.body;
      if (typeof data === 'string') {
        try { data = JSON.parse(data); } catch (e) { data = {}; }
      }
      data = data || {};
      const requestId = data.requestId || '';

      if (requestId) {
        const prev = await findRequestLog(sheets, requestId);
        if (prev) { res.status(200).json({ ok: true, consecutivo: prev.consecutivo, replay: true }); return; }
      }

      if (data.action === 'append') {
        const result = await appendRecord(sheets, data);
        if (requestId) await logRequest(sheets, requestId, result.consecutivo, 'append');
        res.status(200).json({ ok: true, consecutivo: result.consecutivo });
        return;
      }
      if (data.action === 'update') {
        await updateMainSheet(sheets, data.consecutivo, data.row);
        await updateQueueStatus(sheets, data.consecutivo, data.row, data.status);
        if (requestId) await logRequest(sheets, requestId, data.consecutivo, 'update');
        res.status(200).json({ ok: true });
        return;
      }
      if (data.action === 'removeQueue') {
        await removeFromQueue(sheets, data.consecutivo);
        await addToHistory(sheets, data.record);
        if (requestId) await logRequest(sheets, requestId, data.consecutivo, 'removeQueue');
        res.status(200).json({ ok: true });
        return;
      }
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ ok: false, error: 'Method not allowed' });
  } catch (err) {
    console.error('api/action error:', err);
    res.status(500).json({ ok: false, error: String((err && err.message) || err) });
  }
};
