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
function formatDateTime(val) {
  // Full date+time formatter for trace columns (registradoEn/iniciadoEn/
  // finalizadoEn). Legacy rows from before the text-format guard was in
  // place can arrive as a numeric datetime serial (date+time combined);
  // newer rows are already plain "dd/mm/yyyy hh:mm:ss" text and pass
  // through unchanged.
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') {
    const d = serialToDate(val);
    return pad2(d.getUTCDate()) + '/' + pad2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear() + ' ' +
      pad2(d.getUTCHours()) + ':' + pad2(d.getUTCMinutes()) + ':' + pad2(d.getUTCSeconds());
  }
  const s = String(val).trim();
  if (s.includes('GMT') || s.length > 25) {
    const d = new Date(s);
    if (!isNaN(d)) {
      return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' +
        pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
    }
  }
  return s;
}

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
// Resolves the actual tab title to use for the "queue" sheet, matching
// case-insensitively against whatever already exists in the spreadsheet
// (e.g. "Cola" vs "COLA") instead of assuming the literal string 'Cola'.
// This guards against a silent name mismatch causing every write to land
// in (or try to create) a different sheet than the one visible to users.
// Falls back to creating 'Cola' only if no such tab exists at all.
let colaTitleCache = null;
async function resolveColaTitle(sheets) {
  if (colaTitleCache) return colaTitleCache;
  const meta = await getMeta(sheets);
  const found = meta.find(s => String(s.properties.title || '').trim().toLowerCase() === 'cola');
  if (found) {
    colaTitleCache = found.properties.title;
    return colaTitleCache;
  }
  await ensureSheetExists(sheets, 'Cola');
  colaTitleCache = 'Cola';
  return colaTitleCache;
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
    registradoEn: formatDateTime(r[base + 1]),
    iniciadoPor: String(r[base + 2] || ''),
    iniciadoEn: formatDateTime(r[base + 3]),
    finalizadoPor: String(r[base + 4] || ''),
    finalizadoEn: formatDateTime(r[base + 5]),
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
// Writes a row at an EXPLICITLY computed next-empty-row using values.update,
// instead of values.append()'s "guess where the table is" heuristic. This
// exists because values.append() on the COLA sheet was found to silently
// misdetect the table's starting column (writing a 34-column row at
// 'COLA'!Z30:BG30 instead of 'COLA'!A30:AH30), which is why appended queue
// rows were invisible in the sheet the user actually looks at. Reading the
// real next row first and writing with an explicit A:AH-style range sidesteps
// that heuristic entirely.
async function appendRowExact(sheets, title, row, lastColLetter) {
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${title}'!A:A`, valueRenderOption: 'UNFORMATTED_VALUE'
  });
  const existing = resp.data.values || [];
  const nextRow = existing.length + 1;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range: `'${title}'!A${nextRow}:${lastColLetter}${nextRow}`, valueInputOption: 'RAW',
    requestBody: { values: [row] }
  });
  return nextRow;
}

async function getQueue(sheets) {
  const colaTitle = await resolveColaTitle(sheets);
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${colaTitle}'`, valueRenderOption: 'UNFORMATTED_VALUE'
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
async function appendRecord(sheets, data, requestId) {
  const consec = await incrementConsecutivo(sheets);
  const consecStr = String(consec).padStart(6, '0');
  const row = (data.row || []).slice();
  row[0] = consecStr;

  const mainTitle = await getFirstSheetTitle(sheets);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${mainTitle}'!A:AH`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });

  // IMPORTANT: seal the requestId as soon as the main-sheet row (with its
  // freshly-assigned consecutivo) is safely written. This is the piece that
  // was missing before: if it were only logged after the Cola write below,
  // then any failure in that Cola step meant a client retry (same
  // requestId, up to 4 attempts) would never find a log entry and would
  // re-run this whole function — re-incrementing the consecutivo and
  // re-appending to the main sheet each time. That produced the
  // 000333-000336 duplicate-row incident. Sealing here means a retry after
  // this point always short-circuits via findRequestLog() and replays this
  // same consecutivo instead of creating a new row.
  if (requestId) await logRequest(sheets, requestId, consecStr, 'append');

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
  try {
    const colaTitle = await resolveColaTitle(sheets);
    await appendRowExact(sheets, colaTitle, colaRow, 'AH');
  } catch (e) {
    // Don't let a Cola-sheet failure turn into a duplicate DATOS row or a
    // failed save from the user's point of view — the record is already
    // safely stored and the requestId is already sealed above. Log it so
    // it's visible in Vercel's function logs for follow-up.
    console.error('appendRecord: failed to append to Cola/queue sheet for consecutivo ' + consecStr + ':', e);
  }

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
  const colaTitle = await resolveColaTitle(sheets);
  const idx = await findRowIndexByConsecutivo(sheets, colaTitle, consecutivo);
  if (idx === -1) return;
  const rowNum = idx + 1;
  const data = [
    { range: `'${colaTitle}'!L${rowNum}`, values: [[row[11] || '']] },   // operario
    { range: `'${colaTitle}'!O${rowNum}`, values: [[row[14] || '']] },   // horaMuelle
    { range: `'${colaTitle}'!P${rowNum}`, values: [[row[15] || '']] },   // horaSalida
    { range: `'${colaTitle}'!Q${rowNum}`, values: [[row[16] || '']] },   // tiempoEspera
    { range: `'${colaTitle}'!R${rowNum}`, values: [[row[17] || '']] },   // tiempoDescarga
    { range: `'${colaTitle}'!S${rowNum}`, values: [[row[18] || '']] },   // tiempoTotal
    { range: `'${colaTitle}'!Z${rowNum}`, values: [[status || 'waiting']] }, // status
    { range: `'${colaTitle}'!AA${rowNum}`, values: [[row[26] || '']] },  // registradoPor
    { range: `'${colaTitle}'!AB${rowNum}`, values: [[row[27] || '']] },  // registradoEn
    { range: `'${colaTitle}'!AC${rowNum}`, values: [[row[28] || '']] },  // iniciadoPor
    { range: `'${colaTitle}'!AD${rowNum}`, values: [[row[29] || '']] },  // iniciadoEn
    { range: `'${colaTitle}'!AE${rowNum}`, values: [[row[30] || '']] },  // finalizadoPor
    { range: `'${colaTitle}'!AF${rowNum}`, values: [[row[31] || '']] },  // finalizadoEn
    { range: `'${colaTitle}'!AG${rowNum}`, values: [[row[32] || '']] },  // requestId
    { range: `'${colaTitle}'!AH${rowNum}`, values: [[row[33] || '']] }   // refsExtraJSON
  ];
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { valueInputOption: 'RAW', data }
  });
}

async function removeFromQueue(sheets, consecutivo) {
  const colaTitle = await resolveColaTitle(sheets);
  const idx = await findRowIndexByConsecutivo(sheets, colaTitle, consecutivo);
  if (idx === -1) return;
  const colaSheetId = await getSheetId(sheets, colaTitle);
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
      else if (action === 'debugCola') {
        // TEMPORARY diagnostic route: tries to resolve + write a throwaway
        // test row into the queue sheet and reports back exactly what
        // happened, since Vercel's log viewer hasn't been showing the
        // console.error() from appendRecord's catch block. Safe to remove
        // once the Cola-write issue is confirmed fixed.
        const debug = { ok: true };
        try {
          const meta = await getMeta(sheets);
          debug.allSheetTitles = meta.map(s => s.properties.title);
          const colaTitle = await resolveColaTitle(sheets);
          debug.resolvedColaTitle = colaTitle;
          const testRow = ['__DEBUG__', 'DEBUG', '', '', '', '', '', '', '', '', '', '', '', '',
            '', '', '', '', '', '', '', '', '', '', '', 'waiting', '', '', '', '', '', '', 'debug-' + Date.now(), ''];
          const nextRow = await appendRowExact(sheets, colaTitle, testRow, 'AH');
          debug.wroteOk = true;
          debug.wroteAtRow = nextRow;
        } catch (e) {
          debug.wroteOk = false;
          debug.error = String((e && e.message) || e);
          debug.errorCode = e && e.code;
          debug.errorDetails = e && e.errors ? e.errors : (e && e.response && e.response.data ? e.response.data : undefined);
        }
        result = debug;
      }
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
        // requestId is sealed inside appendRecord() itself, right after the
        // main-sheet row is written (see comment there) — not here — so a
        // retry can never re-run the consecutivo/DATOS-append step.
        const result = await appendRecord(sheets, data, requestId);
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
        // Seal here, right after the row leaves the queue — same reasoning
        // as appendRecord(): if the Historial write below fails, a retry
        // (same requestId) must not re-run and append a second Historial
        // row for the same dispatch.
        if (requestId) await logRequest(sheets, requestId, data.consecutivo, 'removeQueue');
        try {
          await addToHistory(sheets, data.record);
        } catch (e) {
          console.error('removeQueue: failed to append to Historial for consecutivo ' + data.consecutivo + ':', e);
        }
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
