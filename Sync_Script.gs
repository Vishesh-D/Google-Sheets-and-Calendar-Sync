/**
 * Google Sheets → Calendar Sync
 * Copyright (c) 2025 Vishesh Dalal
 * License: MIT
 * Repository: https://github.com/yourname/sheets-calendar-sync
 */


/***** =========================
 *  CONFIG (EDIT THESE)
 *  ======================== *****/

// Source (institute) → filtered copy (your sheet) → Calendar
const CFG = {
  // You can paste either the Spreadsheet ID or the full URL of the institute sheet:
  SOURCE_SPREADSHEET_ID_OR_URL: '1Ws6G9HmVKFQcDzGA7IfxTpAp6P2jKHnXEhJtjuvOOY4',
  SOURCE_SHEET_NAME:            'Term V Schedule', // exact tab name in the institute file
  DEST_SHEET_NAME:              'MyEvents',                 // filtered output tab in your sheet
  ALLOWLIST_SHEET_NAME:         'Allowlist',                // terms in A:A (one per line)
  TZ:                           'Asia/Kolkata',

  // WHERE headers & data live in the SOURCE (matches your screenshots)
  EVENT_HEADER_ROW_INDEX:   5,  // row with "PGP-28 D1", "PGP-28 D2", ..., etc.
  DATETIME_HEADER_ROW_INDEX:5,  // row with "DATE" and "TIME"
  DATA_START_ROW_INDEX:     7,  // first real timetable row (after banner/registration rows)

  // Event columns in the SOURCE layout (use 1-based: A=1, B=2...)
  // These are the columns that contain the per-slot event titles (e.g., D1, D2, …)
  EVENT_COLS: [3,4,5,6,7,8,9,10], // C..J — adjust if your source has more/fewer

  // Allow-list behavior (only keep events matching these terms)
  MATCH_MODE: 'exact',                // 'substring' | 'wholeword' | 'exact'
  KEEP_WHEN_ALLOWLIST_EMPTY: true,    // fail-closed: if allowlist empty → keep nothing
  ALLOWLIST_TERMS_INLINE: [/* 'ER (FIN)', 'GT-B' */],
  DROP_ROWS_WITH_NO_MATCHES: false,

  // Calendar
  CALENDAR_ID: 'primary',             // or paste a specific calendar ID
  TITLE_TEMPLATE: '{COL} - {event}',          // e.g., '{H:D1} — {event}' to prefix with column header

  // Your filtered table layout (DEST_SHEET_NAME)
  DATE_COL: 1,                        // A = DATE
  TIME_COL: 2,                        // B = "HH:MM-HH:MM" (strict parser, minutes preserved)
  LOCATION_COL: null,                 // set to a column number if you add one
  DESCRIPTION_COL: null,              // set to a column number if you add one
  AUTO_DELETE_MISSING: true,
  // If null, auto-detect event columns in DEST as C..last
  DEST_EVENT_COLS: null,

  // Safety / rate limits
  SYNC_BATCH_ROWS: 40,                // sheet rows per run
  SLEEP_MS_BETWEEN_OP: 250,           // throttle between Calendar ops (ms)

  // Automation
  TRIGGER_EVERY_MINUTES: 30,          // 0 = don’t auto-install trigger
};

/***** =========================
 *  ONE-TIME SETUP (run manually)
 *  ======================== *****/

function setupAll() {
  ensureSheet_(CFG.DEST_SHEET_NAME);
  ensureEventMap_();
  maybeCreateMenu_();
  if (CFG.TRIGGER_EVERY_MINUTES > 0) {
    installTimeTrigger_('runPipeline', CFG.TRIGGER_EVERY_MINUTES);
  }
  Logger.log('Setup done.');
}
function bestEventHeaderToken_(vals, col1, rowPref, rowFallbackBandEnd) {
  // 1) Try the preferred row (where D1/D2 live)
  const prefer = (vals[rowPref - 1][col1 - 1] || '').toString().trim();
  if (/\b[A-Z]\d+\b/i.test(prefer)) return prefer;

  // 2) Scan nearby rows (above/below) for something like D1/E2…
  const top = Math.max(1, rowPref - 2);
  const bot = Math.max(rowPref + 1, rowFallbackBandEnd + 1);
  for (let r = top; r <= bot; r++) {
    const v = (vals[r - 1][col1 - 1] || '').toString().trim();
    if (/\b[A-Z]\d+\b/i.test(v)) return v;
  }

  // 3) Fall back to first non-empty within the band
  for (let r = top; r <= bot; r++) {
    const v = (vals[r - 1][col1 - 1] || '').toString().trim();
    if (v) return v;
  }
  return '';
}

/***** =========================
 *  SINGLE ENTRY POINT (for trigger)
 *  ======================== *****/

function runPipeline() {
  pullFilter_();     // 1) Source → (skip red + allowlist) → MyEvents (with headers)
  syncToCalendar_(); // 2) MyEvents → Google Calendar (batched + throttled)
}


function makeKey_(title, start, end, location, description) {
  // Build a unique string that identifies one calendar event
  const parts = [
    title || '',
    start ? start.getTime() : '',
    end ? end.getTime() : '',
    location || '',
    description || ''
  ];
  return parts.join('||');
}




function buildTitle_(eventText, row, headers, colIndex) {
  let title = CFG.TITLE_TEMPLATE || '{event}';
  title = title.replace('{event}', eventText);

  // {H:HeaderName} → looks for that header in row 1
  title = title.replace(/\{H:([^}]+)\}/g, (_, hName) => {
    const idx = headers.indexOf(hName);
    if (idx >= 0) return String(row[idx] || '');
    return '';
  });

  // {C:n} → raw cell value from column n
  title = title.replace(/\{C:(\d+)\}/g, (_, n) => {
    const i = parseInt(n, 10) - 1;
    return (i >= 0 && i < row.length) ? String(row[i] || '') : '';
  });

  // NEW: {COL} → the header of the current column (passed in as colIndex)
  if (colIndex != null) {
    const header = headers[colIndex] || '';
    title = title.replace('{COL}', header);
  }

  return title.trim();
}



/***** =========================
 *  MENU
 *  ======================== *****/

function maybeCreateMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Sync')
      .addItem('Run pipeline now', 'runPipeline')
      .addItem('Install time trigger', 'installDefaultTrigger')
      .addItem('Remove time trigger', 'removePipelineTriggers')
      .addItem('Cleanup missing calendar events', 'cleanupMissingCalendarEvents')
      .addToUi();
  } catch (e) {
    // No UI context (e.g., running from a trigger) — ignore
  }
}

function installDefaultTrigger() {
  installTimeTrigger_('runPipeline', CFG.TRIGGER_EVERY_MINUTES || 30);
}

function removePipelineTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runPipeline') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Removed triggers for runPipeline.');
}

/***** =========================
 *  PART 1: Pull + per-cell filter into MyEvents (with headers)
 *  ======================== *****/

function pullFilter_() {
  const srcSS = openByIdOrUrl_(CFG.SOURCE_SPREADSHEET_ID_OR_URL);
  if (!srcSS) throw new Error('Could not open source spreadsheet. Check ID/URL or access.');
  const srcSh = srcSS.getSheetByName(CFG.SOURCE_SHEET_NAME);
  if (!srcSh) throw new Error('Source sheet not found: ' + CFG.SOURCE_SHEET_NAME);

  const lastRow = srcSh.getLastRow(), lastCol = srcSh.getLastColumn();
  if (lastRow < CFG.DATA_START_ROW_INDEX) { writeDest_([['(empty source)']]); return; }

  const vals = srcSh.getRange(1, 1, lastRow, lastCol).getValues();
  const bgs  = srcSh.getRange(1, 1, lastRow, lastCol).getBackgrounds();

// --- Build clean headers for MyEvents ---
const dateLabel = (vals[CFG.DATETIME_HEADER_ROW_INDEX - 1][0] || 'DATE').toString().trim() || 'DATE';
const timeLabel = (vals[CFG.DATETIME_HEADER_ROW_INDEX - 1][1] || 'TIME').toString().trim() || 'TIME';

// Prefer tokens like D1/E1/etc if present; otherwise fall back to first non-empty above
const eventHeaders = [];
for (const c1 of CFG.EVENT_COLS) {
  const raw = bestEventHeaderToken_(vals, c1, CFG.EVENT_HEADER_ROW_INDEX, CFG.DATETIME_HEADER_ROW_INDEX);
  eventHeaders.push(normalizeEventHeader_(raw));
}
const cleanHeaders = [dateLabel, timeLabel, ...eventHeaders];


  // --- Data after the banner rows ---
  const data  = vals.slice(CFG.DATA_START_ROW_INDEX - 1);
  const dataB = bgs.slice(CFG.DATA_START_ROW_INDEX - 1);

  const allow = readAllowlist_();
  const out = [cleanHeaders];

  for (let r = 0; r < data.length; r++) {
    const srcRow = data[r].slice();
    const rowBg  = dataB[r].slice();

    const outRow = [ srcRow[0], srcRow[1] ]; // DATE, TIME
    let keptAny = false;

    for (let k = 0; k < CFG.EVENT_COLS.length; k++) {
      const c1 = CFG.EVENT_COLS[k];
      const i = c1 - 1;

      const cellText = String(srcRow[i] ?? '').trim();
      const isRed = isReddish_(normalizeHex_(rowBg[i]));
      const matches = (allow.length ? cellMatchesAllow_(cellText, allow) : !CFG.KEEP_WHEN_ALLOWLIST_EMPTY);

      if (isRed || !matches) {
        outRow.push('');
      } else {
        outRow.push(cellText);
        if (cellText) keptAny = true;
      }
    }

    if (CFG.DROP_ROWS_WITH_NO_MATCHES && !keptAny) continue;
    out.push(outRow);
  }

  writeDest_(out.length ? out : [cleanHeaders]);
}

// Scan header rows [top..bot] (inclusive) for the first non-empty cell in column c1
function findFirstHeaderAbove_(vals, topRowIdx1, botRowIdx1, col1) {
  const top = Math.max(1, topRowIdx1);
  const bot = Math.max(top, botRowIdx1);
  const c = Math.max(1, col1) - 1;
  for (let r1 = top; r1 <= bot; r1++) {
    const v = (vals[r1 - 1][c] || '').toString().trim();
    if (v) return v;
  }
  return '';
}

/***** =========================
 *  PART 2: Calendar sync (batched + throttled)
 *  ======================== *****/

function syncToCalendar_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DEST_SHEET_NAME);
  if (!sh) throw new Error('Dest sheet not found: ' + CFG.DEST_SHEET_NAME);

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const rows    = sh.getRange(2,1,lastRow-1,lastCol).getValues();

  const eventCols = Array.isArray(CFG.DEST_EVENT_COLS) && CFG.DEST_EVENT_COLS.length
    ? CFG.DEST_EVENT_COLS
    : (function(){ const a=[]; for (let i=3;i<=lastCol;i++) a.push(i); return a; })();

  const map = readEventMap_();                   // key -> { id, seen:false }
  const cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID);
  if (!cal) throw new Error('Calendar not found: ' + CFG.CALENDAR_ID);

  let made=0, updated=0, skipped=0;

  for (const row of rows) {
    const dt = getStartEndFromRowIST_(row);
    if (!dt) { skipped++; continue; }
    const {start, end} = dt;

    const baseDesc = getColVal_(row, CFG.DESCRIPTION_COL);
    const location = getColVal_(row, CFG.LOCATION_COL);

    for (const c1 of eventCols) {
      const i = c1 - 1;
      const txt = (row[i] ?? '').toString().trim();
      if (!txt) continue;

      // Column header for this event cell (e.g., "D1", "E1", …)
      const cohortHeader = String(headers[i] || '').trim();

      const title = buildTitle_(txt, row, headers, i);

      const finalDesc = [baseDesc, cohortHeader ? `Cohort: ${cohortHeader}` : '']
        .filter(Boolean)
        .join('\n');

      // New key (includes cohort note) + legacy key (without it)
      const key       = makeKey_(title, start, end, location, finalDesc);
      const legacyKey = makeKey_(title, start, end, location, baseDesc || '');

      let existing = map.get(key) || map.get(legacyKey);

      try {
        if (existing && existing.id) {
          const ev = getEventByIdSafe_(existing.id);
          if (ev) {
            if (ev.getTitle() !== title) ev.setTitle(title);
            if (+ev.getStartTime() !== +start || +ev.getEndTime() !== +end) ev.setTime(start, end);
            if ((ev.getLocation()||'') !== (location||'')) ev.setLocation(location||'');
            if ((ev.getDescription()||'') !== (finalDesc||'')) ev.setDescription(finalDesc||'');

            // If we matched legacy key, migrate mapping to the new key
            if (!map.has(key)) {
              map.delete(legacyKey);
              map.set(key, { id: existing.id, seen: true });
            } else {
              existing.seen = true;
            }
            updated++;
          } else {
            const ne = createEvent_(cal, title, start, end, location, finalDesc);
            map.set(key, { id: ne.getId(), seen: true });
            made++;
          }
        } else {
          const ne = createEvent_(cal, title, start, end, location, finalDesc);
          map.set(key, { id: ne.getId(), seen: true });
          made++;
        }
      } catch (e) {
        if (String(e).match(/too many/i)) Utilities.sleep(1500);
        else throw e;
      }

      Utilities.sleep(CFG.SLEEP_MS_BETWEEN_OP);
    }
  }

  // --- Auto-delete events that are no longer present in MyEvents ---
  if (CFG.AUTO_DELETE_MISSING) {
    for (const [k, rec] of Array.from(map.entries())) {
      if (rec && rec.id && rec.seen !== true) {
        const ev = getEventByIdSafe_(rec.id);
        try { if (ev) ev.deleteEvent(); } catch (_) {/* swallow: already gone or perms */}
        map.delete(k);
        Utilities.sleep(100); // tiny backoff
      } else if (rec) {
        // reset seen flag for next run
        rec.seen = false;
      }
    }
  }

  writeEventMap_(map);
  Logger.log(`All rows: created=${made}, updated=${updated}, skippedRows=${skipped}, deletedMissing=${CFG.AUTO_DELETE_MISSING ? 'yes' : 'no'}`);
}



function cleanupMissingCalendarEvents() {
  const map = readEventMap_();
  const cal = CalendarApp.getCalendarById(CFG.CALENDAR_ID);
  for (const [key, rec] of map.entries()) {
    if (rec && rec.id && rec.seen === false) {
      const ev = getEventByIdSafe_(rec.id);
      if (ev) ev.deleteEvent();
      map.delete(key);
    }
  }
  writeEventMap_(map);
}

/***** =========================
 *  Allow-list + color helpers
 *  ======================== *****/

function readAllowlist_() {
  const sheetTerms = (() => {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.ALLOWLIST_SHEET_NAME);
    if (!sh) return [];
    const n = sh.getLastRow(); if (n < 1) return [];
    return sh.getRange(1,1,n,1).getValues()
      .map(r => (r[0] || '').toString().trim())
      .filter(s => s && !s.startsWith('#'));
  })();

  const inline = (CFG.ALLOWLIST_TERMS_INLINE || [])
    .map(s => String(s || '').trim())
    .filter(Boolean);

  const list = [...sheetTerms, ...inline]
    .map(s => s.toLowerCase().replace(/\s+/g,' ').trim());

  return Array.from(new Set(list));
}

function cellMatchesAllow_(text, allow) {
  const t = String(text || '').toLowerCase().replace(/\s+/g,' ').trim();
  if (!t) return false;

  if (CFG.MATCH_MODE === 'exact') {
    const set = new Set(allow);
    return set.has(t);
  }
  if (CFG.MATCH_MODE === 'wholeword') {
    const tokens = new Set(t.split(/\W+/).filter(Boolean));
    return allow.some(term => tokens.has(term));
  }
  // substring
  return allow.some(term => t.indexOf(term) !== -1);
}

function normalizeHex_(hex) {
  let h = String(hex || '').trim();
  if (!h) return '#ffffff';
  if (!h.startsWith('#')) h = '#' + h;
  if (h.length === 4) h = '#' + h.slice(1).split('').map(c => c + c).join('');
  return h.length === 7 ? h.toLowerCase() : '#ffffff';
}
function isReddish_(hex) {
  const {r,g,b} = hexToRgb_(hex);
  return r > 140 && r > g + 20 && r > b + 20;
}
function hexToRgb_(hex) {
  const h = hex.replace('#',''); const num = parseInt(h || 'ffffff', 16);
  return { r: (num >> 16) & 255, g: (num >> 8) & 255, b: num & 255 };
}

/***** =========================
 *  Date/time (IST) helpers
 *  ======================== *****/

function getStartEndFromRowIST_(row) {
  const dateVal = getColVal_(row, CFG.DATE_COL);
  const timeVal = getColVal_(row, CFG.TIME_COL);
  return parseTimeRangeIST_(dateVal, timeVal);
}

// STRICT: preserves minutes; accepts "HH:MM-HH:MM", "HH.MM-HH.MM", "HHMM-HHMM"
function parseTimeRangeIST_(dateVal, timeRangeStr) {
  const tz = CFG.TZ || 'Asia/Kolkata';
  const base = toDateSafeIST_(dateVal, tz); if (!base) return null;

  let s = String(timeRangeStr || '').trim()
    .replace(/[–—]/g, '-')    // en/em dash -> hyphen
    .replace(/\s+/g, '');     // remove spaces

  const m = s.match(/^(\d{1,2})[:\.]?(\d{2})-(\d{1,2})[:\.]?(\d{2})$/);
  if (!m) return null;

  const sh = clampHour_(+m[1]), sm = clampMin_(+m[2]);
  const eh = clampHour_(+m[3]), em = clampMin_(+m[4]);

  const start = new Date(base); start.setHours(sh, sm, 0, 0);
  const end   = new Date(base); end.setHours(eh, em, 0, 0);

  if (end <= start) return null;
  return { start, end };
}

function toDateSafeIST_(val, tz) {
  const d = (val instanceof Date && !isNaN(val)) ? val : new Date(val);
  if (!(d instanceof Date) || isNaN(d)) return null;
  const y = +Utilities.formatDate(d, tz, 'yyyy');
  const m = +Utilities.formatDate(d, tz, 'MM');
  const dd= +Utilities.formatDate(d, tz, 'dd');
  return new Date(y, m - 1, dd, 0, 0, 0, 0);
}
function clampHour_(h){ return Math.max(0, Math.min(23, h)); }
function clampMin_(m){ return Math.max(0, Math.min(59, m)); }

/***** =========================
 *  Header helper
 *  ======================== *****/

function normalizeEventHeader_(raw) {
  // "PGP-28 D1" → "D1", "PGPFIN05 E1" → "E1", otherwise keep trimmed text
  const s = (raw || '').toString().trim();
  if (!s) return '';
  const m = s.match(/\b([A-Z]\d+)\b/i);
  return m ? m[1].toUpperCase() : s;
}

/***** =========================
 *  Calendar + EventMap
 *  ======================== *****/

function ensureEventMap_() {
  const sh = ensureSheet_('EventMap');
  const header = sh.getRange(1,1,1,3).getValues()[0];
  if (header[0] !== 'Key' || header[1] !== 'EventID' || header[2] !== 'Seen') {
    sh.clear(); sh.getRange(1,1,1,3).setValues([['Key','EventID','Seen']]);
  }
}

function readEventMap_() {
  ensureEventMap_();
  const sh = SpreadsheetApp.getActive().getSheetByName('EventMap');
  const last = sh.getLastRow();
  const map = new Map();
  if (last < 2) return map;
  const rows = sh.getRange(2,1,last-1,3).getValues();
  for (const [key,id] of rows) if (key && id) map.set(String(key), { id: String(id), seen: false });
  return map;
}

function writeEventMap_(map) {
  ensureEventMap_();
  const sh = SpreadsheetApp.getActive().getSheetByName('EventMap');
  sh.clear(); sh.getRange(1,1,1,3).setValues([['Key','EventID','Seen']]);
  const out = [];
  for (const [key, rec] of map.entries()) out.push([key, rec.id || '', rec.seen ? 'true' : 'false']);
  if (out.length) sh.getRange(2,1,out.length,3).setValues(out);
}

function createEvent_(cal, title, start, end, location, description) {
  const opt = {};
  if (location) opt.location = String(location);
  if (description) opt.description = String(description);
  return cal.createEvent(String(title), start, end, opt);
}
function getEventByIdSafe_(id) {
  let ev = null;
  try { ev = CalendarApp.getEventById(id); } catch(_) {}
  if (!ev && id && !String(id).includes('@google.com')) {
    try { ev = CalendarApp.getEventById(String(id) + '@google.com'); } catch(_) {}
  }
  return ev;
}

/***** =========================
 *  IO / Utils / Triggers
 *  ======================== *****/

function writeDest_(rows) {
  const sh = ensureSheet_(CFG.DEST_SHEET_NAME);
  sh.clear();
  if (rows && rows.length) sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
}

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function openByIdOrUrl_(idOrUrl) {
  const id = extractId_(idOrUrl);
  try { return SpreadsheetApp.openById(id); } catch(_) { return null; }
}
function extractId_(s) {
  const str = String(s || '').trim();
  if (!str) return '';
  if (str.startsWith('http')) {
    const m = str.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return m ? m[1] : str;
  }
  return str;
}

function getColVal_(row, col1) {
  if (!col1) return '';
  const i = col1 - 1;
  return (i >= 0 && i < row.length) ? row[i] : '';
}

function range_(a,b){ const r=[]; for(let i=a;i<=b;i++) r.push(i); return r; }

function getState_(k, def){ return PropertiesService.getScriptProperties().getProperty(k) ?? def; }
function setState_(k, v){ PropertiesService.getScriptProperties().setProperty(k, String(v)); }
function clearState_(){ PropertiesService.getScriptProperties().deleteAllProperties(); }

function installTimeTrigger_(fnName, everyMinutes) {
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === fnName && t.getEventType() === ScriptApp.EventType.CLOCK);
  if (!exists) ScriptApp.newTrigger(fnName).timeBased().everyMinutes(everyMinutes).create();
}
