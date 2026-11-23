/**
 * Macros spreadsheet API — Google Apps Script web app (container-bound to the sheet).
 *
 * Install: Extensions -> Apps Script, paste this file, set Script Property TOKEN,
 * Deploy -> Web app (Execute as: Me, Who has access: Anyone). See sheet_api/README.md.
 * After editing: Deploy -> Manage deployments -> Edit -> Version: New version (keeps the URL).
 *
 * All requests: POST JSON {"token": "...", "action": "...", ...}
 *
 * Columns are addressed by their row-1 header name (or a column letter like "K" if no header
 * matches). Formulas are strings starting with "=" and may use placeholders:
 *   {row}          -> the row being written
 *   {col:<header>} -> the column letter of <header> in the same sheet
 * e.g. "=4*{col:פחמימה}{row} + 4*{col:חלבון}{row} + 9*{col:שומן}{row}"
 *
 * Actions:
 *   tabs                                          list sheets + headers
 *   read          sheet [from] [to] [formulas]    rows as {header: value}
 *   set           sheet row values{}              update cells in one row
 *   append        sheet values{}                  add a row after the last one
 *   delete_rows   sheet rows[]                    delete rows (returns their contents)
 *   add_column    sheet header                    add a header after the last column
 *   rename_column sheet column to                 rename a header
 * Every write returns the previous contents of what it changed.
 */

const MAX_DELETE = 50;

function doGet() {
  return json_({ ok: true, msg: 'macros sheet api alive (use POST)' });
}

function doPost(e) {
  let p;
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || p.token !== token) return json_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return json_(dispatch_(p));
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function dispatch_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (p.action === 'tabs') {
    return {
      ok: true,
      tabs: ss.getSheets().map(s => ({
        name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn(), headers: headers_(s),
      })),
    };
  }
  const sh = ss.getSheetByName(p.sheet);
  if (!sh) throw new Error('sheet_not_found: ' + p.sheet);
  switch (p.action) {
    case 'read': return read_(sh, p);
    case 'set': return set_(sh, p);
    case 'append': return append_(sh, p);
    case 'delete_rows': return deleteRows_(sh, p);
    case 'add_column': return addColumn_(sh, p);
    case 'rename_column': return renameColumn_(sh, p);
  }
  throw new Error('unknown_action: ' + p.action);
}

// ---------- helpers ----------

function headers_(sh) {
  const n = sh.getLastColumn();
  if (!n) return [];
  return sh.getRange(1, 1, 1, n).getValues()[0].map(h => String(h).trim());
}

function colToLetter_(c) {
  let s = '';
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function letterToCol_(l) {
  let c = 0;
  for (const ch of l) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c;
}

function colIndex_(sh, key) {
  const k = String(key).trim();
  const hs = headers_(sh);
  const hits = [];
  hs.forEach((h, i) => { if (h === k) hits.push(i + 1); });
  if (hits.length === 1) return hits[0];
  if (hits.length > 1) throw new Error('ambiguous_column: ' + k);
  if (/^[A-Z]{1,2}$/.test(k)) return letterToCol_(k);
  throw new Error('unknown_column: ' + k);
}

function colKey_(hs, c) {
  return hs[c - 1] || colToLetter_(c);
}

function subst_(sh, s, row) {
  return s
    .replace(/\{row\}/g, String(row))
    .replace(/\{col:([^}]+)\}/g, (_, h) => colToLetter_(colIndex_(sh, h)));
}

function cellContents_(range) {
  const f = range.getFormula();
  return f ? f : range.getValue();
}

function writeCell_(sh, row, col, v) {
  const r = sh.getRange(row, col);
  const prev = cellContents_(r);
  if (v === null || v === undefined || v === '') r.clearContent();
  else if (typeof v === 'string' && v.charAt(0) === '=') r.setFormula(subst_(sh, v, row));
  else r.setValue(v);
  return prev;
}

function checkRow_(row) {
  const n = Number(row);
  if (!Number.isInteger(n) || n < 2) throw new Error('invalid_row (must be >= 2): ' + row);
  return n;
}

function rowObject_(sh, row) {
  const n = sh.getLastColumn();
  const hs = headers_(sh);
  const range = sh.getRange(row, 1, 1, n);
  const vals = range.getValues()[0];
  const fs = range.getFormulas()[0];
  const o = {};
  for (let c = 1; c <= n; c++) {
    const v = fs[c - 1] || vals[c - 1];
    if (v !== '') o[colKey_(hs, c)] = v;
  }
  return o;
}

// ---------- actions ----------

function read_(sh, p) {
  const last = sh.getLastRow();
  const ncol = sh.getLastColumn();
  const from = Math.max(2, Number(p.from) || 2);
  const to = Math.min(last, Number(p.to) || last);
  const hs = headers_(sh);
  if (to < from || !ncol) return { ok: true, headers: hs, rows: [] };
  const range = sh.getRange(from, 1, to - from + 1, ncol);
  const vals = range.getValues();
  const fs = p.formulas ? range.getFormulas() : null;
  const rows = [];
  vals.forEach((v, i) => {
    if (v.every(x => x === '')) return;
    const values = {};
    const formulas = {};
    v.forEach((x, j) => {
      const key = colKey_(hs, j + 1);
      if (x !== '') values[key] = x;
      if (fs && fs[i][j]) formulas[key] = fs[i][j];
    });
    const r = { row: from + i, values };
    if (fs) r.formulas = formulas;
    rows.push(r);
  });
  return { ok: true, headers: hs, rows };
}

function set_(sh, p) {
  const row = checkRow_(p.row);
  const entries = Object.entries(p.values || {});
  if (!entries.length) throw new Error('no_values');
  const cols = entries.map(([k]) => colIndex_(sh, k)); // validate all before writing
  const previous = {};
  entries.forEach(([k, v], i) => { previous[k] = writeCell_(sh, row, cols[i], v); });
  return { ok: true, action: 'set', row, previous };
}

function append_(sh, p) {
  const entries = Object.entries(p.values || {});
  if (!entries.length) throw new Error('no_values');
  const cols = entries.map(([k]) => colIndex_(sh, k));
  const row = sh.getLastRow() + 1;
  entries.forEach(([, v], i) => writeCell_(sh, row, cols[i], v));
  return { ok: true, action: 'appended', row };
}

function deleteRows_(sh, p) {
  const rows = [...new Set((p.rows || []).map(checkRow_))].sort((a, b) => b - a);
  if (!rows.length) throw new Error('no_rows');
  if (rows.length > MAX_DELETE) throw new Error('too_many_rows (max ' + MAX_DELETE + ')');
  const deleted = rows.map(r => ({ row: r, contents: rowObject_(sh, r) }));
  rows.forEach(r => sh.deleteRow(r));
  return { ok: true, action: 'deleted', deleted };
}

function addColumn_(sh, p) {
  const h = String(p.header || '').trim();
  if (!h) throw new Error('missing_header');
  if (headers_(sh).includes(h)) throw new Error('column_exists: ' + h);
  const col = sh.getLastColumn() + 1;
  sh.getRange(1, col).setValue(h);
  return { ok: true, action: 'column_added', column: colToLetter_(col), header: h };
}

function renameColumn_(sh, p) {
  const col = colIndex_(sh, p.column);
  const to = String(p.to || '').trim();
  if (!to) throw new Error('missing_to');
  if (headers_(sh).includes(to)) throw new Error('column_exists: ' + to);
  const prev = writeCell_(sh, 1, col, to);
  return { ok: true, action: 'renamed', column: colToLetter_(col), previous: prev };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
