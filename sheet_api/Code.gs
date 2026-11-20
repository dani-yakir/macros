/**
 * Macros "dict" sheet API — Google Apps Script web app.
 *
 * Install: open the Google Sheet -> Extensions -> Apps Script, paste this file,
 * set Script Property TOKEN, then Deploy -> New deployment -> Web app
 * (Execute as: Me, Who has access: Anyone). See sheet_api/README.md.
 *
 * All requests are POST with a JSON body:
 *   {"token": "...", "action": "list"}
 *   {"token": "...", "action": "upsert", "name": "...", "unit": 100,
 *    "carbs": 1.2, "protein": 3.4, "fat": 5.6, "calories": 78, "overwrite": false}
 */

const SHEET_NAME = 'dict';
// Columns: A name | B unit | C carbs | D protein | E fat | F calories | G validation formula

function doGet() {
  return json_({ ok: true, msg: 'macros dict api alive (use POST)' });
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

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return json_({ ok: false, error: 'sheet_not_found' });

  if (p.action === 'list') return json_({ ok: true, rows: readRows_(sheet) });

  if (p.action === 'upsert') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      return json_(upsert_(sheet, p));
    } finally {
      lock.releaseLock();
    }
  }

  return json_({ ok: false, error: 'unknown_action' });
}

function readRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 6).getValues()
    .map((v, i) => ({
      row: i + 2, name: String(v[0]).trim(), unit: v[1],
      carbs: v[2], protein: v[3], fat: v[4], calories: v[5],
    }))
    .filter(r => r.name !== '');
}

function upsert_(sheet, p) {
  const name = String(p.name || '').trim();
  if (!name) return { ok: false, error: 'missing_name' };
  if (/^[=+\-@]/.test(name)) return { ok: false, error: 'invalid_name' };

  const nums = {};
  for (const k of ['unit', 'carbs', 'protein', 'fat', 'calories']) {
    const n = Number(p[k]);
    if (p[k] === undefined || p[k] === null || p[k] === '' || !isFinite(n) || n < 0) {
      return { ok: false, error: 'invalid_' + k };
    }
    nums[k] = n;
  }
  if (nums.unit <= 0) return { ok: false, error: 'invalid_unit' };

  const values = [nums.unit, nums.carbs, nums.protein, nums.fat, nums.calories];
  const existing = readRows_(sheet).find(r => r.name === name);

  if (existing) {
    if (!p.overwrite) return { ok: false, error: 'exists', existing };
    sheet.getRange(existing.row, 2, 1, 5).setValues([values]);
    return { ok: true, action: 'updated', row: existing.row, previous: existing };
  }

  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 6).setValues([[name].concat(values)]);
  sheet.getRange(row, 7).setFormula(`=4*C${row} + 4*D${row} + 9*E${row}`);
  return { ok: true, action: 'added', row };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
