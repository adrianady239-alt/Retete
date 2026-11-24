/**
 * Rețete – server (Google Apps Script).
 * 1. Creează un fișier nou în Google Sheets, numește-l „Retete”.
 * 2. Extensii → Apps Script, șterge tot, lipește codul ăsta.
 * 3. Pune parola ta la TOKEN (linia de mai jos), Ctrl+S.
 * 4. Implementare → Implementare nouă → Aplicație web (Execută ca: Eu, Acces: Oricine).
 * 5. URL-ul /exec + parola se pun în aplicație la Setări → Server.
 * Foile se creează singure la prima rulare.
 */
const TOKEN = 'SCHIMBA_PAROLA';

const SHEETS = {
  retete:  ['id', 'nume', 'emoji', 'categorie', 'portii', 'timp', 'ingrediente', 'pasi',
            'etichete', 'favorit', 'nota', 'creat'],
  plan:    ['id', 'data', 'masa', 'retetaId', 'nume', 'portii', 'mancat'],
  preturi: ['id', 'produs', 'pret', 'cantitate', 'unitate', 'magazin', 'data']
};
const DATA_SHEET = 'APP_DATA';

function doGet() {
  return out({ ok: true, msg: 'Rețete – server activ. Folosește URL-ul în aplicație.' });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    if (String(p.token || '') !== TOKEN) {
      return out({ ok: false, code: 'bad_token', error: 'Parolă greșită.' });
    }
    switch (p.action) {
      case 'ping':   return out({ ok: true, rev: rev() });
      case 'all':    return out({ ok: true, rev: rev(),
                                  retete: readRows('retete'),
                                  plan: readRows('plan'),
                                  preturi: readRows('preturi'),
                                  cfg: docData('cfg'), lista: docData('lista') });
      case 'upsert': return out(withLock(function () { return upsert(p.sheet, p.rows || []); }));
      case 'del':    return out(withLock(function () { return delRows(p.sheet, p.ids || []); }));
      case 'doc':    return out(withLock(function () { putDoc(p.id, p.data); return { ok: true }; }));
      case 'ocr':    return out(ocrBon(p.b64, p.mime));
      default: return out({ ok: false, code: 'tool_error', error: 'Acțiune necunoscută: ' + p.action });
    }
  } catch (err) {
    return out({ ok: false, code: 'tool_error', error: String((err && err.message) || err) });
  }
}

/* ---------- foi ---------- */
function sheetFor(key) {
  const head = SHEETS[key];
  if (!head) throw new Error('Foaie necunoscută: ' + key);
  const name = key.charAt(0).toUpperCase() + key.slice(1);
  let sh = ss().getSheetByName(name);
  if (!sh) {
    sh = ss().insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
  }
  return sh;
}
function colMap(sh, head) {
  const n = Math.max(sh.getLastColumn(), head.length);
  const row = sh.getRange(1, 1, 1, n).getDisplayValues()[0];
  const map = {};
  for (let i = 0; i < row.length; i++) {
    const k = String(row[i] || '').trim().toLowerCase();
    if (k) map[k] = i;
  }
  for (let i = 0; i < head.length; i++) if (!(head[i] in map)) map[head[i]] = i;
  return map;
}
function readRows(key) {
  const head = SHEETS[key], sh = sheetFor(key), last = sh.getLastRow();
  if (last < 2) return [];
  const map = colMap(sh, head);
  const width = Math.max(sh.getLastColumn(), head.length);
  const vals = sh.getRange(2, 1, last - 1, width).getDisplayValues();
  const res = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i], o = {};
    let gol = true;
    for (let c = 0; c < head.length; c++) {
      const v = r[map[head[c]]];
      o[head[c]] = (v == null) ? '' : String(v);
      if (o[head[c]] !== '') gol = false;
    }
    if (!gol && o.id) res.push(o);
  }
  return res;
}
function upsert(key, rows) {
  if (!Array.isArray(rows) || !rows.length) return { ok: true, scrise: 0 };
  const head = SHEETS[key], sh = sheetFor(key), map = colMap(sh, head);
  const last = sh.getLastRow();
  const width = Math.max(sh.getLastColumn(), head.length);
  const ids = last < 2 ? [] :
    sh.getRange(2, map.id + 1, last - 1, 1).getDisplayValues().map(function (x) { return String(x[0]); });
  let add = 0, mod = 0;
  for (let i = 0; i < rows.length; i++) {
    const src = rows[i] || {};
    if (!src.id) continue;
    const line = new Array(width).fill('');
    for (let c = 0; c < head.length; c++) {
      const v = src[head[c]];
      line[map[head[c]]] = (v == null) ? '' : String(v);
    }
    const at = ids.indexOf(String(src.id));
    if (at >= 0) { sh.getRange(at + 2, 1, 1, width).setValues([line]); mod++; }
    else { sh.getRange(sh.getLastRow() + 1, 1, 1, width).setValues([line]); ids.push(String(src.id)); add++; }
  }
  SpreadsheetApp.flush();
  bump();
  return { ok: true, adaugate: add, modificate: mod };
}
function delRows(key, ids) {
  if (!Array.isArray(ids) || !ids.length) return { ok: true, sterse: 0 };
  const head = SHEETS[key], sh = sheetFor(key), map = colMap(sh, head), last = sh.getLastRow();
  if (last < 2) return { ok: true, sterse: 0 };
  const col = sh.getRange(2, map.id + 1, last - 1, 1).getDisplayValues();
  const kill = [];
  for (let i = 0; i < col.length; i++) if (ids.indexOf(String(col[i][0])) >= 0) kill.push(i + 2);
  kill.sort(function (a, b) { return b - a; });
  for (let i = 0; i < kill.length; i++) sh.deleteRow(kill[i]);
  if (kill.length) { SpreadsheetApp.flush(); bump(); }
  return { ok: true, sterse: kill.length };
}

/* ---------- setări (foaia ascunsă APP_DATA) ---------- */
function dataSheet() {
  let sh = ss().getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss().insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['id', 'json', 'updatedAt']]);
    sh.hideSheet();
  }
  return sh;
}
function allDocs() {
  const sh = dataSheet(), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, 3).getValues();
}
function docData(id) {
  const rows = allDocs();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === id) { try { return JSON.parse(rows[i][1] || '{}'); } catch (e) { return null; } }
  }
  return null;
}
function putDoc(id, data) {
  if (!id) throw new Error('Lipsește id.');
  const sh = dataSheet(), rows = allDocs();
  let at = -1;
  for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === id) { at = i; break; }
  const line = [[id, JSON.stringify(data || {}), new Date()]];
  if (at >= 0) sh.getRange(at + 2, 1, 1, 3).setValues(line);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues(line);
  bump();
}

/* ---------- backup lunar (Declanșatoare → backupLunar, Cronometru lunar) ---------- */
function backupLunar() {
  const file = DriveApp.getFileById(ss().getId());
  const name = 'Retete backup';
  const it = DriveApp.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  const stamp = Utilities.formatDate(new Date(), ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  file.makeCopy('Retete ' + stamp, folder);
  const all = [], files = folder.getFiles();
  while (files.hasNext()) { const f = files.next(); all.push({ f: f, d: f.getDateCreated().getTime() }); }
  all.sort(function (a, b) { return b.d - a.d; });
  for (let i = 12; i < all.length; i++) all[i].f.setTrashed(true);
  return all.length;
}

/* ---------- utilitare ---------- */
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function rev() { return Number(PropertiesService.getScriptProperties().getProperty('rev') || 0); }
function bump() { PropertiesService.getScriptProperties().setProperty('rev', String(rev() + 1)); }
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- citirea bonului: OCR gratuit prin Google Drive ----------
   Necesită serviciul avansat Drive: în Apps Script, Servicii (+) → Drive API → Adaugă.
   Poza se urcă temporar în Drive, se convertește în document (OCR), se citește textul
   și fișierele se șterg imediat. Nu rămâne nimic în Drive. */
function ocrBon(b64, mime) {
  let docId = null, imgId = null;
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg', 'bon.jpg');
    const copie = Drive.Files.insert(
      { title: 'bon-temp-' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
      blob,
      { ocr: true, ocrLanguage: 'ro' }
    );
    docId = copie.id;
    const text = DocumentApp.openById(docId).getBody().getText();
    return { ok: true, text: text, randuri: parseBon(text) };
  } catch (err) {
    return { ok: false, code: 'tool_error',
      error: 'OCR-ul nu a mers: ' + String((err && err.message) || err) +
             ' (verifică dacă ai adăugat serviciul Drive API în Apps Script)' };
  } finally {
    try { if (docId) Drive.Files.remove(docId); } catch (e) {}
    try { if (imgId) Drive.Files.remove(imgId); } catch (e) {}
  }
}

// din textul bonului scot liniile „produs ... 12,34"
const BON_SARI = /(total|subtotal|tva|rest|numerar|card|bon fiscal|cui|c\.u\.i|cif|nr\.|casier|data|ora|multumim|mulțumim|puncte|reducere|discount|cod|serie|amef|client|adresa|adresă|str\.|tel)/i;
function parseBon(text) {
  const linii = String(text || '').split('\n');
  const out = [];
  for (let i = 0; i < linii.length; i++) {
    const l = linii[i].trim();
    if (l.length < 4) continue;
    if (BON_SARI.test(l)) continue;
    // ultimul număr cu zecimale de pe linie = prețul
    const m = l.match(/^(.*?)[\s.]*(\d{1,4}[.,]\d{2})\s*(?:[A-D]|lei|ron)?\s*$/i);
    if (!m) continue;
    let nume = m[1].replace(/[*x×]\s*\d+([.,]\d+)?\s*$/i, '')   // „2 x" la final
                   .replace(/\b\d+([.,]\d+)?\s*(buc|kg|g|l|ml)\b/gi, function (s2) { return s2; })
                   .replace(/[^\p{L}\p{N}\s.,%-]/gu, ' ')
                   .replace(/\s{2,}/g, ' ').trim();
    if (nume.length < 3) continue;
    if (/^\d+$/.test(nume)) continue;
    const pret = Number(m[2].replace(',', '.'));
    if (!(pret > 0) || pret > 5000) continue;
    // cantitatea din ambalaj, dacă apare în nume (ex. „LAPTE 1L", „FAINA 1KG")
    let cant = 0, unit = '';
    const c = nume.match(/(\d+([.,]\d+)?)\s*(kg|g|l|ml|buc)\b/i);
    if (c) { cant = Number(c[1].replace(',', '.')); unit = c[3].toLowerCase(); }
    out.push({ nume: nume.toLowerCase(), pret: pret, cant: cant, unit: unit });
  }
  return out;
}
