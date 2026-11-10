/** @OnlyCurrentDoc */
/**
 * Zarobki – most między Twoim arkuszem Google a aplikacją na telefonie.
 *
 * CZYTA wszystkie zakładki arkusza (nic w nich nie zmienia).
 * ZMIENIA tylko komórkę STATUS sprzedanej rzeczy, gdy klikniesz przełącznik w aplikacji.
 * ZAPISUJE tylko w jednej zakładce „Dashboard – zapis” (zaznaczenia: otrzymane pieniądze, dostawy, cel, konta, pomysły).
 * PYTA Claude (czat i generator opisów) – klucz API zostaje tutaj, w ustawieniach skryptu, nigdy na stronie.
 *
 * Jednorazowo: uruchom funkcję setup() i skopiuj z dziennika (Wykonania) klucz dostępu.
 */
const STATE_SHEET = 'Dashboard – zapis';
const MODEL = 'claude-opus-5-5';

function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('TOKEN')) p.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
  stateSheet_();
  Logger.log('KLUCZ DOSTĘPU (wklej w aplikacji): ' + p.getProperty('TOKEN'));
  Logger.log(p.getProperty('ANTHROPIC_API_KEY') ? 'Klucz Anthropic: ustawiony' : 'Klucz Anthropic: BRAK – dodaj ANTHROPIC_API_KEY w Ustawieniach projektu → Właściwości skryptu, jeśli chcesz czat i generator opisów');
}

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) { let b = {}; try { b = JSON.parse(e.postData.contents); } catch (_) {} return handle_(b); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(q) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || q.t !== token) return json_({error: 'unauthorized'});
  try {
    if (q.a === 'data') return json_({sheets: readSheets_(), state: readState_(), at: Date.now()});
    if (q.a === 'set') { writeState_(String(q.k), q.v); return json_({ok: true}); }
    if (q.a === 'ask') return json_(askClaude_(q));
    if (q.a === 'status') return json_(setStatus_(q));
    return json_({error: 'unknown_action'});
  } catch (err) {
    return json_({error: String(err && err.message || err)});
  }
}

/* ---------- zmiana STATUSU sprzedaży: tylko jedna komórka w kolumnie STATUS, nic więcej ---------- */
function setStatus_(b) {
  const want = String(b.want || ''), pat = {in: /zako[nń]czon/i, wallet: /portfel/i, paying: /pieni[aą]dze w drodze/i, pending: /pieni[aą]dze w drodze/i, shipped: /wys[łl]an/i}[want];
  if (!pat) throw new Error('zły status');
  const sh = SpreadsheetApp.getActive().getSheets().find(s => s.getName().trim() === String(b.sheet || '').trim());
  if (!sh || !/sprzeda/i.test(sh.getName()) || sh.getName() === STATE_SHEET) throw new Error('nie ma takiej zakładki sprzedaży');
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(h => String(h).toUpperCase());
  const sc = head.findIndex(h => h.indexOf('STATUS') >= 0), pc = head.findIndex(h => h.indexOf('PRODUKT') >= 0);
  if (sc < 0 || pc < 0) throw new Error('brak kolumny STATUS');
  const row = Math.floor(+b.row);
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('zły wiersz');
  const prod = String(sh.getRange(row, pc + 1).getDisplayValue()).trim().toLowerCase();
  if (!b.product || prod !== String(b.product).trim().toLowerCase()) throw new Error('wiersz się nie zgadza, odśwież dane');
  const cell = sh.getRange(row, sc + 1);
  let opts = [];
  const dv = cell.getDataValidation();
  if (dv) {
    const ty = dv.getCriteriaType(), v = dv.getCriteriaValues(), C = SpreadsheetApp.DataValidationCriteria;
    if (ty === C.VALUE_IN_LIST) opts = v[0] || [];
    else if (ty === C.VALUE_IN_RANGE && v[0]) opts = [].concat.apply([], v[0].getDisplayValues());
  }
  const def = {in: 'Sprzedaż zakończona', wallet: 'Pieniądze w portfelu Vinted', paying: 'Sprzedane - pieniądze w drodze', pending: 'Sprzedane - pieniądze w drodze', shipped: 'Sprzedane - wysłane'}[want];
  const value = opts.filter(String).find(o => pat.test(o)) || def;
  const old = cell.getDisplayValue();
  if (old !== value) cell.setValue(value);
  return {ok: true, old: old, value: value};
}

/* ---------- odczyt arkusza (tylko odczyt) ---------- */
function readSheets_() {
  const out = {};
  SpreadsheetApp.getActive().getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === STATE_SHEET) return;
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (!lr || !lc) { out[name] = {rows: [], links: {}}; return; }
    const range = sh.getRange(1, 1, lr, lc);
    const disp = range.getDisplayValues();
    // drop empty trailing rows/cols to keep the response small
    let last = disp.length; while (last > 0 && disp[last - 1].every(v => v === '' || v === '0' || v === '0,00 zł')) last--;
    const rows = disp.slice(0, Math.max(last, 1));
    const links = {};
    if (/^\s*\d+\s*$/.test(name) || /zamówienie/i.test(name)) {
      const rt = sh.getRange(1, 1, rows.length, Math.min(lc, 4)).getRichTextValues();
      rt.forEach((row, r) => row.forEach((cell, c) => { const u = cell && cell.getLinkUrl(); if (u) links[r + ',' + c] = u; }));
    }
    out[name] = {rows, links};
  });
  return out;
}

/* ---------- zapis zaznaczeń (tylko zakładka „Dashboard – zapis”) ---------- */
function stateSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(STATE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STATE_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['KLUCZ', 'WARTOŚĆ (zapisuje aplikacja – nie edytuj ręcznie)', 'ZMIENIONO']]).setFontWeight('bold');
    sh.setColumnWidth(2, 600);
  }
  return sh;
}
function readState_() {
  const sh = stateSheet_(), n = sh.getLastRow();
  const st = {};
  if (n < 2) return st;
  sh.getRange(2, 1, n - 1, 2).getValues().forEach(([k, v]) => { if (k) { try { st[k] = JSON.parse(v); } catch (_) {} } });
  return st;
}
function writeState_(key, value) {
  if (!/^[a-z0-9_]{1,40}$/.test(key)) throw new Error('bad_key');
  const text = JSON.stringify(value);
  if (text.length > 45000) throw new Error('too_big');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = stateSheet_(), n = sh.getLastRow();
    const keys = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => r[0]) : [];
    const i = keys.indexOf(key);
    const row = i >= 0 ? i + 2 : n + 1;
    sh.getRange(row, 1, 1, 3).setValues([[key, text, new Date()]]);
  } finally { lock.releaseLock(); }
}

/* ---------- Claude (czat i opisy) ---------- */
function askClaude_(q) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return {error: 'no_key'};
  const messages = Array.isArray(q.messages) ? q.messages.slice(-20) : [{role: 'user', content: String(q.input || '')}];
  const body = {
    model: MODEL,
    max_tokens: q.kind === 'gen' ? 6000 : 4000,
    output_config: {effort: 'low'},
    cache_control: {type: 'ephemeral'},   // stała część rozmowy (zasady + dane) czytana z pamięci podręcznej przy kolejnych pytaniach
    fallbacks: 'default',
    messages: messages
  };
  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {'x-api-key': key, 'anthropic-version': '2026-10-21', 'anthropic-beta': 'server-side-fallback-2026-07-01'},
    payload: JSON.stringify(body)
  });
  const code = r.getResponseCode();
  let j = {}; try { j = JSON.parse(r.getContentText()); } catch (_) {}
  if (code === 429) return {error: 'rate_limited'};
  if (code >= 400) return {error: (j.error && j.error.message) || ('HTTP ' + code)};
  if (j.stop_reason === 'refusal') return {error: 'refusal'};
  const text = (j.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  addUsage_(j.usage || {});
  return {text: text, stop: j.stop_reason, usage: j.usage};
}
function addUsage_(u) {
  const month = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'yyyy-MM');
  const st = readState_().api_usage || {};
  const m = st[month] || {input: 0, output: 0, cache_read: 0, cache_write: 0, calls: 0};
  m.input += u.input_tokens || 0; m.output += u.output_tokens || 0;
  m.cache_read += u.cache_read_input_tokens || 0; m.cache_write += u.cache_creation_input_tokens || 0; m.calls++;
  st[month] = m;
  writeState_('api_usage', st);
}
