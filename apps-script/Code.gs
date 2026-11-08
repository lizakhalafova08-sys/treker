/**
 * Тихие руки — приём записей дневника в Google-таблицу.
 *
 * Скрипт привязан к таблице (Расширения → Apps Script) и развёрнут как веб-приложение.
 * Приложение отправляет сюда запись (POST, JSON), скрипт кладёт её в строку нужного дня
 * на листе «Дневник». Повторная отправка той же записи обновляет её строку (по ID).
 * Действие list отдаёт все заполненные строки дневника — из них приложение считает итоги.
 *
 * Ключ доступа хранится в свойствах скрипта, а не в коде: запустите функцию setup() один раз
 * и скопируйте ключ из журнала выполнения в настройки приложения.
 */

const SHEET_NAME = 'Дневник';
const HEAD = {
  time: 'Время', kinds: 'Что было', situation: 'Где была', before: 'Позыв, 0', after: 'Позыв после',
  zone: 'Зона', methods: 'Какой метод применила', helped: 'Помогло', could: 'Какой метод мог бы',
  sources: 'Источник', notes: 'Заметки', id: 'ID записи'
};
const DAYS = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

/** Запустить один раз: создаёт ключ доступа и печатает его в журнал. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    props.setProperty('KEY', key);
  }
  Logger.log('Ключ для приложения: ' + key);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkKey_(p.key)) return json_({ ok: false, error: 'bad_key' });
  return json_({ ok: true, sheet: SHEET_NAME, ping: true });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (!checkKey_(body.key)) return json_({ ok: false, error: 'bad_key' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'delete') return json_(deleteEntry_(body.id));
    if (body.action === 'list') return json_(listEntries_());
    return json_(writeEntry_(body.entry));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function checkKey_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('KEY');
  return !!real && key === real;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function layout_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  const hi = values.findIndex(r => r[1] === 'Дата');
  if (hi < 0) throw new Error('На листе «' + SHEET_NAME + '» не найдена строка заголовков');
  const C = {};
  Object.keys(HEAD).forEach(k => { C[k] = values[hi].findIndex(h => String(h).indexOf(HEAD[k]) === 0); });
  ['time', 'before', 'methods', 'id'].forEach(k => { if (C[k] < 0) throw new Error('Нет столбца «' + HEAD[k] + '»'); });
  return { values, hi, C };
}

function ruDate_(iso) { const p = iso.split('-'); return p[2] + '.' + p[1] + '.' + p[0]; }
function text_(v) { v = v == null ? '' : String(v); return /^[=+\-@]/.test(v) ? "'" + v : v; }

function writeEntry_(rec) {
  if (!rec || !rec.id || !rec.date) throw new Error('Пустая запись');
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Нет листа «' + SHEET_NAME + '»');
  const { values, hi, C } = layout_(sh);
  const dataCols = ['time', 'kinds', 'situation', 'before', 'after', 'methods', 'helped', 'could', 'sources', 'notes']
    .map(k => C[k]).filter(i => i >= 0);
  const dd = ruDate_(rec.date);

  let row = values.findIndex((r, i) => i > hi && r[C.id] === rec.id);
  let fresh = null;
  if (row < 0) {
    const same = [];
    values.forEach((r, i) => { if (i > hi && r[1] === dd) same.push(i); });
    const free = same.find(i => dataCols.every(c => !values[i][c]));
    if (free != null) row = free;
    else if (same.length) {
      const last = same[same.length - 1];
      sh.insertRowAfter(last + 1);
      row = last + 1;
      fresh = { week: sh.getRange(last + 1, 1).getValue(), day: values[last][2] };
    } else {
      row = sh.getLastRow();
      const d = new Date(rec.date + 'T12:00:00');
      fresh = { week: '', day: DAYS[d.getDay()] };
    }
  }
  const n = row + 1;
  if (fresh) {
    sh.getRange(n, 1, 1, 3).setValues([[fresh.week, dd, fresh.day]]);
  }
  const bl = colLetter_(C.before);
  const notes = [rec.notes, rec.answers].filter(Boolean).join('\n');
  const val = {
    time: text_(rec.time), kinds: text_((rec.kinds || []).join(', ')), situation: text_(rec.situation),
    before: rec.before == null ? '' : rec.before, after: rec.after == null ? '' : rec.after,
    methods: text_((rec.methods || []).join(', ')), helped: text_(rec.helped), could: text_((rec.couldHelp || []).join(', ')),
    sources: text_((rec.sources || []).join(', ')), notes: text_(notes), id: rec.id
  };
  Object.keys(val).forEach(k => { if (C[k] >= 0) sh.getRange(n, C[k] + 1).setValue(val[k]); });
  if (C.zone >= 0) {
    sh.getRange(n, C.zone + 1).setFormula(
      '=IF(' + bl + n + '="","",IF(' + bl + n + '>=7,"Красная",IF(' + bl + n + '>=4,"Жёлтая","Зелёная")))');
  }
  return { ok: true, row: n };
}

/** Все заполненные строки «Дневника» — для итогов в приложении. */
function listEntries_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const { values, hi, C } = layout_(sh);
  const dataCols = ['time', 'kinds', 'situation', 'before', 'after', 'methods', 'helped', 'could', 'sources', 'notes']
    .map(k => C[k]).filter(i => i >= 0);
  const get = (r, k) => (C[k] >= 0 ? String(r[C[k]] || '').trim() : '');
  const list = s => s ? s.split(',').map(x => x.trim()).filter(Boolean) : [];
  const num = s => { const n = parseFloat(String(s).replace(',', '.')); return isNaN(n) ? null : n; };
  const rows = [];
  values.forEach((r, i) => {
    if (i <= hi) return;
    const m = String(r[1] || '').match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (!m || !dataCols.some(c => String(r[c] || '').trim())) return;
    rows.push({
      id: get(r, 'id') || ('row' + (i + 1)), date: m[3] + '-' + m[2] + '-' + m[1], time: get(r, 'time'),
      kinds: list(get(r, 'kinds')), situation: get(r, 'situation'), before: num(get(r, 'before')), after: num(get(r, 'after')),
      methods: list(get(r, 'methods')), helped: get(r, 'helped'), couldHelp: list(get(r, 'could')),
      sources: list(get(r, 'sources')), notes: get(r, 'notes')
    });
  });
  return { ok: true, rows: rows, at: new Date().toISOString() };
}

function deleteEntry_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const { values, hi, C } = layout_(sh);
  const row = values.findIndex((r, i) => i > hi && r[C.id] === id);
  if (row < 0) return { ok: true, row: null };
  const dd = values[row][1];
  const sameDay = values.filter((r, i) => i > hi && r[1] === dd).length;
  if (sameDay > 1) sh.deleteRow(row + 1);
  else {
    ['time', 'kinds', 'situation', 'before', 'after', 'methods', 'helped', 'could', 'sources', 'notes', 'id']
      .forEach(k => { if (C[k] >= 0) sh.getRange(row + 1, C[k] + 1).clearContent(); });
  }
  return { ok: true, row: row + 1 };
}

function colLetter_(i) {
  let s = ''; i++;
  while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
  return s;
}
