/* Разбор исходных файлов месяца. Используется в tools/build-data.js. */
'use strict';

const MONTHS = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
const MONTHS_LABEL = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
const PLATFORMS = ['ВК', 'ОК', 'ТГ', 'МАКС'];
const PLATFORM_ALIASES = {
  'вк': 'ВК', 'вконтакте': 'ВК', 'vk': 'ВК',
  'ок': 'ОК', 'одноклассники': 'ОК', 'ok': 'ОК',
  'тг': 'ТГ', 'telegram': 'ТГ', 'телеграм': 'ТГ', 'tg': 'ТГ',
  'макс': 'МАКС', 'max': 'МАКС'
};
const METRIC_ALIASES = {'просмотры': 'views', 'охваты': 'reach', 'подписчики': 'subs', 'посты': 'posts', 'публикаций': 'posts', 'лайки': 'likes'};

const norm = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
const num = v => {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string' && /^\s*-?[\d\s]+([.,]\d+)?\s*$/.test(v)) return Number(v.replace(/\s/g, '').replace(',', '.'));
  return null;
};

/* ---------- CSV ---------- */
function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const firstLine = text.slice(0, text.indexOf('\n') > 0 ? text.indexOf('\n') : text.length);
  const delim = (firstLine.split(';').length > firstLine.split(',').length) ? ';' : ',';
  const rows = [];
  let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += c;
    } else if (c === '"' && field === '') q = true;
    else if (c === delim) { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); field = '';
      if (row.length > 1 || row[0] !== '') rows.push(row);
      row = [];
    } else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function readSemicolonTable(text) {
  const rows = parseCSV(text);
  const hdr = rows[0].map(h => h.trim());
  return rows.slice(1).filter(r => r.some(x => String(x).trim() !== '')).map(r => {
    const o = {}; hdr.forEach((h, i) => { o[h] = (r[i] || '').trim(); }); return o;
  });
}

/* ---------- Справочник ресурсов ---------- */
function loadResources(text) {
  const rows = readSemicolonTable(text);
  const list = [];
  for (const r of rows) {
    const platform = PLATFORM_ALIASES[norm(r['Площадка'])];
    if (!platform) throw new Error(`resources.csv: неизвестная площадка «${r['Площадка']}»`);
    list.push({
      editorial: r['Редакция'], block: r['Блок'], platform,
      stat: r['Статистика'] ? norm(r['Статистика']) : '',
      links: (r['Ссылки'] || '').split('|').map(norm).filter(Boolean),
      weekly: (r['Съём'] || '').split('|').map(norm).filter(Boolean)
    });
  }
  return list;
}

/* Ключ владельца из ссылки barkov: club123, handle, group/123 */
function ownerKey(url) {
  let s = norm(url).replace(/^https?:\/\//, '').replace(/^(www\.|m\.)/, '');
  s = s.replace(/^(vk\.ru|vk\.com|t\.me|ok\.ru|max\.ru)\//, '').replace(/\/+$/, '');
  s = s.replace(/^public(\d+)$/, 'club$1');
  return s;
}

/* ---------- «Статистика охваты и просмотры» ---------- */
function parseStat(wb, XLSX, year, monthIdx, fileName) {
  const errors = [], warns = [];
  const ws = wb.Sheets[String(year)];
  if (!ws) { errors.push(`${fileName}: нет листа «${year}».`); return {errors, warns, data: null}; }
  const rows = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: null, blankrows: true});
  const mName = MONTHS[monthIdx];
  const data = {ВК: {}, ОК: {}, ТГ: {}, МАКС: {}};
  const found = {ВК: false, ОК: false, ТГ: false, МАКС: false};
  let platform = null, vIdx = -1, rIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    const a = norm(r[0]);
    if (PLATFORM_ALIASES[a]) platform = PLATFORM_ALIASES[a];
    const cells = r.map(norm);
    if (cells.includes('январь')) {
      const idx = cells.reduce((acc, c, k) => (c === mName ? acc.concat(k) : acc), []);
      vIdx = idx.length ? idx[0] : -1;
      rIdx = idx.length > 1 ? idx[1] : -1;
      if (platform == null && cells[0] && PLATFORM_ALIASES[cells[0]]) platform = PLATFORM_ALIASES[cells[0]];
      continue;
    }
    const label = r[1];
    if (!platform || typeof label !== 'string' || !label.trim()) continue;
    if (vIdx < 0) continue;
    found[platform] = true;
    const key = norm(label);
    if (data[platform][key]) warns.push(`${fileName}: в блоке ${platform} строка «${label.trim()}» встречается дважды, взята первая.`);
    else data[platform][key] = {
      views: num(r[vIdx]),
      reach: (platform === 'ВК' || platform === 'ОК') && rIdx >= 0 ? num(r[rIdx]) : null,
      row: i + 1
    };
  }
  for (const p of ['ВК', 'ОК', 'ТГ']) {
    if (!found[p]) errors.push(`${fileName}, лист «${year}»: не найден блок ${p} с колонкой «${mName}».`);
  }
  if (!found['МАКС']) warns.push(`${fileName}: блока МАКС нет — просмотры МАКС только из ручных данных.`);
  const reachCol = Object.values(data['ВК']).some(x => x.reach != null);
  if (!reachCol) warns.push(`${fileName}: в колонках охватов за «${mName}» пусто.`);
  return {errors, warns, data};
}

/* ---------- Недельный съём ---------- */
function parseWeekLabel(text) {
  const m = String(text || '').match(/(\d{1,2})(?:\.(\d{1,2}))?\s*[–—-]\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{4}))?/);
  if (!m) return null;
  return {d1: +m[1], m1: m[2] ? +m[2] : +m[4], d2: +m[3], m2: +m[4], y: m[5] ? +m[5] : null};
}

function parseWeekly(wb, XLSX, year, fileName) {
  const errors = [], warns = [];
  let snapshot = null;
  const sheets = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {header: 1, raw: true, defval: null, blankrows: true});
    for (let i = 0; i < Math.min(rows.length, 8); i++) {
      for (const c of rows[i] || []) {
        const m = String(c || '').match(/снято\s+(\d{2})\.(\d{2})\.(\d{4})/i);
        if (m && !snapshot) snapshot = `${m[3]}-${m[2]}-${m[1]}`;
      }
    }
    const h = rows.findIndex(r => r && norm(r[0]) === 'площадка' && norm(r[1]) === 'ресурс');
    if (h < 0) continue;
    const hdr = rows[h].map(norm);
    if (!hdr.includes('подписчики')) continue;
    const per = parseWeekLabel(rows[0] && rows[0][0]) || parseWeekLabel(name);
    if (!per) continue;
    sheets.push({name, rows, h, hdr, per});
  }
  if (!sheets.length) { errors.push(`${fileName}: не найден лист недели со строкой «Площадка | Ресурс | … | Подписчики».`); return {errors, warns}; }
  const key = s => (s.per.y || year) * 10000 + s.per.m2 * 100 + s.per.d2;
  sheets.sort((a, b) => key(b) - key(a));
  const s = sheets[0];
  const C = s.hdr.indexOf('подписчики');
  const data = {ВК: {}, ОК: {}, ТГ: {}, МАКС: {}};
  let platform = null;
  for (let i = s.h + 1; i < s.rows.length; i++) {
    const r = s.rows[i]; if (!r) continue;
    const a = norm(r[0]), b = norm(r[1]);
    if (a && PLATFORM_ALIASES[a]) platform = PLATFORM_ALIASES[a];
    if (!b || /^итого/.test(b) || !platform) continue;
    data[platform][b] = num(r[C]);
  }
  const y = s.per.y || year;
  const pad = n => String(n).padStart(2, '0');
  const weekEnd = `${y}-${pad(s.per.m2)}-${pad(s.per.d2)}`;
  if (!snapshot) { warns.push(`${fileName}: не найдена дата «снято …», срез считаю по концу недели ${weekEnd}.`); snapshot = weekEnd; }
  return {errors, warns, data, sheet: s.name, weekEnd, snapshot};
}

/* ---------- Выгрузки barkov ---------- */
const BARKOV = {
  ВК: {owner: 'ссылка на владельца стены', date: 'дата публикации', likes: 'лайков'},
  ТГ: {owner: 'ссылка на владельца', date: 'дата публикации', likes: 'лайков'},
  ОК: {owner: 'ссылка на владельца', date: 'дата и время', likes: 'классов'}
};

function parseBarkov(text, platform, monthKey, fileName) {
  const errors = [], warns = [];
  const rows = parseCSV(text);
  if (!rows.length) { errors.push(`${fileName}: файл пустой.`); return {errors, warns}; }
  const hdr = rows[0].map(norm);
  const spec = BARKOV[platform];
  const C = {}, missing = [];
  for (const [k, label] of Object.entries(spec)) { C[k] = hdr.indexOf(label); if (C[k] < 0) missing.push(label.toUpperCase()); }
  if (missing.length) { errors.push(`${fileName}: нет колонок ${missing.map(x => '«' + x + '»').join(', ')}.`); return {errors, warns}; }
  const byOwner = {};
  let outside = 0, total = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length < 2) continue;
    const d = String(r[C.date] || '').trim();
    if (!d) continue;
    total++;
    if (d.slice(0, 7) !== monthKey) { outside++; continue; }
    const k = ownerKey(r[C.owner]);
    const o = byOwner[k] || (byOwner[k] = {posts: 0, likes: 0, first: d, last: d});
    o.posts++;
    o.likes += num(r[C.likes]) || 0;
    if (d < o.first) o.first = d;
    if (d > o.last) o.last = d;
  }
  if (outside) warns.push(`${fileName}: ${outside} из ${total} постов вне месяца ${monthKey}, не учтены.`);
  return {errors, warns, byOwner};
}

/* ---------- Popster: выгрузка постов одного паблика (ВК или ТГ) ---------- */
function parsePopster(wb, XLSX, monthKey, fileName) {
  const errors = [], warns = [];
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: null, blankrows: false});
  const hdr = (rows[0] || []).map(norm);
  const C = {page: hdr.indexOf('page url'), likes: hdr.indexOf('likes'), date: hdr.indexOf('date'), social: hdr.indexOf('social')};
  const missing = Object.entries({page: 'Page url', likes: 'Likes', date: 'Date'}).filter(([k]) => C[k] < 0).map(([, v]) => v);
  if (missing.length) { errors.push(`${fileName}: нет колонок ${missing.map(x => '«' + x + '»').join(', ')}.`); return {errors, warns}; }
  const byOwner = {}; let platform = null, outside = 0, total = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (!r || !r[C.page]) continue;
    const v = r[C.date];
    let d;
    if (v instanceof Date) d = v.toISOString();
    else if (typeof v === 'number') { const t = XLSX.SSF.parse_date_code(v); d = `${t.y}-${String(t.m).padStart(2, '0')}-${String(t.d).padStart(2, '0')} ${String(t.H).padStart(2, '0')}:${String(t.M).padStart(2, '0')}`; }
    else d = String(v || '');
    total++;
    if (d.slice(0, 7) !== monthKey) { outside++; continue; }
    if (C.social >= 0 && !platform) platform = {vk: 'ВК', tg: 'ТГ', ok: 'ОК'}[norm(r[C.social])] || null;
    const k = ownerKey(r[C.page]);
    const o = byOwner[k] || (byOwner[k] = {posts: 0, likes: 0, first: d, last: d});
    o.posts++; o.likes += num(r[C.likes]) || 0;
    if (d < o.first) o.first = d; if (d > o.last) o.last = d;
  }
  if (!platform) platform = /_вк_/i.test(fileName) ? 'ВК' : /_тг_/i.test(fileName) ? 'ТГ' : null;
  if (!platform) errors.push(`${fileName}: не понял площадку (нет колонки «Social» и в имени нет _ВК_/_ТГ_).`);
  if (outside) warns.push(`${fileName}: ${outside} из ${total} постов вне месяца ${monthKey}, не учтены.`);
  return {errors, warns, byOwner, platform};
}

/* ---------- ОК: админский экспорт группы (.xls), посты = строки листа Topics, лайки = сумма LIKES листа Actions ---------- */
function parseOkAdmin(wb, XLSX, monthKey, fileName) {
  const errors = [], warns = [];
  const toKey = v => {
    if (v instanceof Date) return v.toISOString().slice(0, 7);
    if (typeof v === 'number') { const d = XLSX.SSF.parse_date_code(v); return `${d.y}-${String(d.m).padStart(2, '0')}`; }
    return String(v || '').slice(0, 7);
  };
  const sheet = name => { const n = wb.SheetNames.find(x => norm(x) === name); return n ? XLSX.utils.sheet_to_json(wb.Sheets[n], {header: 1, raw: true, defval: null}) : null; };
  const topics = sheet('topics'), actions = sheet('actions');
  if (!topics || !actions) { errors.push(`${fileName}: нет листа ${!topics ? '«Topics»' : '«Actions»'}.`); return {errors, warns}; }
  const th = topics[0].map(norm), ah = actions[0].map(norm);
  const cD = th.indexOf('date'), cU = th.indexOf('url'), cL = ah.indexOf('likes');
  if (cD < 0) errors.push(`${fileName}, лист Topics: нет колонки «DATE».`);
  if (cL < 0) errors.push(`${fileName}, лист Actions: нет колонки «LIKES».`);
  if (errors.length) return {errors, warns};
  let gid = (fileName.match(/ok_group(\d+)/i) || [])[1];
  if (!gid) for (const r of topics.slice(1)) { const m = String(r[cU] || '').match(/group\/(\d+)/); if (m) { gid = m[1]; break; } }
  if (!gid) { errors.push(`${fileName}: не нашёл ID группы ОК (ни в имени файла ok_group…, ни в ссылках постов).`); return {errors, warns}; }
  const toDay = v => {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'number') { const d = XLSX.SSF.parse_date_code(v); return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`; }
    return String(v || '').slice(0, 10);
  };
  let posts = 0, likes = 0, outside = 0, first = null;
  for (const r of topics.slice(1)) {
    if (r[cD] == null) continue;
    if (toKey(r[cD]) === monthKey) { posts++; const d = toDay(r[cD]); if (!first || d < first) first = d; } else outside++;
  }
  for (const r of actions.slice(1)) { if (r[0] == null) continue; if (toKey(r[0]) === monthKey) likes += num(r[cL]) || 0; }
  if (outside) warns.push(`${fileName}: ${outside} постов вне месяца ${monthKey}, не учтены.`);
  return {errors, warns, byOwner: {[gid]: {posts, likes, first: first || monthKey + '-01'}}};
}

/* ---------- Посты МАКС (формат «Итог по каналам») ---------- */
function parseMaxPosts(wb, XLSX, fileName) {
  const errors = [], warns = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {header: 1, raw: true, defval: null, blankrows: true});
    const h = rows.findIndex(r => r && r.map(norm).includes('ресурс') && r.map(norm).includes('публикаций'));
    if (h < 0) continue;
    const hdr = rows[h].map(norm);
    const cR = hdr.indexOf('ресурс'), cP = hdr.indexOf('публикаций'), cL = hdr.indexOf('лайки');
    if (cL < 0) { errors.push(`${fileName}, лист «${name}»: нет колонки «Лайки».`); return {errors, warns}; }
    const data = {};
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r || !r[cR]) continue;
      const k = norm(r[cR]); if (/^итого/.test(k)) continue;
      data[k] = {posts: num(r[cP]), likes: num(r[cL])};
    }
    return {errors, warns, data, sheet: name};
  }
  errors.push(`${fileName}: не найден лист с колонками «Ресурс | Публикаций | … | Лайки».`);
  return {errors, warns};
}

/* ---------- Ручные данные ---------- */
function parseManual(text, fileName) {
  const errors = [], warns = [];
  const rows = readSemicolonTable(text);
  const need = ['Площадка', 'Редакция', 'Метрика', 'Значение', 'Источник'];
  const have = rows.length ? Object.keys(rows[0]) : [];
  const missing = need.filter(n => !have.includes(n));
  if (missing.length) { errors.push(`${fileName}: нет колонок ${missing.map(x => '«' + x + '»').join(', ')}.`); return {errors, warns}; }
  const items = [];
  for (const r of rows) {
    const platform = PLATFORM_ALIASES[norm(r['Площадка'])];
    const metric = METRIC_ALIASES[norm(r['Метрика'])];
    const value = num(r['Значение']);
    if (!platform || !metric || value == null) { warns.push(`${fileName}: строка пропущена (${r['Площадка']} / ${r['Редакция']} / ${r['Метрика']} / ${r['Значение']}).`); continue; }
    items.push({platform, editorial: r['Редакция'], metric, value, source: r['Источник']});
  }
  return {errors, warns, items};
}

module.exports = {MONTHS, MONTHS_LABEL, PLATFORMS, norm, num, parseCSV, readSemicolonTable, loadResources, ownerKey,
  parseStat, parseWeekly, parseBarkov, parsePopster, parseOkAdmin, parseMaxPosts, parseManual};
