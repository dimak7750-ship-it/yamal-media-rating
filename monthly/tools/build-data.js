#!/usr/bin/env node
/* Собирает data/months.json из папок data/raw/ГГГГ-ММ/.
   В папке месяца ищутся файлы по названию:
     «Статистика…xlsx»                 — просмотры и охваты (обязательно)
     «…съём…xlsx»                      — подписчики на конец месяца (обязательно, если нет в ручных данных)
     «vk.barkov.net-wallposts…csv»     — посты и лайки ВК
     «vk.barkov.net-tgwallposts…csv»   — посты и лайки ТГ
     «vk.barkov.net-okwallposts…csv»   — посты и классы ОК
     «МАКС_посты…xlsx»                 — посты и лайки МАКС (лист «Итог по каналам»)
     «ручные…csv»                      — то, чего нет в файлах (база июля и т. п.); заполняет только пустые клетки
   Если в файле нет нужных колонок, сборка падает, и на сайте остаются прежние данные. */
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require('../../lib/xlsx.full.min.js');
const L = require('./lib.js');

const root = path.join(__dirname, '..');
const rawDir = path.join(root, 'data', 'raw');
const resources = L.loadResources(fs.readFileSync(path.join(root, 'data', 'resources.csv'), 'utf8'));
const notes = L.readSemicolonTable(fs.readFileSync(path.join(root, 'data', 'notes.csv'), 'utf8'))
  .map(n => ({month: n['Месяц'], platform: n['Площадка'], editorial: n['Редакция'], text: n['Пометка']}));

const monthDirs = fs.readdirSync(rawDir).filter(d => /^\d{4}-\d{2}$/.test(d) && fs.statSync(path.join(rawDir, d)).isDirectory()).sort();
if (!monthDirs.length) { console.error('ОШИБКА: в data/raw нет ни одной папки вида ГГГГ-ММ.'); process.exit(1); }

let failed = false;
const err = (m) => { console.error('ОШИБКА ' + m); failed = true; };
const months = [];

for (const mk of monthDirs) {
  const [year, mm] = mk.split('-').map(Number);
  const mIdx = mm - 1;
  const dir = path.join(rawDir, mk);
  // файлы ищутся и во вложенных папках (например, 2026-09/popster/…)
  const walk = (d, pre = '') => fs.readdirSync(d).flatMap(f => {
    if (f.startsWith('~$') || f.startsWith('.')) return [];
    const full = path.join(d, f);
    return fs.statSync(full).isDirectory() ? walk(full, pre + f + '/') : [pre + f];
  });
  const files = walk(dir);
  const warns = [];
  const roles = {stat: [], weekly: [], vk: [], tg: [], ok: [], popster: [], okadmin: [], max: [], manual: [], other: []};
  for (const f of files) {
    const n = path.basename(f).toLowerCase();
    if (/^popster.*\.xlsx$/.test(n)) roles.popster.push(f);
    else if (/ok_group\d+.*\.xls$/.test(n)) roles.okadmin.push(f);
    else if (/tgwallposts/.test(n) && n.endsWith('.csv')) roles.tg.push(f);
    else if (/okwallposts/.test(n) && n.endsWith('.csv')) roles.ok.push(f);
    else if (/wallposts/.test(n) && n.endsWith('.csv')) roles.vk.push(f);
    else if (/статистик/.test(n) && /\.xlsx?$/.test(n)) roles.stat.push(f);
    else if (/съ[её]м/.test(n) && /\.xlsx?$/.test(n)) roles.weekly.push(f);
    else if (/(макс|max)_?посты/.test(n) && /\.xlsx?$/.test(n)) roles.max.push(f);
    else if (/ручн/.test(n) && n.endsWith('.csv')) roles.manual.push(f);
    else if (!/\.(md|txt)$/.test(n)) roles.other.push(f);
  }
  roles.other.forEach(f => warns.push(`Файл «${f}» не распознан и не используется.`));
  const readWb = f => XLSX.read(fs.readFileSync(path.join(dir, f)), {type: 'buffer', cellDates: true});
  const readTxt = f => fs.readFileSync(path.join(dir, f), 'utf8');
  const collect = (res, label) => { (res.errors || []).forEach(e => err(`${mk}: ${e}`)); (res.warns || []).forEach(w => warns.push(w)); return res; };

  // ручные данные
  const manual = [];
  for (const f of roles.manual) { const r = collect(L.parseManual(readTxt(f), f)); if (r.items) r.items.forEach(x => manual.push(x)); }
  const manualHas = metric => manual.some(x => x.metric === metric);

  // Статистика
  if (roles.stat.length !== 1) err(`${mk}: нужен ровно один файл «Статистика охваты и просмотры…», найдено ${roles.stat.length}.`);
  const stat = roles.stat.length ? collect(L.parseStat(readWb(roles.stat[0]), XLSX, year, mIdx, roles.stat[0])) : {data: null};

  // съём
  let weekly = {data: null};
  if (roles.weekly.length > 1) err(`${mk}: недельных съёмов в папке ${roles.weekly.length}, нужен один — ближайший к концу месяца.`);
  else if (roles.weekly.length === 1) weekly = collect(L.parseWeekly(readWb(roles.weekly[0]), XLSX, year, roles.weekly[0]));
  else if (!manualHas('subs')) err(`${mk}: нет недельного съёма (подписчики) и нет подписчиков в ручных данных.`);
  if (weekly.snapshot) {
    const end = new Date(year, mm, 0), snap = new Date(weekly.snapshot);
    const gap = Math.round((snap - end) / 864e5);
    if (Math.abs(gap) > 3) warns.push(`Подписчики сняты ${weekly.snapshot.split('-').reverse().join('.')} — ${Math.abs(gap)} дн. ${gap > 0 ? 'после' : 'до'} конца месяца.`);
  }

  // посты и лайки ВК/ТГ/ОК: barkov, Popster (ВК, ТГ), админские выгрузки ОК.
  // Если по каналу есть несколько выгрузок, берётся одна — с самым ранним постом в месяце
  // (при равенстве — где постов больше). Выгрузки не складываются, чтобы не задвоить посты.
  const barkov = {};
  const pool = {ВК: {}, ТГ: {}, ОК: {}};
  // ключ канала приводится к первой ссылке из справочника, чтобы club123 и короткое имя одного паблика считались одним каналом
  const canon = (p, k) => { const r = resources.find(x => x.platform === p && x.links.includes(k)); return r ? r.links[0] : k; };
  const offer = (p, byOwner, label, file) => {
    for (const [k0, v] of Object.entries(byOwner)) {
      const k = canon(p, k0);
      const o = pool[p][k];
      if (!o || v.first < o.first || (v.first === o.first && v.posts > o.posts)) pool[p][k] = {...v, label, file};
    }
  };
  for (const [p, key] of [['ВК', 'vk'], ['ТГ', 'tg'], ['ОК', 'ok']])
    for (const f of roles[key].slice().sort()) { const r = collect(L.parseBarkov(readTxt(f), p, mk, f)); if (r.byOwner) offer(p, r.byOwner, 'barkov', f); }
  for (const f of roles.popster) {
    const r = collect(L.parsePopster(readWb(f), XLSX, mk, path.basename(f)));
    if (r.byOwner && r.platform) offer(r.platform, r.byOwner, 'Popster', f);
  }
  for (const f of roles.okadmin) {
    const r = collect(L.parseOkAdmin(readWb(f), XLSX, mk, path.basename(f)));
    if (r.byOwner) offer('ОК', r.byOwner, 'админка ОК', f);
  }
  // просмотры прошлого месяца из той же «Статистики» — чтобы отличить обрезанную выгрузку от канала, который начал выходить в этом месяце
  let statPrevM = null;
  if (roles.stat.length && mIdx > 0) { const r = L.parseStat(readWb(roles.stat[0]), XLSX, year, mIdx - 1, roles.stat[0]); statPrevM = r.data; }
  const nameOf = (p, k) => (resources.find(r => r.platform === p && r.links.includes(k)) || {}).editorial || k;
  for (const p of ['ВК', 'ТГ', 'ОК']) {
    const agg = pool[p];
    if (!Object.keys(agg).length) { warns.push(`Нет выгрузки постов ${p} (barkov / Popster${p === 'ОК' ? ' / админка ОК' : ''}): посты, отдача и ER ${p} — н/д.`); continue; }
    const byLabel = {};
    for (const [k, v] of Object.entries(agg)) if (resources.some(r => r.platform === p && r.links.includes(k))) (byLabel[v.label] = byLabel[v.label] || []).push(nameOf(p, k));
    warns.push(`Посты и лайки ${p}: ` + Object.entries(byLabel).map(([l, ks]) => `${l} — ${ks.length} кан` + (Object.keys(byLabel).length > 1 && ks.length <= 5 ? `. (${ks.join(', ')})` : '')).join('; ') + '.');
    // обрезанная выгрузка: канал с сотней и больше постов, первый пост месяца позже 4-го числа, а в прошлом месяце у канала были просмотры
    const cut = [], late = [];
    for (const [k, v] of Object.entries(agg)) {
      const day = +String(v.first).slice(8, 10);
      if (!(v.posts >= 100 && day > 4)) continue;
      const res = resources.find(r => r.platform === p && r.links.includes(k));
      const pv = res && res.stat && statPrevM && statPrevM[p][res.stat] ? statPrevM[p][res.stat].views : null;
      const when = `посты с ${String(v.first).slice(8, 10)}.${String(v.first).slice(5, 7)}, ${v.posts} шт.`;
      if (pv) { cut.push(`${nameOf(p, k)} (${when})`); agg[k] = {posts: null, likes: null, cut: true}; }
      else late.push(`${nameOf(p, k)} (${when}, в прошлом месяце просмотров не было)`);
    }
    if (cut.length) warns.push(`${p}: выгрузка обрезана, начало месяца не попало — посты и лайки н/д: ${cut.join(', ')}.`);
    if (late.length) warns.push(`${p}: канал начал выходить в середине месяца, посты взяты как есть: ${late.join(', ')}.`);
    barkov[p] = {byOwner: agg, used: new Set()};
  }

  // МАКС посты
  let maxPosts = null;
  if (roles.max.length > 1) err(`${mk}: файлов «МАКС_посты» несколько, нужен один за месяц.`);
  else if (roles.max.length === 1) maxPosts = collect(L.parseMaxPosts(readWb(roles.max[0]), XLSX, roles.max[0]));
  else warns.push('Нет файла «МАКС_посты» за месяц: посты, отдача и ER МАКС — н/д.');

  // клетки
  // охваты: если в файле месяца колонка пустая, берём из «Статистики» следующего месяца (только охваты)
  let statNext = null;
  const nextMk = monthDirs[monthDirs.indexOf(mk) + 1];
  if (nextMk) {
    const nd = path.join(rawDir, nextMk);
    const nf = fs.readdirSync(nd).find(f => /статистик/i.test(f) && /\.xlsx?$/i.test(f) && !f.startsWith('~$'));
    if (nf) { const r = L.parseStat(XLSX.read(fs.readFileSync(path.join(nd, nf)), {type: 'buffer'}), XLSX, year, mIdx, nf); if (r.data) statNext = {data: r.data, file: nf}; }
  }
  const reachFilled = [];
  const cells = {};
  const missStat = [], missSubs = [], missPosts = {ВК: [], ТГ: [], ОК: [], МАКС: []};
  for (const r of resources) {
    const c = {views: null, reach: null, subs: null, posts: null, likes: null, src: {}};
    if (stat.data && r.stat) {
      const s = stat.data[r.platform][r.stat];
      if (s) { c.views = s.views; c.src.views = 'Статистика'; if (s.reach != null) { c.reach = s.reach; c.src.reach = 'Статистика'; } }
      else if (Object.keys(stat.data[r.platform]).length) missStat.push(`${r.platform} «${r.stat}»`);
    }
    if (weekly.data) {
      const w = weekly.data[r.platform];
      const k = r.weekly.find(a => a in w);
      if (k != null && w[k] != null) { c.subs = w[k]; c.src.subs = 'съём'; }
      else missSubs.push(`${r.platform} ${r.editorial}`);
    }
    if (barkov[r.platform]) {
      const b = barkov[r.platform];
      const k = r.links.find(l => b.byOwner[l]);
      if (k) { b.used.add(k); if (!b.byOwner[k].cut) { c.posts = b.byOwner[k].posts; c.likes = b.byOwner[k].likes; c.src.posts = c.src.likes = b.byOwner[k].label; } }
      else missPosts[r.platform].push(r.editorial);
    }
    if (r.platform === 'МАКС' && maxPosts && maxPosts.data) {
      const k = r.weekly.find(a => maxPosts.data[a]);
      if (k) { c.posts = maxPosts.data[k].posts; c.likes = maxPosts.data[k].likes; c.src.posts = c.src.likes = 'МАКС_посты'; }
      else missPosts['МАКС'].push(r.editorial);
    }
    for (const m of manual) {
      if (m.platform === r.platform && L.norm(m.editorial) === L.norm(r.editorial) && c[m.metric] == null) {
        c[m.metric] = m.value; c.src[m.metric] = 'ручные: ' + m.source;
      }
    }
    if ((r.platform === 'ВК' || r.platform === 'ОК') && c.reach == null && statNext && r.stat) {
      const s2 = statNext.data[r.platform][r.stat];
      if (s2 && s2.reach != null) { c.reach = s2.reach; c.src.reach = 'Статистика следующего месяца: ' + statNext.file; reachFilled.push(`${r.platform} ${r.editorial}`); }
    }
    if (r.platform === 'ТГ' || r.platform === 'МАКС') c.reach = null;
    cells[`${r.editorial}|${r.platform}`] = c;
  }
  if (reachFilled.length) warns.push(`Охваты взяты из «${statNext.file}», потому что в файле месяца они пустые: ${reachFilled.join(', ')}.`);
  if (missStat.length) warns.push(`В «Статистике» нет строк: ${missStat.join(', ')}.`);
  if (missSubs.length && weekly.data) warns.push(`В съёме нет подписчиков: ${missSubs.join(', ')}.`);
  for (const [p, list] of Object.entries(missPosts)) if (list.length) warns.push(`В выгрузке постов ${p} нет каналов: ${list.join(', ')} — посты н/д.`);
  for (const [p, b] of Object.entries(barkov)) {
    const foreign = Object.entries(b.byOwner).filter(([k]) => !b.used.has(k));
    if (foreign.length) warns.push(`Посты ${p}: не наши владельцы, не учтены — ${foreign.map(([k, v]) => `${k} (${v.posts})`).join(', ')}.`);
  }

  months.push({
    month: mk, label: `${L.MONTHS_LABEL[mIdx]} ${year}`, days: new Date(year, mm, 0).getDate(),
    files: files.filter(f => !roles.other.includes(f) && !roles.popster.includes(f) && !roles.okadmin.includes(f))
      .concat(roles.popster.length ? [`Popster: ${roles.popster.length} файлов`] : [])
      .concat(roles.okadmin.length ? [`админка ОК: ${roles.okadmin.length} файлов`] : []),
    subsSnapshot: weekly.snapshot || null, weekSheet: weekly.sheet || null,
    warnings: warns, cells
  });
  console.log(`${mk}: файлов ${files.length}, предупреждений ${warns.length}`);
  warns.forEach(w => console.warn('  внимание: ' + w));
}

if (failed) { console.error('Сборка остановлена: исправьте файлы, чтобы не опубликовать неполные данные.'); process.exit(1); }

const editorials = [];
for (const r of resources) {
  let e = editorials.find(x => x.name === r.editorial);
  if (!e) editorials.push(e = {name: r.editorial, block: r.block, platforms: []});
  e.platforms.push(r.platform);
}
// пометки и предупреждения на публичную страницу не выводятся — они остаются в логе сборки (вкладка Actions) и в data/notes.csv
const out = {updated: new Date().toISOString(), editorials, notes: [], months: months.map(m => ({...m, warnings: [], files: []}))};
fs.writeFileSync(path.join(root, 'data', 'months.json'), JSON.stringify(out));
console.log(`data/months.json: месяцев ${months.length} (${months[0].month} … ${months[months.length - 1].month})`);
