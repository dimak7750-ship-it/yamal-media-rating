#!/usr/bin/env node
/* Собирает data/weeks.json из всех файлов съёма в data/raw/*.xlsx.
   Файлы читаются по алфавиту; если одна и та же неделя есть в нескольких файлах,
   берётся версия из файла, который идёт позже (имена вида 2026-09-27.xlsx — по дате конца недели). */
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require('../lib/xlsx.full.min.js');
const P = require('../parser.js');

const root = path.join(__dirname, '..');
const rawDir = path.join(root, 'data', 'raw');
const files = fs.readdirSync(rawDir).filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$')).sort();
if (!files.length) { console.error('В data/raw нет ни одного .xlsx'); process.exit(1); }

const byEnd = new Map();
const warnings = [];
let failed = false;
for (const f of files) {
  const wb = XLSX.read(fs.readFileSync(path.join(rawDir, f)), {type: 'buffer'});
  const res = P.parseWorkbook(wb, XLSX);
  res.errors.forEach(e => { console.error(`ОШИБКА ${f}: ${e}`); failed = true; });
  res.warns.forEach(w => { console.warn(`внимание ${f}: ${w}`); warnings.push(`${f}: ${w}`); });
  for (const w of res.weeks) byEnd.set(P.iso(w.end), {...P.toJSON(w), source: f});
  console.log(`${f}: недель ${res.weeks.length}`);
}
if (failed) { console.error('Сборка остановлена: исправьте файл, чтобы не опубликовать неполные данные.'); process.exit(1); }

const weeks = [...byEnd.values()].sort((a, b) => a.end.localeCompare(b.end));
const out = {updated: new Date().toISOString(), sources: files, warnings, weeks};
fs.writeFileSync(path.join(root, 'data', 'weeks.json'), JSON.stringify(out));
console.log(`data/weeks.json: недель ${weeks.length}, последняя ${weeks[weeks.length - 1].start} — ${weeks[weeks.length - 1].end}`);
