/* Разбор файла еженедельного съёма (.xlsx) — общий для страницы и для сборки data/weeks.json.
   Работает и в браузере (window.YMParser), и в Node (module.exports). */
(function (root) {
  'use strict';

  const PLATFORM_KEYS = ['ВКонтакте', 'МАКС', 'Telegram'];
  const PLATFORM_ALIASES = {'вконтакте':'ВКонтакте','вк':'ВКонтакте','telegram':'Telegram','телеграм':'Telegram','тг':'Telegram','макс':'МАКС','max':'МАКС'};
  const COLS = {pubs:'публикаций', views:'просмотры', likes:'лайки', subs:'подписчики', growth:'прирост подписчиков'};
  const METRIC_RU = {views:'просмотры', likes:'лайки', subs:'подписчики', growth:'прирост подписчиков'};

  const norm = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
  const num = v => (typeof v === 'number' && isFinite(v)) ? v : null;
  const fmt = v => Math.round(v).toLocaleString('ru-RU');
  const pad = n => String(n).padStart(2, '0');
  const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const fromIso = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };

  function parseWeekLabel(text, sheetName) {
    const re = /(\d{1,2})(?:\.(\d{1,2}))?\s*[–—-]\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{4}))?/;
    const m = String(text || '').match(re) || String(sheetName || '').match(re);
    if (!m) return null;
    const d2 = +m[3], m2 = +m[4], y = m[5] ? +m[5] : new Date().getFullYear();
    const d1 = +m[1], m1 = m[2] ? +m[2] : m2;
    const y1 = m1 > m2 ? y - 1 : y;
    return {start: new Date(y1, m1 - 1, d1), end: new Date(y, m2 - 1, d2)};
  }

  function parseWorkbook(wb, XLSX) {
    const weeks = [], errors = [], warns = [];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: null, blankrows: true});
      const title = rows[0] && rows[0][0];
      const isWeekByName = /^\d{1,2}(\.\d{1,2})?\s*[-–]\s*\d{1,2}\.\d{1,2}$/.test(name.trim());
      const isWeekByTitle = /съ[её]м/i.test(String(title || '')) && parseWeekLabel(title, '') != null;
      if (!isWeekByName && !isWeekByTitle) continue;

      const h = rows.findIndex(r => r && norm(r[0]) === 'площадка' && norm(r[1]) === 'ресурс');
      if (h < 0) { errors.push(`Лист «${name}»: не найдена строка заголовков «Площадка | Ресурс».`); continue; }
      const hdr = rows[h].map(norm);
      const C = {}, missing = [];
      for (const [k, label] of Object.entries(COLS)) { C[k] = hdr.indexOf(label); if (C[k] < 0) missing.push(label); }
      if (missing.length) { errors.push(`Лист «${name}»: нет колонок ${missing.map(x => '«' + x + '»').join(', ')}.`); continue; }

      const period = parseWeekLabel(title, name);
      if (!period) { errors.push(`Лист «${name}»: не удалось понять даты недели.`); continue; }

      const items = [], totals = {}, empty = [];
      let currentPlatform = null;
      for (let i = h + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const a = r[0] != null ? String(r[0]).trim() : '';
        const b = r[1] != null ? String(r[1]).trim() : '';
        if (!b) continue;
        if (/^итого/i.test(b)) {
          const p = PLATFORM_ALIASES[norm(b.replace(/^итого\s*/i, ''))];
          if (p) totals[p] = {views: num(r[C.views]), likes: num(r[C.likes]), subs: num(r[C.subs]), growth: num(r[C.growth])};
          continue;
        }
        if (a) currentPlatform = PLATFORM_ALIASES[norm(a)] || a;
        const platform = currentPlatform;
        if (!PLATFORM_KEYS.includes(platform)) continue;
        const rec = {platform, name: b, pubs: num(r[C.pubs]), views: num(r[C.views]), likes: num(r[C.likes]), subs: num(r[C.subs]), growth: num(r[C.growth])};
        if (rec.views == null && rec.subs == null) { empty.push(`${platform} · ${b}`); continue; }
        rec.er = (rec.likes != null && rec.subs) ? rec.likes / rec.subs / 7 * 100 : null;
        items.push(rec);
      }
      // сверка с итоговыми строками листа
      for (const p of PLATFORM_KEYS) {
        const t = totals[p]; if (!t) continue;
        for (const k of ['views', 'likes', 'subs', 'growth']) {
          if (t[k] == null) continue;
          const sum = items.filter(x => x.platform === p).reduce((s, x) => s + (x[k] || 0), 0);
          if (Math.abs(sum - t[k]) > 0.5) warns.push(`Лист «${name}», ${p}: сумма «${METRIC_RU[k]}» по строкам (${fmt(sum)}) не совпадает с итогом (${fmt(t[k])}).`);
        }
      }
      weeks.push({sheet: name, title: String(title || name), start: period.start, end: period.end, items, empty});
    }
    weeks.sort((a, b) => a.end - b.end);

    // текущая неделя — из ячейки B2 листа «Сводка», иначе последняя
    let def = null;
    const sv = wb.SheetNames.find(n => norm(n) === 'сводка');
    if (sv) {
      const c = wb.Sheets[sv]['B2'];
      if (c && c.v != null) def = weeks.find(w => norm(w.sheet) === norm(c.v)) || null;
    }
    return {weeks, def: def || weeks[weeks.length - 1] || null, errors, warns};
  }

  // недели <-> JSON (даты как YYYY-MM-DD)
  const toJSON = w => ({...w, start: iso(w.start), end: iso(w.end)});
  const fromJSON = w => ({...w, start: fromIso(w.start), end: fromIso(w.end)});

  const api = {parseWorkbook, parseWeekLabel, toJSON, fromJSON, iso, PLATFORM_KEYS};
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.YMParser = api;
})(typeof window !== 'undefined' ? window : globalThis);
