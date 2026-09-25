/* 金鶯診所˙醫師回診率分析系統 — 共用計算核心 engine.js
 * 前台 index.html 與後台 admin.html 共用；不需修改。
 */
(function (root) {
  'use strict';

  /* ---------------- 基本工具 ---------------- */
  const normCode = c => String(c == null ? '' : c).trim().replace(/^0+(?=\d)/, '');
  const num = v => { if (v === null || v === undefined || v === '') return 0; const n = Number(String(v).replace(/[,\s]/g, '')); return isNaN(n) ? 0 : n; };
  const isBlank = v => v === null || v === undefined || v === '' || (typeof v === 'number' && isNaN(v));
  const numOrNull = v => { if (isBlank(v)) return null; const n = Number(String(v).replace(/[,\s%]/g, '')); return isNaN(n) ? null : n; };
  const r2 = x => Math.round(x * 100) / 100;
  const r4 = x => Math.round(x * 10000) / 10000;

  // 民國 1150806 → 天數序號（UTC 日）
  function rocDay(s) {
    s = String(s == null ? '' : s).trim();
    if (!/^\d{7}$/.test(s)) return null;
    const y = +s.slice(0, 3) + 1911, m = +s.slice(3, 5), d = +s.slice(5, 7);
    if (m < 1 || m > 12 || d < 1 || d > 31) return null;
    return Date.UTC(y, m - 1, d) / 86400000;
  }
  const isoDay = s => { const m = String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 : null; };
  const dayIso = d => new Date(d * 86400000).toISOString().slice(0, 10);
  const dayRoc = d => { const t = new Date(d * 86400000); return (t.getUTCFullYear() - 1911) + '/' + String(t.getUTCMonth() + 1).padStart(2, '0') + '/' + String(t.getUTCDate()).padStart(2, '0'); };

  function monthLabel(m) { // 2026-08 → 115年8月
    const [y, mo] = String(m).split('-').map(Number);
    return (y - 1911) + '年' + mo + '月';
  }
  function prevMonth(m) { const [y, mo] = m.split('-').map(Number); const d = new Date(Date.UTC(y, mo - 2, 1)); return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0'); }
  function monthRange(m) { const [y, mo] = m.split('-').map(Number); const s = Date.UTC(y, mo - 1, 1) / 86400000; const e = Date.UTC(y, mo, 0) / 86400000; return [s, e]; }

  /* ---------------- 讀 Excel 列（二維陣列） ---------------- */
  function findHeader(rows, keys) {
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const cells = (rows[i] || []).map(c => String(c == null ? '' : c).trim());
      if (keys.every(k => cells.indexOf(k) >= 0)) return i;
    }
    return -1;
  }
  function headerIndex(row) { const o = {}; (row || []).forEach((c, i) => { const k = String(c == null ? '' : c).trim(); if (k && o[k] === undefined) o[k] = i; }); return o; }

  /** 檔名 OO_看診名單：排除「說明含還卡日」「欠款有金額」（含其對沖列），其餘病患資料皆辨識 */
  function parseVisitRows(rows) {
    const h = findHeader(rows, ['證號', '日期', '醫']);
    if (h < 0) throw new Error('看診名單找不到欄位名稱列（需有「證號、日期、醫」）');
    const H = headerIndex(rows[h]);
    const g = (r, k) => H[k] === undefined ? '' : r[H[k]];
    const all = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const id = String(g(r, '證號') == null ? '' : g(r, '證號')).trim();
      if (!/^\d+$/.test(id)) continue;                 // 小計、總計、統計表等非病患列
      const pid = id.replace(/^0+(?=\d)/, '');
      const day = rocDay(g(r, '日期'));
      if (day === null) continue;
      all.push({ id: pid, name: String(g(r, '姓名') || '').trim(), day, code: normCode(g(r, '醫')), note: String(g(r, '說明') == null ? '' : g(r, '說明')),
                 debt: num(g(r, '欠款')), sub: num(g(r, '小計')), reg: num(g(r, '掛號費')), line: i + 1 });
    }
    const ex = new Array(all.length).fill('');
    all.forEach((v, i) => {
      if (v.note.indexOf('還卡') >= 0) ex[i] = '還卡';
      else if (v.debt !== 0) ex[i] = '欠款';
    });
    // 欠款列常伴隨一筆金額相反、欠款為 0 的對沖列（如 -2400 / +2400），一併排除
    all.forEach((v, i) => {
      if (ex[i] !== '欠款') return;
      const j = all.findIndex((w, k) => !ex[k] && w.id === v.id && w.day === v.day && w.debt === 0 && w.sub === -v.debt && w.reg === 0);
      if (j >= 0) ex[j] = '欠款對沖';
    });
    const kept = all.filter((_, i) => !ex[i]);
    const excluded = all.map((v, i) => ex[i] ? Object.assign({ reason: ex[i] }, v) : null).filter(Boolean);
    const days = kept.map(v => v.day);
    return { rows: kept, excluded, totalRows: all.length,
             stats: { 還卡: ex.filter(x => x === '還卡').length, 欠款: ex.filter(x => x === '欠款').length, 欠款對沖: ex.filter(x => x === '欠款對沖').length },
             minDay: days.length ? Math.min.apply(null, days) : null, maxDay: days.length ? Math.max.apply(null, days) : null };
  }

  /** 檔名 OO_BACK：欄位名稱列以下皆為資料（前 1-3 列抬頭忽略） */
  function parseBackRows(rows) {
    const h = findHeader(rows, ['病歷號', '看診日期', '醫師別']);
    if (h < 0) throw new Error('BACK 檔找不到欄位名稱列（需有「病歷號、看診日期、醫師別」）');
    const H = headerIndex(rows[h]);
    const g = (r, k) => H[k] === undefined ? '' : r[H[k]];
    const out = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const id = String(g(r, '病歷號') == null ? '' : g(r, '病歷號')).trim();
      if (!/^\d+$/.test(id)) continue;
      const day = rocDay(g(r, '看診日期'));
      if (day === null) continue;
      const pid = id.replace(/^0+(?=\d)/, '');
      if (H['藥代'] !== undefined) { const dc = String(g(r, '藥代') || '').toUpperCase(); if (dc && dc.indexOf('BACK') < 0) continue; }
      out.push({ id: pid, name: String(g(r, '姓名') || '').trim(), day, code: normCode(g(r, '醫師別')), line: i + 1 });
    }
    return { rows: out };
  }

  /* ---------------- 院區計算 ---------------- */
  /**
   * @param visits  parseVisitRows().rows
   * @param backs   parseBackRows().rows
   * @param opt { clinic, visitMap:[{clinic,code,doctor,active}], backMap:[...], doctors:[name], days:5, stripFirst:true }
   */
  function computeClinic(visits, backs, opt) {
    const days = +opt.days || 5;
    const vmap = {}, bmap = {};
    (opt.visitMap || []).filter(x => x.clinic === opt.clinic && x.active !== false).forEach(x => vmap[normCode(x.code)] = x.doctor);
    (opt.backMap || []).filter(x => x.clinic === opt.clinic && x.active !== false).forEach(x => bmap[normCode(x.code)] = x.doctor);
    const docOf = code => vmap[code] || null;
    const backDocOf = code => {
      if (bmap[code]) return bmap[code];
      if (opt.stripFirst !== false && code.length > 1) { const c2 = normCode(code.slice(1)); if (vmap[c2]) return vmap[c2]; }
      return null;
    };
    const res = {}; const unmappedV = {}, unmappedB = {};
    const doctorSet = new Set(opt.doctors || []);
    const ensure = n => res[n] || (res[n] = { patients: 0, backN: 0, nRevN: 0, wRevN: 0 });
    // 索引：病患|醫師 → 看診日
    const idx = {};
    visits.forEach(v => {
      const d = docOf(v.code);
      const key = v.id + '|' + (d || ('#' + v.code));
      (idx[key] = idx[key] || []).push(v.day);
      if (!d) { unmappedV[v.code] = (unmappedV[v.code] || 0) + 1; return; }
      if (d === '不列入') return;
      ensure(d).patients++;
    });
    const hasRevisit = (key, day) => { const a = idx[key]; if (!a) return false; for (let i = 0; i < a.length; i++) { const dd = a[i] - day; if (dd >= 1 && dd <= days) return true; } return false; };
    // 廣義（總）回診：同一病患 N 天內再回同一醫師
    visits.forEach(v => {
      const d = docOf(v.code); if (!d || d === '不列入') return;
      if (hasRevisit(v.id + '|' + d, v.day)) ensure(d).wRevN++;
    });
    // BACK 開立數與狹義回診
    backs.forEach(b => {
      const d = backDocOf(b.code);
      if (!d) { unmappedB[b.code] = (unmappedB[b.code] || 0) + 1; return; }
      if (d === '不列入') return;
      const o = ensure(d); o.backN++;
      if (hasRevisit(b.id + '|' + d, b.day)) o.nRevN++;
    });
    const outside = Object.keys(res).filter(n => !doctorSet.has(n));
    return { doctors: res, unmappedVisit: unmappedV, unmappedBack: unmappedB, outside };
  }

  /** 由原始數字算比例（數值皆取至小數點後兩位；比例以小數存，0.1580 = 15.80%） */
  function metrics(o) {
    const hours = numOrNull(o.hours), patients = numOrNull(o.patients), visits = numOrNull(o.visits);
    const slowN = numOrNull(o.slowN), backN = numOrNull(o.backN), nRevN = numOrNull(o.nRevN), wRevN = numOrNull(o.wRevN);
    const out = Object.assign({}, o, { hours, patients, visits, slowN, backN, nRevN, wRevN });
    out.slowRate = visits && slowN !== null ? r4(slowN / visits) : null;
    out.vph = hours && patients !== null ? r2(patients / hours) : null;
    out.backRate = visits && backN !== null ? r4(backN / visits) : null;
    out.nRevRate = backN && nRevN !== null ? r4(nRevN / backN) : null;
    out.wRevRate = visits && wRevN !== null ? r4(wRevN / visits) : null;
    return out;
  }

  /* ---------------- 班表 → 上診時數 ---------------- */
  const WEEK = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'];
  const SKIP = new Set(['休診', '~', '～', '', 'TRUE', 'FALSE', '彩色版', '系統', '確認ok', '異動ok'].concat(WEEK));
  function timeMin(v) {
    let s = String(v == null ? '' : v).trim();
    if (s.indexOf('T:') === 0) s = s.slice(2);
    const m = s.match(/^(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
    if (!m) return null;
    let h = +m[2]; const ap = (m[1] || m[4] || '').toUpperCase();
    if ((ap === '下午' || ap === 'PM') && h < 12) h += 12;
    if ((ap === '上午' || ap === 'AM') && h === 12) h = 0;
    return h * 60 + (+m[3]);
  }
  function timeRange(v) {
    const s = String(v == null ? '' : v).trim().replace(/^T:/, '');
    const m = s.match(/^(\d{1,2}):(\d{2})\s*[-~～－–]\s*(\d{1,2}):(\d{2})$/);
    return m ? [+m[1] * 60 + (+m[2]), +m[3] * 60 + (+m[4])] : null;
  }
  const looksName = s => s && !SKIP.has(s) && s.indexOf('D:') !== 0 && s.indexOf('T:') !== 0 && !/\d/.test(s) && s.length >= 2 && s.length <= 5 && !/[~～:：\/()（）]/.test(s);

  /**
   * @param sheet {grid:[[string]], merges:[[r,c,rows,cols]]}  grid 由後台 readSchedule 傳回（日期 'D:yyyy-mm-dd'、時間 'T:7:30'）
   * @param opt {start, end (天數序號), exclude:['減重']}
   * @return {hours:{醫師:時數}, detail:[{day,name,start,end,section}]}
   */
  function parseSchedule(sheet, opt) {
    const g = sheet.grid.map(r => r.map(v => v == null ? '' : String(v)));
    const R = g.length, C = Math.max.apply(null, g.map(r => r.length).concat([15]));
    g.forEach(r => { while (r.length < C) r.push(''); });
    (sheet.merges || []).forEach(([r0, c0, nr, nc]) => {
      const v = g[r0][c0];
      for (let r = r0; r < r0 + nr && r < R; r++) for (let c = c0; c < c0 + nc && c < C; c++) g[r][c] = v;
    });
    const exclude = (opt.exclude || []).filter(Boolean);
    // 週區塊：含「星期一」的列
    const blocks = [];
    for (let r = 0; r < R; r++) {
      if (g[r].slice(1, 15).some(v => v.trim() === '星期一')) blocks.push({ wr: r });
    }
    blocks.forEach(b => {
      const dates = {};
      for (let k = 0; k < 7; k++) {
        const c = 1 + k * 2;
        for (const rr of [b.wr - 1, b.wr - 2]) {
          if (rr >= 0 && g[rr][c].indexOf('D:') === 0) { dates[k] = isoDay(g[rr][c].slice(2)); break; }
        }
      }
      const ks = Object.keys(dates);
      if (ks.length) {           // 補齊缺漏的日期
        const k0 = +ks[0], d0 = dates[k0];
        for (let k = 0; k < 7; k++) if (dates[k] === undefined) dates[k] = d0 + (k - k0);
        b.dated = true; b.dates = dates;
      } else b.dated = false;
    });
    blocks.forEach((b, i) => { b.end = i + 1 < blocks.length ? blocks[i + 1].wr - 2 : R; });
    // 每一天使用哪個區塊：有日期的週區塊優先（後面的異動覆蓋前面）；沒有則用當時的固定班表（無日期區塊）
    const chosen = {};
    for (let d = opt.start; d <= opt.end; d++) {
      let pick = -1;
      blocks.forEach((b, i) => { if (b.dated && Object.values(b.dates).indexOf(d) >= 0) pick = i; });
      if (pick < 0) {
        const nextDated = blocks.findIndex(b => b.dated && Math.min.apply(null, Object.values(b.dates)) > d);
        const limit = nextDated < 0 ? blocks.length : nextDated;
        for (let i = 0; i < limit; i++) if (!blocks[i].dated) pick = i;
        if (pick < 0) pick = blocks.findIndex(b => !b.dated);
      }
      if (pick >= 0) chosen[d] = pick;
    }
    const hours = {}, detail = [], seen = new Set();
    blocks.forEach((b, bi) => {
      const dayFor = {};  // 星期 k → 天數序號陣列
      Object.keys(chosen).forEach(d => {
        if (chosen[d] !== bi) return;
        d = +d; const k = (new Date(d * 86400000).getUTCDay() + 6) % 7;
        (dayFor[k] = dayFor[k] || []).push(d);
      });
      if (!Object.keys(dayFor).length) return;
      let section = '';
      for (let r = b.wr + 1; r < b.end; r++) {
        const a = g[r][0].trim();
        if (a && (r === b.wr + 1 || g[r - 1][0].trim() !== a)) section = a;
        if (exclude.some(x => section.indexOf(x) >= 0)) continue;
        for (let c = 1; c < 15; c++) {
          const name = g[r][c].trim();
          if (!looksName(name)) continue;
          const k = Math.floor((c - 1) / 2);
          const ds = dayFor[k]; if (!ds) continue;
          let st = null, en = null;
          const rg = r + 1 < R ? timeRange(g[r + 1][c]) : null;
          if (rg) { st = rg[0]; en = rg[1]; }
          else if (r + 3 < R && timeMin(g[r + 1][c]) !== null && timeMin(g[r + 3][c]) !== null) { st = timeMin(g[r + 1][c]); en = timeMin(g[r + 3][c]); }
          if (st === null || en <= st) continue;
          ds.forEach(d => {
            const key = d + '|' + name + '|' + st + '|' + en;
            if (seen.has(key)) return; seen.add(key);
            hours[name] = (hours[name] || 0) + (en - st) / 60;
            detail.push({ day: d, name, start: st, end: en, section });
          });
        }
      }
    });
    Object.keys(hours).forEach(n => hours[n] = r2(hours[n]));
    return { hours, detail };
  }

  /* ---------------- 排名、增減、平均 ---------------- */
  const METRICS = [
    { k: 'hours',    group: '上診時數',   label: '上診時數', fmt: 'h',   color: '#6b7280' },
    { k: 'visits',   group: '看診人次',   label: '看診人次', fmt: 'int', color: '#6b7280' },
    { k: 'slowN',    group: '醫師慢專比例', label: '慢專人次', fmt: 'int', color: '#d98b0b' },
    { k: 'slowRate', group: '醫師慢專比例', label: '慢專比例', fmt: 'pct', color: '#d98b0b', avg: true },
    { k: 'vph',      group: '平均人次/hr', label: '平均人次/hr', fmt: 'n2', color: '#2f6fb3', avg: true },
    { k: 'backN',    group: 'BACK開立',   label: 'BACK開立數', fmt: 'int', color: '#7c5cc4' },
    { k: 'backRate', group: 'BACK開立',   label: 'BACK開立率', fmt: 'pct', color: '#7c5cc4', avg: true },
    { k: 'nRevN',    group: '狹義回診率', label: '狹義回診數', fmt: 'int', color: '#e07a2e' },
    { k: 'nRevRate', group: '狹義回診率', label: '狹義回診率', fmt: 'pct', color: '#e07a2e', avg: true, gated: true },
    { k: 'wRevN',    group: '廣義回診率', label: '廣義回診數', fmt: 'int', color: '#3f9b5a' },
    { k: 'wRevRate', group: '廣義回診率', label: '廣義回診率', fmt: 'pct', color: '#3f9b5a', avg: true }
  ];
  const METRIC = {}; METRICS.forEach(m => METRIC[m.k] = m);

  /** 依集團全院區排名（數值大者名次前；同值同名次）。狹義回診率需 BACK 開立率 ≥ 門檻 */
  function rankAll(records, threshold) {
    const out = {};
    METRICS.forEach(m => {
      const vals = [];
      records.forEach(r => {
        const v = numOrNull(r[m.k]);
        if (v === null) return;
        if (m.gated) { const br = numOrNull(r.backRate); if (br === null || br < threshold) return; }
        vals.push(v);
      });
      const n = vals.length;
      records.forEach(r => {
        const key = r.clinic + '|' + r.doctor;
        out[key] = out[key] || {};
        const v = numOrNull(r[m.k]);
        if (v === null) { out[key][m.k] = null; return; }
        if (m.gated) { const br = numOrNull(r.backRate); if (br === null || br < threshold) { out[key][m.k] = { none: true }; return; } }
        out[key][m.k] = { rank: 1 + vals.filter(x => x > v + 1e-9).length, n };
      });
    });
    return out;
  }
  function averages(records) {
    const out = {};
    METRICS.forEach(m => {
      const vals = records.map(r => numOrNull(r[m.k])).filter(v => v !== null);
      out[m.k] = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
    });
    return out;
  }
  function diffOf(cur, prev, k) {
    if (!prev) return null;
    const a = numOrNull(cur[k]), b = numOrNull(prev[k]);
    if (a === null || b === null) return null;
    return a - b;
  }

  /** 區間彙總（多個月加總後重算比例） */
  function aggregate(recs) {
    const sum = k => { const v = recs.map(r => numOrNull(r[k])).filter(x => x !== null); return v.length ? v.reduce((a, b) => a + b, 0) : null; };
    // 比例只用「分子、分母都有值」的月份計算，避免缺資料的月份拉低比例
    const gv = (r, k) => numOrNull(k === 'patients' && isBlank(r.patients) ? r.visits : r[k]);
    const pair = (n, d) => { let a = 0, b = 0, ok = false; recs.forEach(r => { const x = gv(r, n), y = gv(r, d); if (x !== null && y) { a += x; b += y; ok = true; } }); return ok ? [a, b] : null; };
    const o = { hours: sum('hours'), patients: sum('patients'), visits: sum('visits'), slowN: sum('slowN'), backN: sum('backN'), nRevN: sum('nRevN'), wRevN: sum('wRevN') };
    if (o.patients === null) o.patients = o.visits;
    const q = (p, f) => p ? f(p[0], p[1]) : null;
    o.slowRate = q(pair('slowN', 'visits'), (a, b) => r4(a / b));
    o.vph = q(pair('patients', 'hours'), (a, b) => r2(a / b));
    o.backRate = q(pair('backN', 'visits'), (a, b) => r4(a / b));
    o.nRevRate = q(pair('nRevN', 'backN'), (a, b) => r4(a / b));
    o.wRevRate = q(pair('wRevN', 'visits'), (a, b) => r4(a / b));
    return o;
  }

  /* ---------------- 格式化 ---------------- */
  function fmt(v, f, pctDigits) {
    if (v === null || v === undefined || v === '' || (typeof v === 'number' && isNaN(v))) return '';
    const n = Number(v);
    const pd = pctDigits == null ? 2 : pctDigits;
    switch (f) {
      case 'pct': return (n * 100).toFixed(pd) + '%';
      case 'n2': return n.toFixed(2);
      case 'h': return String(Math.round(n * 100) / 100);
      case 'int': return String(Math.round(n));
      default: return String(v);
    }
  }
  function fmtDiff(v, f, pctDigits) {
    if (v === null || v === undefined || isNaN(v)) return '';
    const pd = pctDigits == null ? 2 : pctDigits;
    if (f === 'pct') return (v * 100).toFixed(pd) + '%';
    if (f === 'n2') return (Math.round(v * 100) / 100).toFixed(2);
    if (f === 'h') return (Math.round(v * 10) / 10).toFixed(1);
    return String(Math.round(v));
  }

  /* ---------------- 報表（結果表）HTML ---------------- */
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /**
   * data: {month, analysisDate, records, prev, clinicSum, clinicOrder:[院區], doctorOrder:{院區|醫師:序}, titles:{院區|醫師:職稱}}
   * cfg:  {clinics:[..] 空=全部, doctors:[院區|醫師] 空=全部, cols:{k:{v,d,r}}, showAvg, showClinic:{total,slow,rate,rank}, showGrand, markNew,
   *        threshold, topColor, showLimit, pctDigits}
   */
  function renderReport(data, cfg) {
    const pd = cfg.pctDigits == null ? 2 : +cfg.pctDigits;
    const all = data.records || [];
    const ranks = rankAll(all, cfg.threshold == null ? 0.3 : +cfg.threshold);
    const avg = averages(all);
    const prevMap = {}; (data.prev || []).forEach(r => { prevMap[r.clinic + '|' + r.doctor] = r; prevMap['*' + r.doctor] = prevMap['*' + r.doctor] || r; });
    const csMap = {}; (data.clinicSum || []).forEach(c => csMap[c.clinic] = c);
    // 院區排名（院區慢箋比例）
    const csList = (data.clinicSum || []).filter(c => numOrNull(c.slowRate) !== null);
    const csRank = {}; csList.forEach(c => { csRank[c.clinic] = 1 + csList.filter(x => +x.slowRate > +c.slowRate + 1e-9).length; });
    const csAvg = csList.length ? csList.reduce((a, c) => a + (+c.slowRate), 0) / csList.length : null;
    const grand = (data.clinicSum || []).reduce((a, c) => a + (numOrNull(c.total) || 0), 0);

    const order = data.clinicOrder || [];
    const clinicsAll = Array.from(new Set(order.concat(all.map(r => r.clinic)))).filter(c => all.some(r => r.clinic === c));
    const clinics = clinicsAll.filter(c => !cfg.clinics || !cfg.clinics.length || cfg.clinics.indexOf(c) >= 0);
    const docOk = r => !cfg.doctors || !cfg.doctors.length || cfg.doctors.indexOf(r.clinic + '|' + r.doctor) >= 0;
    const dOrd = data.doctorOrder || {};

    // 欄位
    const cols = [];
    METRICS.forEach(m => {
      const c = (cfg.cols || {})[m.k] || {};
      ['v', 'd', 'r'].forEach(s => { if (c[s]) cols.push({ m, s }); });
    });
    const cs = cfg.showClinic || {};
    const csCols = [['total', '院區總人次'], ['slow', '院區慢箋人次'], ['rate', '院區慢箋比例'], ['rank', '院區排名']].filter(x => cs[x[0]]);

    // 表頭：群組 / 項目 / 值增減排名
    const groups = [];
    cols.forEach(c => { const last = groups[groups.length - 1]; if (last && last.g === c.m.group) last.cols.push(c); else groups.push({ g: c.m.group, cols: [c] }); });
    const title = monthLabel(data.month) + (data.analysisDate ? '<br><small>(' + esc(data.analysisDate) + '分析)</small>' : '');
    let h = '<table class="rpt"><thead>';
    h += '<tr><th class="t-title" colspan="2" rowspan="3">' + title + '</th>';
    groups.forEach(gp => h += '<th colspan="' + gp.cols.length + '" class="g">' + esc(gp.g) + '</th>');
    if (csCols.length) h += '<th class="gap" rowspan="' + (3 + 0) + '"></th>' + csCols.map(x => '<th rowspan="3" class="cs-h">' + x[1].replace(/(院區)/, '$1<br>') + '</th>').join('');
    h += '</tr><tr>';
    groups.forEach(gp => {
      const ms = []; gp.cols.forEach(c => { const l = ms[ms.length - 1]; if (l && l.m === c.m) l.n++; else ms.push({ m: c.m, n: 1 }); });
      ms.forEach(x => h += '<th colspan="' + x.n + '" class="m">' + (x.m.label === gp.g ? '' : esc(x.m.label)) + (x.m.gated ? '<div class="note">開立率未達' + Math.round((cfg.threshold == null ? 0.3 : cfg.threshold) * 100) + '%不列排名</div>' : '') + '</th>');
    });
    h += '</tr><tr>';
    cols.forEach(c => h += '<th class="s">' + ({ v: '值', d: '增減', r: '排名' }[c.s]) + '</th>');
    h += '</tr></thead><tbody>';

    // 集團平均值列
    if (cfg.showAvg !== false) {
      h += '<tr class="avg"><td colspan="2">集團平均值</td>';
      let i = 0;
      while (i < cols.length) {
        // 平均值以「值」欄位為中心，合併該項目的所有子欄
        const m = cols[i].m; let n = 0; while (i + n < cols.length && cols[i + n].m === m) n++;
        const hasV = cols.slice(i, i + n).some(c => c.s === 'v');
        h += '<td colspan="' + n + '">' + (m.avg && hasV ? fmt(avg[m.k], m.fmt, pd) : '') + '</td>';
        i += n;
      }
      if (csCols.length) {
        h += '<td class="gap"></td>';
        const k = csCols.map(x => x[0]);
        const labelSpan = k.filter(x => x !== 'rate' && x !== 'rank' ).length;
        csCols.forEach(x => {
          if (x[0] === 'rate') h += '<td class="cs-avg">' + fmt(csAvg, 'pct', pd) + '</td>';
          else if (x[0] === 'total') h += '<td class="cs-avg">集團平均</td>';
          else h += '<td class="cs-avg"></td>';
        });
      }
      h += '</tr>';
    }

    const topC = +cfg.topColor || 5, lim = +cfg.showLimit || 10;
    clinics.forEach(cl => {
      const rows = all.filter(r => r.clinic === cl && docOk(r))
        .sort((a, b) => (dOrd[a.clinic + '|' + a.doctor] || 999) - (dOrd[b.clinic + '|' + b.doctor] || 999));
      if (!rows.length) return;
      rows.forEach((r, ri) => {
        const key = r.clinic + '|' + r.doctor;
        const prev = prevMap[key] || prevMap['*' + r.doctor];
        const isNew = false;
        h += '<tr' + (ri === 0 ? ' class="first"' : '') + '>';
        if (ri === 0) h += '<td class="clinic" rowspan="' + rows.length + '"><span>' + esc(cl) + '</span></td>';
        h += '<td class="doc' + (isNew ? ' new' : '') + '">' + esc(r.doctor) + esc(r.title || (data.titles || {})[key] || '醫師') + '</td>';
        cols.forEach(c => {
          const m = c.m;
          if (c.s === 'v') h += '<td class="v">' + fmt(r[m.k], m.fmt, pd) + '</td>';
          else if (c.s === 'd') { const dv = diffOf(r, prev, m.k); h += '<td class="d' + (dv < 0 ? ' neg' : '') + '">' + fmtDiff(dv, m.fmt, pd) + '</td>'; }
          else {
            const rk = (ranks[key] || {})[m.k];
            if (!rk) h += '<td class="r"></td>';
            else if (rk.none) h += '<td class="r none">無</td>';
            else if (rk.rank > lim) h += '<td class="r"></td>';
            else h += '<td class="r' + (rk.rank <= topC ? ' top' : '') + '"' + (rk.rank <= topC ? ' style="--c:' + m.color + '"' : '') + '><b>' + rk.rank + '</b><sub>/' + rk.n + '</sub></td>';
          }
        });
        if (csCols.length && ri === 0) {
          const c = csMap[cl] || {};
          h += '<td class="gap" rowspan="' + rows.length + '"></td>';
          csCols.forEach(x => {
            let v = '';
            if (x[0] === 'total') v = fmt(c.total, 'int');
            if (x[0] === 'slow') v = fmt(c.slow, 'int');
            if (x[0] === 'rate') v = fmt(c.slowRate, 'pct', pd);
            if (x[0] === 'rank') v = csRank[cl] || '';
            h += '<td class="cs" rowspan="' + rows.length + '">' + v + '</td>';
          });
        }
        h += '</tr>';
      });
    });
    h += '</tbody>';
    if (cfg.showGrand && csCols.some(x => x[0] === 'total')) {
      h += '<tfoot><tr><td colspan="' + (2 + cols.length) + '" class="nb"></td><td class="gap"></td>';
      csCols.forEach(x => h += '<td class="grand">' + (x[0] === 'total' ? grand.toLocaleString() : '') + '</td>');
      h += '</tr><tr><td colspan="' + (2 + cols.length) + '" class="nb"></td><td class="gap"></td><td class="grand-l" colspan="' + csCols.length + '">全院區總人次</td></tr></tfoot>';
    } else if (cfg.showGrand) {
      h += '<tfoot><tr><td colspan="' + (2 + cols.length) + '" class="grand-row">全院區總人次：<b>' + grand.toLocaleString() + '</b></td></tr></tfoot>';
    }
    h += '</table>';
    return h;
  }

  const REPORT_CSS = `
  .rpt{border-collapse:collapse;font-size:11px;font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif;color:#111;background:#fff}
  .rpt th,.rpt td{border:1px solid #333;padding:3px 5px;text-align:center;white-space:nowrap;height:22px}
  .rpt thead th{background:#fff;font-weight:700}
  .rpt th.g{font-size:12px}
  .rpt th.s{font-weight:500;font-size:10.5px}
  .rpt th.m .note{color:#d11;font-size:9px;font-weight:500}
  .rpt th.t-title{min-width:120px;font-size:12px}
  .rpt th.cs-h{font-size:10.5px;background:#fff;min-width:44px}
  .rpt td.d{background:#ececec}
  .rpt td.d.neg{color:#b42318}
  .rpt tr.avg td{font-weight:700;height:26px}
  .rpt td.clinic{writing-mode:vertical-rl;letter-spacing:4px;font-weight:600;width:22px;padding:4px 2px}
  .rpt td.doc{min-width:80px}
  .rpt td.doc.new{background:#fff200}
  .rpt td.r{min-width:36px;position:relative}
  .rpt td.r b{font-weight:500}
  .rpt td.r sub{font-size:8.5px;color:#444;margin-left:1px}
  .rpt td.r.top{background:color-mix(in srgb,var(--c) 18%,#fff)}
  .rpt td.r.top b{color:var(--c);font-weight:700}
  .rpt td.r.top sub{color:var(--c)}
  .rpt td.gap,.rpt th.gap{border:none;background:transparent;width:10px;min-width:10px;padding:0}
  .rpt td.cs{font-size:11px}
  .rpt td.cs-avg{background:#fdf1d8;font-weight:700}
  .rpt td.nb{border:none}
  .rpt td.grand{border:none;color:#c0271d;font-weight:700;font-size:12px}
  .rpt td.grand-l{border:none;color:#c0271d;font-size:10px}
  .rpt td.grand-row{text-align:right;border:none;color:#c0271d;padding-top:8px}
  `;

  /* ---------------- 美化圖片版報表（給股東／老闆，下載 PNG / JPEG） ---------------- */
  const hexA = (hex, a) => { const h = hex.replace('#', ''); const n = parseInt(h.length === 3 ? h.split('').map(c => c + c).join('') : h, 16); return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`; };

  function prettyPrep(data, cfg) {
    const all = data.records || [];
    const ranks = rankAll(all, cfg.threshold == null ? 0.3 : +cfg.threshold);
    const avg = averages(all);
    const prevMap = {}; (data.prev || []).forEach(r => { prevMap[r.clinic + '|' + r.doctor] = r; prevMap['*' + r.doctor] = prevMap['*' + r.doctor] || r; });
    const csMap = {}; (data.clinicSum || []).forEach(c => csMap[c.clinic] = c);
    const pcsMap = {}; (data.prevClinicSum || []).forEach(c => pcsMap[c.clinic] = c);
    const csList = (data.clinicSum || []).filter(c => numOrNull(c.slowRate) !== null);
    const csRank = {}; csList.forEach(c => { csRank[c.clinic] = 1 + csList.filter(x => +x.slowRate > +c.slowRate + 1e-9).length; });
    const csAvg = csList.length ? csList.reduce((a, c) => a + (+c.slowRate), 0) / csList.length : null;
    const grand = (data.clinicSum || []).reduce((a, c) => a + (numOrNull(c.total) || 0), 0);
    const pgrand = (data.prevClinicSum || []).reduce((a, c) => a + (numOrNull(c.total) || 0), 0);
    const order = data.clinicOrder || [];
    const clinics = Array.from(new Set(order.concat(all.map(r => r.clinic)))).filter(c => all.some(r => r.clinic === c))
      .filter(c => !cfg.clinics || !cfg.clinics.length || cfg.clinics.indexOf(c) >= 0);
    const docOk = r => !cfg.doctors || !cfg.doctors.length || cfg.doctors.indexOf(r.clinic + '|' + r.doctor) >= 0;
    const dOrd = data.doctorOrder || {};
    const rowsOf = cl => all.filter(r => r.clinic === cl && docOk(r)).sort((a, b) => (dOrd[a.clinic + '|' + a.doctor] || 999) - (dOrd[b.clinic + '|' + b.doctor] || 999));
    const cols = [];
    METRICS.forEach(m => { const c = (cfg.cols || {})[m.k] || {}; if (c.v || c.d || c.r) cols.push({ m, v: !!c.v, d: !!c.d, r: !!c.r }); });
    const cs = cfg.showClinic || {};
    return { all, ranks, avg, prevMap, csMap, pcsMap, csRank, csAvg, grand, pgrand, clinics, rowsOf, cols, cs };
  }

  function prettyDiff(v, f, pd) {
    if (v === null || v === undefined || isNaN(v)) return '';
    const t = fmtDiff(Math.abs(v), f, pd);
    if (Math.abs(v) < 1e-9) return `<span class="pd z">－ ${t}</span>`;
    return `<span class="pd ${v > 0 ? 'up' : 'dn'}">${v > 0 ? '▲' : '▼'} ${t}</span>`;
  }
  function prettyRank(rk, m, topC, lim) {
    if (!rk) return '';
    if (rk.none) return '';
    if (rk.rank > lim) return '';
    const top = rk.rank <= topC;
    return `<span class="pr${top ? ' top' : ''}" style="${top ? `background:${m.color};border-color:${m.color};color:#fff` : `border-color:${hexA(m.color, .55)};color:${m.color}`}">${rk.rank}<i>/${rk.n}</i></span>`;
  }

  /**
   * variant 'table'：整張表格美化版；'cards'：院區卡片版
   * opts: {title, subtitle, clinicName}
   */
  function renderPretty(data, cfg, variant) {
    const pd = cfg.pctDigits == null ? 2 : +cfg.pctDigits;
    const P = prettyPrep(data, cfg);
    const topC = +cfg.topColor || 5, lim = +cfg.showLimit || 10;
    const month = monthLabel(data.month);
    const kpis = [];
    if (P.grand) kpis.push(['全院區總人次', P.grand.toLocaleString(), (P.pgrand ? prettyDiff(P.grand - P.pgrand, 'int') + '<em>較上月</em>' : '') + (P.cs.rate && P.csAvg !== null ? `<em>院區慢箋平均 ${fmt(P.csAvg, 'pct', pd)}</em>` : '')]);
    kpis.push(['列入醫師', String(P.clinics.reduce((a, c) => a + P.rowsOf(c).length, 0)), `<em>${P.clinics.length} 個院區</em>`]);
    [['slowRate', '集團平均 慢專比例'], ['vph', '集團平均 人次/hr'], ['nRevRate', '集團平均 狹義回診率'], ['wRevRate', '集團平均 廣義回診率']].forEach(([k, l]) => {
      if (P.avg[k] !== null) kpis.push([l, fmt(P.avg[k], METRIC[k].fmt, pd), '']);
    });
    let h = `<div class="pretty pv-${variant}">
      <div class="p-head">
        <img src="${LOGO}" class="p-logo" alt="">
        <div class="p-t"><div class="p-org">金鶯診所 Elite Clinic</div><div class="p-title">${esc(month)} 醫師回診率分析結果</div></div>
        <div class="p-meta">${data.analysisDate ? esc(data.analysisDate) + ' 分析' : ''}</div>
      </div>
      <div class="p-kpis" style="grid-template-columns:repeat(${kpis.length},1fr)">${kpis.map(k => `<div class="p-kpi"><div class="l">${k[0]}</div><div class="v">${k[1]}</div><div class="s">${k[2]}</div></div>`).join('')}</div>`;

    if (variant === 'cards') h += prettyCards(P, cfg, pd, topC, lim);
    else h += prettyTable(P, cfg, pd, topC, lim, data);

    h += `<div class="p-foot"><span>排名：集團全院區醫師排名，前 ${topC} 名實心標示、${lim} 名以後不顯示；狹義回診率僅排 BACK 開立率 ≥ ${Math.round((cfg.threshold == null ? .3 : cfg.threshold) * 100)}% 之醫師。▲▼ 為與上月相比。</span><span>金鶯診所˙醫師回診率分析系統</span></div></div>`;
    return h;
  }

  function prettyTable(P, cfg, pd, topC, lim, data) {
    const csCols = [['total', '院區總人次'], ['slow', '慢箋人次'], ['rate', '慢箋比例'], ['rank', '院區排名']].filter(x => P.cs[x[0]]);
    const groups = []; P.cols.forEach(c => { const l = groups[groups.length - 1]; if (l && l.g === c.m.group) l.cols.push(c); else groups.push({ g: c.m.group, cols: [c], color: c.m.color }); });
    let h = '<div class="p-card"><table class="p-tb"><thead><tr><th class="c-cl" rowspan="2">院區</th><th class="c-doc" rowspan="2">醫師</th>';
    groups.forEach(g => h += `<th colspan="${g.cols.length}" class="g" style="border-bottom:3px solid ${g.color}">${esc(g.g)}</th>`);
    if (csCols.length) h += `<th colspan="${csCols.length}" class="g cs" style="border-bottom:3px solid #1e3d3a">院區統計</th>`;
    h += '</tr><tr>';
    P.cols.forEach(c => h += `<th class="m">${c.m.label === c.m.group || groups.find(g => g.g === c.m.group).cols.length === 1 ? '數值' : esc(c.m.label.replace('BACK', ''))}</th>`);
    csCols.forEach(x => h += `<th class="m cs">${x[1]}</th>`);
    h += '</tr></thead><tbody>';
    P.clinics.forEach((cl, ci) => {
      const rows = P.rowsOf(cl); if (!rows.length) return;
      rows.forEach((r, ri) => {
        const key = r.clinic + '|' + r.doctor;
        const prev = P.prevMap[key] || P.prevMap['*' + r.doctor];
        const isNew = false;
        h += `<tr class="${ci % 2 ? 'band' : ''}${ri === 0 ? ' first' : ''}">`;
        if (ri === 0) h += `<td class="c-cl" rowspan="${rows.length}">${esc(cl).split('').join('<br>')}</td>`;
        h += `<td class="c-doc"><b>${esc(r.doctor)}</b><small>${esc(r.title || '醫師')}</small>${isNew ? '<em class="new">新進</em>' : ''}</td>`;
        P.cols.forEach(c => {
          const m = c.m; const rk = (P.ranks[key] || {})[m.k];
          let inner = '';
          if (c.v) inner += `<div class="pv">${fmt(r[m.k], m.fmt, pd) || '<span class="na">—</span>'}</div>`;
          const sub = (c.d ? prettyDiff(diffOf(r, prev, m.k), m.fmt, pd) : '') + (c.r ? prettyRank(rk, m, topC, lim) : '');
          if (sub) inner += `<div class="ps">${sub}</div>`;
          h += `<td>${inner}</td>`;
        });
        if (csCols.length && ri === 0) {
          const c = P.csMap[cl] || {};
          csCols.forEach(x => {
            let v = '';
            if (x[0] === 'total') v = `<div class="pv big">${fmt(c.total, 'int') ? (+c.total).toLocaleString() : ''}</div>` + (P.pcsMap[cl] && c.total ? `<div class="ps">${prettyDiff(+c.total - (+P.pcsMap[cl].total || 0), 'int')}</div>` : '');
            if (x[0] === 'slow') v = `<div class="pv">${fmt(c.slow, 'int')}</div>`;
            if (x[0] === 'rate') v = `<div class="pv">${fmt(c.slowRate, 'pct', pd)}</div>`;
            if (x[0] === 'rank') v = P.csRank[cl] ? `<span class="crk${P.csRank[cl] <= 3 ? ' top' : ''}">${P.csRank[cl]}</span>` : '';
            h += `<td class="cs" rowspan="${rows.length}">${v}</td>`;
          });
        }
        h += '</tr>';
      });
    });
    h += '</tbody></table></div>';
    return h;
  }

  function prettyCards(P, cfg, pd, topC, lim) {
    const cols = P.cols.filter(c => c.v || c.r);
    const per = cols.length <= 5 ? 3 : 2;
    let h = '';
    // 院區總覽長條
    const cl = P.clinics.map(c => ({ c, t: numOrNull((P.csMap[c] || {}).total), r: numOrNull((P.csMap[c] || {}).slowRate) }));
    if (cl.some(x => x.t)) {
      const maxT = Math.max.apply(null, cl.map(x => x.t || 0));
      const sorted = cl.slice().sort((a, b) => (b.t || 0) - (a.t || 0));
      const maxR = Math.max.apply(null, cl.map(x => x.r || 0)) || 1;
      const sortedR = cl.slice().sort((a, b) => (b.r || 0) - (a.r || 0));
      h += `<div class="p-bars"><div class="p-card"><div class="p-h">院區總人次</div>${sorted.map(x => `<div class="bar"><span class="bl">${esc(x.c)}</span><span class="bt"><i style="width:${x.t ? (x.t / maxT * 100).toFixed(1) : 0}%"></i></span><span class="bv">${x.t ? x.t.toLocaleString() : '—'}</span></div>`).join('')}</div>
        <div class="p-card"><div class="p-h">院區慢箋比例 <small>集團平均 ${fmt(P.csAvg, 'pct', pd)}</small></div>${sortedR.map(x => `<div class="bar"><span class="bl">${esc(x.c)}</span><span class="bt"><i class="r" style="width:${x.r ? (x.r / maxR * 100).toFixed(1) : 0}%"></i></span><span class="bv">${fmt(x.r, 'pct', pd) || '—'}</span></div>`).join('')}</div></div>`;
    }
    h += `<div class="p-grid" style="grid-template-columns:repeat(${per},1fr)">`;
    P.clinics.forEach(c => {
      const rows = P.rowsOf(c); if (!rows.length) return;
      const s = P.csMap[c] || {};
      h += `<div class="p-card cc"><div class="cc-h"><div class="cc-n">${esc(c)}</div><div class="cc-s">`;
      if (P.cs.total && s.total) h += `<span>總人次 <b>${(+s.total).toLocaleString()}</b></span>`;
      if (P.cs.rate && s.slowRate !== '' && s.slowRate != null) h += `<span>慢箋 <b>${fmt(s.slowRate, 'pct', pd)}</b></span>`;
      if (P.cs.rank && P.csRank[c]) h += `<span class="crk${P.csRank[c] <= 3 ? ' top' : ''}">${P.csRank[c]}</span>`;
      h += `</div></div><table class="cc-t"><thead><tr><th class="l">醫師</th>${cols.map(x => `<th style="color:${x.m.color}">${esc(x.m.label)}</th>`).join('')}</tr></thead><tbody>`;
      rows.forEach(r => {
        const key = r.clinic + '|' + r.doctor; const prev = P.prevMap[key] || P.prevMap['*' + r.doctor];
        h += `<tr><td class="l"><b>${esc(r.doctor)}</b><small>${esc(r.title || '')}</small></td>` + cols.map(x => {
          const rk = (P.ranks[key] || {})[x.m.k];
          return `<td>${x.v ? `<div class="pv">${fmt(r[x.m.k], x.m.fmt, pd) || '<span class="na">—</span>'}</div>` : ''}<div class="ps">${x.d ? prettyDiff(diffOf(r, prev, x.m.k), x.m.fmt, pd) : ''}${x.r ? prettyRank(rk, x.m, topC, lim) : ''}</div></td>`;
        }).join('') + '</tr>';
      });
      h += '</tbody></table></div>';
    });
    h += '</div>';
    return h;
  }

  const PRETTY_CSS = `
  .pretty{width:max-content;min-width:1680px;background:#f2f0eb;padding:36px 40px 28px;font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif;color:#1d2a28;box-sizing:border-box}
  .pretty *{box-sizing:border-box}
  .pretty .p-head{display:flex;align-items:center;gap:22px;background:#1e3d3a;color:#fff;border-radius:22px;padding:22px 30px}
  .pretty .p-logo{width:78px;height:78px;border-radius:50%;background:#fff;padding:3px}
  .pretty .p-org{font-size:17px;color:#9fd4c7;letter-spacing:2px}
  .pretty .p-title{font-size:34px;font-weight:700;letter-spacing:1px;margin-top:2px}
  .pretty .p-meta{margin-left:auto;text-align:right;font-size:18px;font-weight:500;line-height:1.6}
  .pretty .p-meta span{font-size:14px;color:#9fd4c7;font-weight:400}
  .pretty .p-kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin:18px 0}
  .pretty .p-kpi{background:#fff;border:1px solid #e3dfd5;border-radius:16px;padding:14px 18px}
  .pretty .p-kpi .l{font-size:17px;color:#7b7a72}
  .pretty .p-kpi .v{font-size:38px;font-weight:700;margin-top:2px;color:#1e3d3a}
  .pretty .p-kpi .s{font-size:15px;min-height:18px}
  .pretty .p-kpi em{font-style:normal;color:#7b7a72;margin-left:6px}
  .pretty .p-card{background:#fff;border:1px solid #e3dfd5;border-radius:18px;padding:10px 12px;overflow:hidden}
  .pretty .p-tb{width:100%;border-collapse:collapse;font-size:17px}
  .pretty .p-tb th{font-weight:700;color:#35504c;padding:12px 10px;text-align:center;white-space:nowrap}
  .pretty .p-tb th.g{font-size:19px}
  .pretty .p-tb th.m{font-size:15px;color:#7b7a72;font-weight:500;background:#f6f4ef;border-bottom:1px solid #e3dfd5}
  .pretty .p-tb td{padding:10px 10px;text-align:center;border-bottom:1px solid #efece5;white-space:nowrap;vertical-align:middle}
  .pretty .p-tb tr.band td{background:#fbfaf7}
  .pretty .p-tb tr.first td{border-top:2px solid #d8e6e2}
  .pretty .p-tb tr.avg td{background:#fdf1d8;font-weight:700;color:#7a5412;font-size:15px}
  .pretty .p-tb td.c-cl{background:#e6efed!important;font-weight:700;color:#1e3d3a;font-size:21px;width:56px;letter-spacing:2px}
  .pretty .p-tb td.c-cl span{writing-mode:vertical-rl}
  .pretty .p-tb td.c-doc{text-align:left;padding-left:14px;min-width:150px;font-size:19px}
  .pretty .p-tb td.c-doc small{color:#7b7a72;font-size:14px;margin-left:3px}
  .pretty em.new{font-style:normal;background:#fde68a;color:#7a5412;font-size:11px;border-radius:6px;padding:1px 6px;margin-left:5px}
  .pretty .pv{font-size:20px;font-weight:500}
  .pretty .pv.big{font-size:23px;font-weight:700;color:#1e3d3a}
  .pretty .ps{display:flex;gap:4px;justify-content:center;align-items:center;margin-top:2px;min-height:0}
  .pretty .pd{font-size:14.5px;font-weight:500}
  .pretty .pd.up{color:#2b7a5f}.pretty .pd.dn{color:#c0392b}.pretty .pd.z{color:#9a988f}
  .pretty .pr{display:inline-block;font-size:14.5px;font-weight:700;border:1.5px solid;border-radius:99px;padding:1px 9px;line-height:20px}
  .pretty .pr i{font-style:normal;font-weight:400;font-size:12px;opacity:.85}
  .pretty .pr.none{border-color:#d9d4c8;color:#9a988f;font-weight:400}
  .pretty .na{color:#c9c5bb}
  .pretty td.cs,.pretty th.cs{background:#f3f8f6}
  .pretty .crk{display:inline-grid;place-items:center;width:38px;height:38px;border-radius:50%;border:2px solid #2b5a54;color:#2b5a54;font-weight:700;font-size:19px}
  .pretty .crk.top{background:#2b5a54;color:#fff}
  .pretty .p-foot{display:flex;justify-content:space-between;gap:20px;font-size:14px;color:#7b7a72;margin-top:14px}
  .pretty .p-bars{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
  .pretty .p-h{font-weight:700;font-size:21px;color:#1e3d3a;margin:6px 6px 10px}
  .pretty .p-h small{font-weight:400;font-size:13px;color:#7b7a72;margin-left:8px}
  .pretty .bar{display:flex;align-items:center;gap:10px;margin:6px 6px;font-size:17px}
  .pretty .bar .bl{width:54px;font-weight:500}
  .pretty .bar .bt{flex:1;height:20px;background:#f2f0eb;border-radius:8px;overflow:hidden}
  .pretty .bar .bt i{display:block;height:100%;background:#2b5a54;border-radius:8px}
  .pretty .bar .bt i.r{background:#d98b0b}
  .pretty .bar .bv{width:90px;text-align:right;font-weight:700}
  .pretty .p-grid{display:grid;gap:16px}
  .pretty .cc{padding:14px 16px}
  .pretty .cc-h{display:flex;align-items:center;justify-content:space-between;border-bottom:2px solid #e6efed;padding-bottom:8px;margin-bottom:6px}
  .pretty .cc-n{font-size:27px;font-weight:700;color:#1e3d3a;letter-spacing:2px}
  .pretty .cc-s{display:flex;gap:14px;align-items:center;font-size:16px;color:#7b7a72}
  .pretty .cc-s b{color:#1d2a28;font-size:20px}
  .pretty .cc-t{width:100%;border-collapse:collapse;font-size:17px}
  .pretty .cc-t th{font-size:15px;font-weight:700;padding:6px 4px;text-align:center;white-space:nowrap}
  .pretty .cc-t td{padding:10px 6px;text-align:center;border-top:1px solid #efece5;white-space:nowrap}
  .pretty .cc-t .l{text-align:left}
  .pretty .cc-t td.l small{color:#7b7a72;font-size:13.5px;margin-left:3px}
  `;

  // 清單一律由後台（登入後）載入，不在公開的程式碼裡放醫師名單
  const DEFAULT_CONFIG = { clinics: [], doctors: [], visitMap: [], backMap: [], settings: {} };
  const LOGO = 'data:image/png;base64,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';

  root.Engine = { DEFAULT_CONFIG, LOGO, renderPretty, PRETTY_CSS, normCode, num, numOrNull, rocDay, isoDay, dayIso, dayRoc, monthLabel, prevMonth, monthRange, r2, r4,
    parseVisitRows, parseBackRows, computeClinic, metrics, parseSchedule, METRICS, METRIC, rankAll, averages, diffOf, aggregate,
    fmt, fmtDiff, renderReport, REPORT_CSS, esc };
})(typeof window !== 'undefined' ? window : globalThis);
