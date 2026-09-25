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
        const isNew = cfg.markNew && (data.prev || []).length && !prev;
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

  // 後台尚未連線時使用的預設清單（正式資料以 Google Sheet 為準）
  const DEFAULT_CONFIG = {"clinics": [{"order": 1, "name": "鶯歌", "fileKey": "鶯歌", "scheduleTab": "鶯歌", "active": true}, {"order": 2, "name": "三樹", "fileKey": "三樹", "scheduleTab": "三樹", "active": true}, {"order": 3, "name": "桃園", "fileKey": "桃園", "scheduleTab": "桃園", "active": true}, {"order": 4, "name": "土城", "fileKey": "土城", "scheduleTab": "土城", "active": true}, {"order": 5, "name": "八德", "fileKey": "八德", "scheduleTab": "八德", "active": true}, {"order": 6, "name": "龍潭", "fileKey": "龍潭", "scheduleTab": "龍潭", "active": true}, {"order": 7, "name": "大溪", "fileKey": "大溪", "scheduleTab": "大溪", "active": true}, {"order": 8, "name": "大竹", "fileKey": "大竹", "scheduleTab": "大竹", "active": true}, {"order": 9, "name": "大湳", "fileKey": "大湳", "scheduleTab": "大湳", "active": true}, {"order": 10, "name": "平鎮", "fileKey": "平鎮", "scheduleTab": "平鎮", "active": true}, {"order": 11, "name": "埔心", "fileKey": "埔心", "scheduleTab": "埔心", "active": true}, {"order": 12, "name": "觀音", "fileKey": "觀音", "scheduleTab": "觀音", "active": true}, {"order": 13, "name": "樹林", "fileKey": "樹林", "scheduleTab": "樹林", "active": true}], "doctors": [{"order": 1, "clinic": "鶯歌", "name": "林永昌", "title": "院長", "active": true}, {"order": 2, "clinic": "鶯歌", "name": "陳珠", "title": "醫師", "active": true}, {"order": 3, "clinic": "鶯歌", "name": "鍾偉瑋", "title": "醫師", "active": true}, {"order": 4, "clinic": "鶯歌", "name": "李軍逸", "title": "醫師", "active": true}, {"order": 5, "clinic": "鶯歌", "name": "陳郁凡", "title": "醫師", "active": true}, {"order": 6, "clinic": "三樹", "name": "張慧馨", "title": "院長", "active": true}, {"order": 7, "clinic": "三樹", "name": "闕壯理", "title": "醫師", "active": true}, {"order": 8, "clinic": "三樹", "name": "黃允玫", "title": "醫師", "active": true}, {"order": 9, "clinic": "桃園", "name": "蕭宇伯", "title": "醫師", "active": true}, {"order": 10, "clinic": "桃園", "name": "李豪權", "title": "醫師", "active": true}, {"order": 11, "clinic": "桃園", "name": "曾㨗聖", "title": "醫師", "active": true}, {"order": 12, "clinic": "桃園", "name": "陳佳舜", "title": "醫師", "active": true}, {"order": 13, "clinic": "土城", "name": "張簡千郁", "title": "院長", "active": true}, {"order": 14, "clinic": "土城", "name": "陳麒中", "title": "醫師", "active": true}, {"order": 15, "clinic": "土城", "name": "黃濰", "title": "醫師", "active": true}, {"order": 16, "clinic": "土城", "name": "何金蔚", "title": "醫師", "active": true}, {"order": 17, "clinic": "八德", "name": "廖崑竹", "title": "院長", "active": true}, {"order": 18, "clinic": "八德", "name": "李彥輝", "title": "醫師", "active": true}, {"order": 19, "clinic": "八德", "name": "王志冉", "title": "醫師", "active": true}, {"order": 20, "clinic": "八德", "name": "楊大緯", "title": "醫師", "active": true}, {"order": 21, "clinic": "八德", "name": "曾韋綸", "title": "醫師", "active": true}, {"order": 22, "clinic": "龍潭", "name": "何冠達", "title": "院長", "active": true}, {"order": 23, "clinic": "龍潭", "name": "何中庸", "title": "醫師", "active": true}, {"order": 24, "clinic": "龍潭", "name": "華志倫", "title": "醫師", "active": true}, {"order": 25, "clinic": "大溪", "name": "林鼎盛", "title": "院長", "active": true}, {"order": 26, "clinic": "大溪", "name": "林昆澤", "title": "醫師", "active": true}, {"order": 27, "clinic": "大溪", "name": "呂瑩純", "title": "醫師", "active": true}, {"order": 28, "clinic": "大竹", "name": "陳秉峰", "title": "院長", "active": true}, {"order": 29, "clinic": "大竹", "name": "沈士強", "title": "醫師", "active": true}, {"order": 30, "clinic": "大竹", "name": "彭榆真", "title": "醫師", "active": true}, {"order": 31, "clinic": "大湳", "name": "黃炯達", "title": "院長", "active": true}, {"order": 32, "clinic": "大湳", "name": "洪岱熙", "title": "醫師", "active": true}, {"order": 33, "clinic": "大湳", "name": "高國峯", "title": "醫師", "active": true}, {"order": 34, "clinic": "平鎮", "name": "陳昭銘", "title": "院長", "active": true}, {"order": 35, "clinic": "平鎮", "name": "甘家銘", "title": "醫師", "active": true}, {"order": 36, "clinic": "埔心", "name": "梁培毅", "title": "院長", "active": true}, {"order": 37, "clinic": "埔心", "name": "徐翊庭", "title": "醫師", "active": true}, {"order": 38, "clinic": "觀音", "name": "李逸瑋", "title": "院長", "active": true}, {"order": 39, "clinic": "觀音", "name": "侯進坤", "title": "醫師", "active": true}, {"order": 40, "clinic": "樹林", "name": "林秉承", "title": "院長", "active": true}, {"order": 41, "clinic": "樹林", "name": "廖崇淵", "title": "醫師", "active": true}, {"order": 42, "clinic": "樹林", "name": "尤騰", "title": "醫師", "active": true}, {"order": 43, "clinic": "樹林", "name": "何中誠", "title": "醫師", "active": true}, {"order": 44, "clinic": "樹林", "name": "陳永仁", "title": "醫師", "active": true}], "visitMap": [{"clinic": "觀音", "code": "69", "doctor": "李逸瑋", "active": true}, {"clinic": "觀音", "code": "97", "doctor": "侯進坤", "active": true}], "backMap": [{"clinic": "觀音", "code": "169", "doctor": "李逸瑋", "active": true}, {"clinic": "觀音", "code": "197", "doctor": "侯進坤", "active": true}], "settings": {"回診判定天數": 5, "狹義回診排名門檻": 0.3, "排名標色名次": 5, "排名顯示名次": 10, "班表排除診別關鍵字": "減重", "BACK醫師別去首碼對應": "Y", "最近班表連結": "https://docs.google.com/spreadsheets/d/16dTALt4zDoD0jD9seMA6poNPTYvDbPQBmYt9rLwi-vU/edit", "百分比小數位數": 2}};

  root.Engine = { DEFAULT_CONFIG, normCode, num, numOrNull, rocDay, isoDay, dayIso, dayRoc, monthLabel, prevMonth, monthRange, r2, r4,
    parseVisitRows, parseBackRows, computeClinic, metrics, parseSchedule, METRICS, METRIC, rankAll, averages, diffOf, aggregate,
    fmt, fmtDiff, renderReport, REPORT_CSS, esc };
})(typeof window !== 'undefined' ? window : globalThis);
