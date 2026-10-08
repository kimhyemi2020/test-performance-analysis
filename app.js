'use strict';

/* 수행평가 CSV 분석 — 모든 처리는 브라우저 안에서만 이루어진다. 네트워크 요청 없음. */

const SAMPLE_CSV = `학번,이름,반,발표(20),보고서(30),실험(30),포트폴리오(20)
10101,김서준,1,18,27,26,19
10102,이하은,1,16,24,22,17
10103,박도윤,1,12,18,20,14
10104,최지우,1,19,29,28,20
10105,정민서,1,15,22,25,16
10106,강예준,1,10,15,14,11
10107,조수아,1,17,26,24,18
10108,윤시우,1,14,21,19,15
10109,장하린,1,18,25,27,17
10110,임주원,1,13,20,21,13
10211,한지호,2,16,23,23,16
10212,오서연,2,20,28,29,19
10213,서은우,2,11,17,16,12
10214,신채원,2,17,25,26,18
10215,권유준,2,15,24,20,15
10216,황다은,2,14,22,23,17
10217,안건우,2,8,12,15,9
10218,송지안,2,18,27,25,18
10219,전현우,2,16,21,24,14
10220,홍소율,2,19,26,28,19
10321,유선우,3,13,19,18,15
10322,고나은,3,17,24,27,16
10323,문태윤,3,15,23,22,17
10324,양서윤,3,12,20,17,13
10325,손준서,3,18,28,26,18
10326,배아린,3,16,25,21,16
10327,백승민,3,9,16,13,10
10328,허예린,3,17,22,25,19
10329,남도현,3,14,18,22,14
10330,심유나,3,16,26,24,17
`;

const ID_HEADER = /^(학번|번호|순번|출석번호|반|학년|학급|id|no\.?)$/i;
const NAME_HEADER = /^(이름|성명|학생명|name)$/i;
const CLASS_HEADER = /^(반|학급|class)$/i;
const TOTAL_HEADER = /^(총점|합계|총합|total|sum)/i;
const MAX_IN_HEADER = /\s*[(\[（]\s*(\d+(?:\.\d+)?)\s*점?\s*[)\]）]\s*$/;
const TOTAL_KEY = '__total';
const SVG_NS = 'http://www.w3.org/2000/svg';

const $ = (id) => document.getElementById(id);
const state = { data: null, target: null, classValue: '', sortKey: null, sortDir: 'desc', query: '' };

/* ---------- 파일 읽기 ---------- */

// 엑셀에서 저장한 CSV는 EUC-KR(CP949)인 경우가 많아 UTF-8 해석이 실패하면 EUC-KR로 다시 읽는다.
function decodeBuffer(buffer) {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buffer);
  } catch (e) {
    return new TextDecoder('euc-kr').decode(buffer);
  }
}

function detectDelimiter(text) {
  const firstLine = text.split(/\r?\n/, 1)[0];
  const counts = [',', ';', '\t'].map((d) => [d, firstLine.split(d).length]);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][0];
}

function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const delimiter = detectDelimiter(text);
  const rows = [];
  let row = [], field = '', inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false;
      } else field += ch;
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === delimiter) {
      row.push(field); field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); field = '';
      rows.push(row); row = [];
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows
    .map((r) => r.map((c) => c.trim()))
    .filter((r) => r.some((c) => c !== ''));
}

function toNumber(value) {
  const cleaned = value.replace(/점$/, '').replace(/,/g, '').trim();
  if (cleaned === '') return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

/* ---------- 데이터 구조화 ---------- */

function pickScale(declaredMax, observedMax) {
  if (declaredMax) return declaredMax;
  const candidates = [5, 10, 20, 30, 40, 50, 100];
  const fit = candidates.find((c) => observedMax <= c);
  return fit || Math.ceil(observedMax / 100) * 100;
}

function buildData(rawRows) {
  if (rawRows.length < 2) throw new Error('열 이름 한 줄과 학생 데이터가 한 줄 이상 필요합니다.');
  const headers = rawRows[0].map((h, i) => h || `열${i + 1}`);
  const body = rawRows.slice(1);

  const columns = headers.map((label, idx) => {
    const values = body.map((r) => r[idx] || '');
    const numbers = values.map(toNumber);
    const hasValue = numbers.some((n) => n !== null);
    const allNumeric = numbers.every((n) => n === null || !Number.isNaN(n));
    const isText = ID_HEADER.test(label) || NAME_HEADER.test(label) || !hasValue || !allNumeric;
    const m = label.match(MAX_IN_HEADER);
    return {
      key: String(idx), idx, label,
      name: m ? label.replace(MAX_IN_HEADER, '') : label,
      type: isText ? 'text' : 'score',
      declaredMax: m ? Number(m[1]) : null,
    };
  });

  const scoreCols = columns.filter((c) => c.type === 'score');
  if (!scoreCols.length) throw new Error('점수로 볼 수 있는 숫자 열을 찾지 못했습니다. 열 이름과 값을 확인해 주세요.');

  const givenTotal = scoreCols.find((c) => TOTAL_HEADER.test(c.name));
  const items = scoreCols.filter((c) => c !== givenTotal);

  const rows = body.map((cells) => {
    const scores = {};
    scoreCols.forEach((c) => { scores[c.key] = toNumber(cells[c.idx] || ''); });
    return { cells, scores };
  });

  scoreCols.forEach((c) => {
    const observed = rows.map((r) => r.scores[c.key]).filter((v) => v !== null);
    c.scale = pickScale(c.declaredMax, Math.max(...observed));
  });

  let totalCol = null;
  if (givenTotal) {
    totalCol = givenTotal;
    totalCol.isTotal = true;
    if (!totalCol.declaredMax && items.length && items.every((c) => c.declaredMax)) {
      totalCol.scale = items.reduce((s, c) => s + c.declaredMax, 0);
    }
  } else if (items.length >= 2) {
    // 총점 열이 없으면 항목 점수를 더해 만든다. 한 항목이라도 비어 있으면 총점은 비워 둔다.
    totalCol = {
      key: TOTAL_KEY, label: '총점', name: '총점', type: 'score', isTotal: true, computed: true,
      scale: items.every((c) => c.declaredMax) ? items.reduce((s, c) => s + c.declaredMax, 0) : null,
    };
    rows.forEach((r) => {
      const vals = items.map((c) => r.scores[c.key]);
      r.scores[TOTAL_KEY] = vals.some((v) => v === null) ? null : vals.reduce((s, v) => s + v, 0);
    });
    if (!totalCol.scale) {
      const observed = rows.map((r) => r.scores[TOTAL_KEY]).filter((v) => v !== null);
      totalCol.scale = pickScale(null, observed.length ? Math.max(...observed) : 100);
    }
    columns.push(totalCol);
  }

  return {
    columns, rows, items, totalCol,
    nameCol: columns.find((c) => NAME_HEADER.test(c.label)) || null,
    classCol: columns.find((c) => CLASS_HEADER.test(c.label)) || null,
    targets: totalCol ? [totalCol, ...items] : items,
  };
}

/* ---------- 통계 ---------- */

function computeStats(values) {
  const n = values.length;
  if (!n) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const mean = values.reduce((s, v) => s + v, 0) / n;
  const mid = Math.floor(n / 2);
  const median = n % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  const sd = Math.sqrt(values.reduce((s, v) => s + (v - mean) ** 2, 0) / n);
  return { n, mean, median, sd, min: sorted[0], max: sorted[n - 1] };
}

function buildBins(values, scale) {
  const count = scale >= 10 ? 10 : Math.max(1, Math.ceil(scale));
  const width = scale / count;
  const bins = Array.from({ length: count }, (_, i) => ({ lo: i * width, hi: (i + 1) * width, n: 0 }));
  values.forEach((v) => {
    const i = Math.min(count - 1, Math.max(0, Math.floor(v / width)));
    bins[i].n++;
  });
  return bins;
}

const fmt = (v) => (Number.isInteger(v) ? String(v) : v.toFixed(1));
const fmtStat = (v) => (Math.round(v * 10) / 10).toFixed(1).replace(/\.0$/, '');

/* ---------- 화면 그리기 ---------- */

function el(tag, attrs = {}, text) {
  const node = document.createElement(tag);
  Object.entries(attrs).forEach(([k, v]) => node.setAttribute(k, v));
  if (text !== undefined) node.textContent = text;
  return node;
}

function svg(tag, attrs = {}, text) {
  const node = document.createElementNS(SVG_NS, tag);
  Object.entries(attrs).forEach(([k, v]) => node.setAttribute(k, v));
  if (text !== undefined) node.textContent = text;
  return node;
}

function visibleRows() {
  const { data, classValue } = state;
  if (!data.classCol || !classValue) return data.rows;
  return data.rows.filter((r) => (r.cells[data.classCol.idx] || '') === classValue);
}

function displayName(row) {
  const { nameCol } = state.data;
  return nameCol ? row.cells[nameCol.idx] || '' : '';
}

function namesWith(rows, key, value) {
  const names = rows.filter((r) => r.scores[key] === value).map(displayName).filter(Boolean);
  if (!names.length) return '';
  return names.length > 2 ? `${names.slice(0, 2).join(', ')} 외 ${names.length - 2}명` : names.join(', ');
}

function renderKpis(rows, target, stats) {
  const box = $('kpis');
  box.replaceChildren();
  const missing = rows.length - (stats ? stats.n : 0);
  const cards = stats ? [
    ['평균', fmtStat(stats.mean), `만점 ${fmt(target.scale)}점 기준 ${fmtStat((stats.mean / target.scale) * 100)}%`, true],
    ['최고점', fmt(stats.max), namesWith(rows, target.key, stats.max)],
    ['최저점', fmt(stats.min), namesWith(rows, target.key, stats.min)],
    ['중앙값', fmtStat(stats.median), '점수순 가운데 값'],
    ['표준편차', fmtStat(stats.sd), '점수가 흩어진 정도'],
    ['응시 인원', String(stats.n), missing ? `점수 없음 ${missing}명 제외` : '전원 점수 있음', false, '명'],
  ] : [['응시 인원', '0', '분석할 점수가 없습니다', false, '명']];

  cards.forEach(([label, value, detail, main, unit]) => {
    const card = el('div', { class: main ? 'kpi main' : 'kpi' });
    card.append(el('div', { class: 'kpi-label' }, label));
    const v = el('div', { class: 'kpi-value' }, value);
    v.append(el('small', {}, unit || '점'));
    card.append(v);
    const d = el('div', { class: 'kpi-detail' }, detail || ' ');
    if (detail) d.title = detail;
    card.append(d);
    box.append(card);
  });
}

function niceStep(max) {
  if (max <= 5) return 1;
  if (max <= 10) return 2;
  if (max <= 25) return 5;
  return Math.ceil(max / 5 / 5) * 5;
}

// 위쪽 모서리만 둥근 막대
function barPath(x, y, w, h, r) {
  r = Math.min(r, h, w / 2);
  return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`;
}

function renderHistogram(target, values, stats) {
  const root = $('histogram');
  root.replaceChildren();
  hideTooltip();
  // 글자가 작아지지 않도록 viewBox를 실제 표시 너비에 맞춘다.
  const W = Math.max(300, Math.round($('histWrap').clientWidth) || 720);
  const H = W < 480 ? 260 : 320, m = { top: 52, right: 12, bottom: 46, left: 34 };
  const pw = W - m.left - m.right, ph = H - m.top - m.bottom;
  root.setAttribute('viewBox', `0 0 ${W} ${H}`);

  const bins = buildBins(values, target.scale);
  const maxCount = Math.max(1, ...bins.map((b) => b.n));
  const step = niceStep(maxCount);
  const yMax = Math.ceil(maxCount / step) * step;
  const x = (v) => m.left + (v / target.scale) * pw;
  const y = (c) => m.top + ph - (c / yMax) * ph;

  for (let t = 0; t <= yMax; t += step) {
    root.append(svg('line', { class: t ? 'gridline' : 'axis', x1: m.left, x2: W - m.right, y1: y(t), y2: y(t) }));
    root.append(svg('text', { x: m.left - 8, y: y(t) + 4, 'text-anchor': 'end' }, String(t)));
  }
  root.append(svg('text', { class: 'axis-title', x: m.left - 8, y: m.top - 16, 'text-anchor': 'end' }, '명'));

  const band = pw / bins.length;
  const gap = Math.min(8, band * 0.14);
  bins.forEach((b, i) => {
    const bx = x(b.lo) + gap / 2, bw = band - gap, by = y(b.n), bh = y(0) - by;
    const last = i === bins.length - 1;
    const hit = svg('rect', { class: 'hit', x: x(b.lo), y: m.top, width: band, height: ph });
    const bar = svg('path', { class: 'bar', d: b.n ? barPath(bx, by, bw, bh, 4) : '' });
    const range = `${fmt(b.lo)}점 이상 ${fmt(b.hi)}점 ${last ? '이하' : '미만'}`;
    const share = values.length ? fmtStat((b.n / values.length) * 100) : '0';
    hit.addEventListener('mouseenter', () => showTooltip(bar, hit, range, `${b.n}명 · ${share}%`));
    hit.addEventListener('mouseleave', hideTooltip);
    root.append(hit, bar);
    if (b.n) root.append(svg('text', { class: 'count', x: bx + bw / 2, y: by - 6, 'text-anchor': 'middle' }, String(b.n)));
    root.append(svg('text', { x: x(b.lo), y: H - m.bottom + 18, 'text-anchor': 'middle' }, fmt(b.lo)));
  });
  root.append(svg('text', { x: x(target.scale), y: H - m.bottom + 18, 'text-anchor': 'middle' }, fmt(target.scale)));
  root.append(svg('text', { class: 'axis-title', x: m.left + pw / 2, y: H - 6, 'text-anchor': 'middle' }, `${target.name} 점수 구간`));

  if (stats) {
    const mx = x(Math.min(stats.mean, target.scale));
    root.append(svg('line', { class: 'mean-line', x1: mx, x2: mx, y1: 22, y2: y(0) }));
    const anchor = mx > W - 90 ? 'end' : mx < m.left + 50 ? 'start' : 'middle';
    root.append(svg('text', { class: 'mean-label', x: mx, y: 14, 'text-anchor': anchor }, `평균 ${fmtStat(stats.mean)}`));
  }

  $('histNote').textContent = `${target.name} · ${values.length}명 · 구간 폭 ${fmt(target.scale / bins.length)}점`;
}

function showTooltip(bar, hit, title, detail) {
  const tip = $('tooltip');
  const wrap = $('histWrap').getBoundingClientRect();
  const box = (bar.getAttribute('d') ? bar : hit).getBoundingClientRect();
  tip.replaceChildren(el('b', {}, title), el('span', {}, detail));
  tip.hidden = false;
  const half = tip.offsetWidth / 2;
  const cx = Math.min(wrap.width - half, Math.max(half, box.left + box.width / 2 - wrap.left));
  tip.style.left = `${cx}px`;
  tip.style.top = `${Math.max(tip.offsetHeight, box.top - wrap.top - 8)}px`;
  bar.classList.add('active');
}

function hideTooltip() {
  $('tooltip').hidden = true;
  document.querySelectorAll('#histogram .bar.active').forEach((b) => b.classList.remove('active'));
}

function renderItems(rows) {
  const box = $('items');
  box.replaceChildren();
  const { items } = state.data;
  $('itemCard').hidden = items.length < 2;
  items.forEach((c) => {
    const values = rows.map((r) => r.scores[c.key]).filter((v) => v !== null);
    if (!values.length) return;
    const mean = values.reduce((s, v) => s + v, 0) / values.length;
    const pct = Math.max(0, Math.min(100, (mean / c.scale) * 100));
    const row = el('div', { class: 'item-row' });
    row.append(el('span', { class: 'item-name' }, c.name));
    const value = el('span', { class: 'item-value' });
    value.append(el('b', {}, fmtStat(mean)), ` / ${fmt(c.scale)}점 · ${fmtStat(pct)}%`);
    const track = el('div', { class: 'item-track' });
    const fill = el('div', { class: 'item-fill' });
    fill.style.width = `${pct}%`;
    track.append(fill);
    row.append(value, track);
    box.append(row);
  });
  if (!box.children.length) box.append(el('p', { class: 'empty' }, '표시할 항목 점수가 없습니다.'));
}

function renderTable(rows, target) {
  const { columns } = state.data;
  const thead = $('table').tHead, tbody = $('table').tBodies[0];
  thead.replaceChildren();
  tbody.replaceChildren();

  // 순위는 분석 대상 점수 기준, 동점은 같은 순위
  const ranked = rows.filter((r) => r.scores[target.key] !== null)
    .sort((a, b) => b.scores[target.key] - a.scores[target.key]);
  const rank = new Map();
  ranked.forEach((r, i) => {
    const prev = ranked[i - 1];
    rank.set(r, prev && prev.scores[target.key] === r.scores[target.key] ? rank.get(prev) : i + 1);
  });

  const valueOf = (row, key) => {
    if (key === '__rank') return rank.has(row) ? rank.get(row) : null;
    const col = columns.find((c) => c.key === key);
    if (col.type === 'score') return row.scores[key];
    const raw = row.cells[col.idx] || '';
    const n = Number(raw);
    return raw !== '' && Number.isFinite(n) ? n : raw;
  };

  const headRow = el('tr');
  [{ key: '__rank', label: '순위', type: 'score' }, ...columns].forEach((c) => {
    const th = el('th', c.type === 'text' ? { class: 'text' } : {});
    const btn = el('button', { type: 'button' }, c.label);
    btn.addEventListener('click', () => {
      if (state.sortKey === c.key) state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
      else { state.sortKey = c.key; state.sortDir = c.type === 'text' || c.key === '__rank' ? 'asc' : 'desc'; }
      render();
    });
    if (state.sortKey === c.key) th.setAttribute('aria-sort', state.sortDir === 'asc' ? 'ascending' : 'descending');
    th.append(btn);
    headRow.append(th);
  });
  thead.append(headRow);

  const q = state.query.trim().toLowerCase();
  const textCols = columns.filter((c) => c.type === 'text');
  let list = q
    ? rows.filter((r) => textCols.some((c) => (r.cells[c.idx] || '').toLowerCase().includes(q)))
    : [...rows];

  const dir = state.sortDir === 'asc' ? 1 : -1;
  list.sort((a, b) => {
    const va = valueOf(a, state.sortKey), vb = valueOf(b, state.sortKey);
    if (va === null || va === '') return vb === null || vb === '' ? 0 : 1; // 빈 값은 항상 아래로
    if (vb === null || vb === '') return -1;
    if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
    return String(va).localeCompare(String(vb), 'ko', { numeric: true }) * dir;
  });

  list.forEach((r) => {
    const tr = el('tr');
    tr.append(el('td', rank.has(r) ? {} : { class: 'missing' }, rank.has(r) ? String(rank.get(r)) : '–'));
    columns.forEach((c) => {
      if (c.type === 'text') { tr.append(el('td', { class: 'text' }, r.cells[c.idx] || '')); return; }
      const v = r.scores[c.key];
      const cls = [c.key === target.key ? 'target' : '', v === null ? 'missing' : ''].join(' ').trim();
      tr.append(el('td', cls ? { class: cls } : {}, v === null ? '–' : fmt(v)));
    });
    tbody.append(tr);
  });
  if (!list.length) {
    const td = el('td', { class: 'text missing', colspan: String(columns.length + 1) }, '검색 결과가 없습니다.');
    const tr = el('tr');
    tr.append(td);
    tbody.append(tr);
  }
}

function render() {
  const { data } = state;
  const target = data.targets.find((c) => c.key === state.target) || data.targets[0];
  const rows = visibleRows();
  const values = rows.map((r) => r.scores[target.key]).filter((v) => v !== null);
  const stats = computeStats(values);
  renderKpis(rows, target, stats);
  renderHistogram(target, values, stats);
  renderItems(rows);
  renderTable(rows, target);
}

/* ---------- 불러오기 ---------- */

function showMessage(text) {
  const box = $('message');
  box.textContent = text || '';
  box.hidden = !text;
}

function loadText(text, fileName) {
  let data;
  try {
    data = buildData(parseCSV(text));
  } catch (e) {
    // 이전 파일의 결과가 새 파일의 결과로 오해되지 않도록 대시보드를 닫는다.
    state.data = null;
    $('dashboard').hidden = true;
    showMessage(`"${fileName}" 파일을 분석할 수 없습니다. ${e.message}`);
    return;
  }
  showMessage('');
  Object.assign(state, { data, target: data.targets[0].key, classValue: '', sortKey: '__rank', sortDir: 'asc', query: '' });
  $('search').value = '';

  $('fileName').textContent = fileName;
  $('fileMeta').textContent = `학생 ${data.rows.length}명 · 평가 항목 ${data.items.length}개`
    + (data.totalCol && data.totalCol.computed ? ' · 총점은 항목 합계로 계산' : '');

  const targetSelect = $('targetSelect');
  targetSelect.replaceChildren(...data.targets.map((c) => el('option', { value: c.key }, c.name)));

  const classSelect = $('classSelect');
  $('classFilterWrap').hidden = !data.classCol;
  if (data.classCol) {
    const classes = [...new Set(data.rows.map((r) => r.cells[data.classCol.idx] || '').filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, 'ko', { numeric: true }));
    classSelect.replaceChildren(el('option', { value: '' }, '전체'),
      ...classes.map((c) => el('option', { value: c }, /^\d+$/.test(c) ? `${c}반` : c)));
  }

  $('dashboard').hidden = false;
  render();
}

function loadFile(file) {
  if (!file) return;
  if (!/\.csv$/i.test(file.name) && !/csv|text/.test(file.type)) {
    showMessage('CSV 파일만 분석할 수 있습니다. 엑셀 파일은 "다른 이름으로 저장 → CSV"로 저장한 뒤 올려 주세요.');
    return;
  }
  const reader = new FileReader();
  reader.onload = () => loadText(decodeBuffer(reader.result), file.name);
  reader.onerror = () => showMessage('파일을 읽는 중 오류가 발생했습니다.');
  reader.readAsArrayBuffer(file);
}

function downloadSample() {
  // BOM을 붙여 엑셀에서 한글이 깨지지 않게 한다.
  const blob = new Blob(['﻿' + SAMPLE_CSV], { type: 'text/csv;charset=utf-8' });
  const a = el('a', { href: URL.createObjectURL(blob), download: 'sample_data.csv' });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

$('fileInput').addEventListener('change', (e) => { loadFile(e.target.files[0]); e.target.value = ''; });
$('sampleBtn').addEventListener('click', () => loadText(SAMPLE_CSV, '샘플 데이터 (가상 학생 30명)'));
$('downloadBtn').addEventListener('click', downloadSample);
$('targetSelect').addEventListener('change', (e) => { state.target = e.target.value; render(); });
$('classSelect').addEventListener('change', (e) => { state.classValue = e.target.value; render(); });
$('search').addEventListener('input', (e) => { state.query = e.target.value; renderTable(visibleRows(), state.data.targets.find((c) => c.key === state.target)); });

let resizeTimer = null;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => { if (state.data) render(); }, 120);
});

const dropzone = $('dropzone');
['dragenter', 'dragover'].forEach((type) => dropzone.addEventListener(type, (e) => {
  e.preventDefault();
  dropzone.classList.add('dragover');
}));
['dragleave', 'drop'].forEach((type) => dropzone.addEventListener(type, (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
}));
dropzone.addEventListener('drop', (e) => loadFile(e.dataTransfer.files[0]));
// 드롭 영역 밖에 떨어뜨려도 브라우저가 파일을 열어 버리지 않게 한다.
['dragover', 'drop'].forEach((type) => window.addEventListener(type, (e) => e.preventDefault()));
