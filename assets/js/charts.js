/*
 * Grafik ringan tanpa library: bar horizontal (HTML), bar bertumpuk kondisi,
 * dan grafik kolom (SVG). Semua punya tooltip (hover & fokus keyboard)
 * serta tampilan tabel sebagai alternatif.
 */
(function (root) {
  'use strict';
  const SVGNS = 'http://www.w3.org/2000/svg';

  function h(tag, attrs, children) {
    const el = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      if (k === 'class') el.className = attrs[k];
      else if (k === 'text') el.textContent = attrs[k];
      else el.setAttribute(k, attrs[k]);
    }
    (children || []).forEach(c => c && el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c));
    return el;
  }
  function s(tag, attrs) {
    const el = document.createElementNS(SVGNS, tag);
    for (const k in attrs) el.setAttribute(k, attrs[k]);
    return el;
  }

  // ---------- ikon status (selalu dipasangkan dengan label) ----------
  const ICON_PATH = {
    critical: '<circle cx="8" cy="8" r="7" fill="currentColor"/><path d="M5.5 5.5l5 5M10.5 5.5l-5 5" stroke="#fff" stroke-width="1.8" stroke-linecap="round"/>',
    serious: '<path d="M8 1.2l7 12.6H1z" fill="currentColor"/><path d="M8 6v3.6" stroke="#fff" stroke-width="1.8" stroke-linecap="round"/><circle cx="8" cy="11.6" r="1" fill="#fff"/>',
    warning: '<circle cx="8" cy="8" r="7" fill="currentColor"/><path d="M8 4.5v4.2" stroke="#1a1a19" stroke-width="1.8" stroke-linecap="round"/><circle cx="8" cy="11.2" r="1" fill="#1a1a19"/>',
    info: '<circle cx="8" cy="8" r="7" fill="currentColor"/><path d="M8 7.2v4.3" stroke="#fff" stroke-width="1.8" stroke-linecap="round"/><circle cx="8" cy="4.8" r="1" fill="#fff"/>',
    good: '<circle cx="8" cy="8" r="7" fill="currentColor"/><path d="M4.8 8.3l2.1 2.1 4.3-4.6" stroke="#fff" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/>'
  };
  const SEV_LABEL = { critical: 'Kritis', serious: 'Penting', warning: 'Perhatian', info: 'Info', good: 'Lolos' };
  function icon(sev, size) {
    const z = size || 14;
    return `<svg width="${z}" height="${z}" viewBox="0 0 16 16" aria-hidden="true">${ICON_PATH[sev] || ICON_PATH.info}</svg>`;
  }
  function sevChip(sev) {
    return `<span class="sev sev-${sev}">${icon(sev)}${SEV_LABEL[sev] || sev}</span>`;
  }

  // ---------- tooltip ----------
  let tipEl = null;
  function tip() {
    if (!tipEl) { tipEl = h('div', { class: 'tip hidden', role: 'tooltip' }); document.body.appendChild(tipEl); }
    return tipEl;
  }
  function showTip(evOrEl, title, rows) {
    const t = tip();
    t.innerHTML = '';
    t.appendChild(h('div', { class: 'tt', text: title }));
    rows.forEach(([k, v]) => t.appendChild(h('div', { class: 'tr' }, [h('span', { text: k }), h('b', { text: v })])));
    t.classList.remove('hidden');
    let x, y;
    if (evOrEl.clientX != null) { x = evOrEl.clientX; y = evOrEl.clientY; }
    else { const r = evOrEl.getBoundingClientRect(); x = r.left + r.width / 2; y = r.top; }
    const w = t.offsetWidth, hh = t.offsetHeight;
    let left = x + 14, top = y - hh - 10;
    if (left + w > window.innerWidth - 8) left = Math.max(8, x - w - 14);
    if (top < 8) top = y + 18;
    t.style.left = left + 'px';
    t.style.top = top + 'px';
  }
  function hideTip() { if (tipEl) tipEl.classList.add('hidden'); }
  function bindTip(el, title, rows) {
    el.addEventListener('pointermove', e => showTip(e, title, rows));
    el.addEventListener('pointerleave', hideTip);
    el.addEventListener('focus', () => showTip(el, title, rows));
    el.addEventListener('blur', hideTip);
  }

  function tableView(cols, rows) {
    const d = h('details', { class: 'tbl-view' });
    d.appendChild(h('summary', { text: 'Lihat sebagai tabel' }));
    const wrap = h('div', { class: 'tbl-wrap' });
    const t = h('table');
    const thead = h('thead', null, [h('tr', null, cols.map(c => h('th', { class: c.num ? 'num' : '', text: c.label })))]);
    const tbody = h('tbody');
    rows.forEach(r => tbody.appendChild(h('tr', null, cols.map(c => h('td', { class: c.num ? 'num' : '', text: c.get(r) })))));
    t.appendChild(thead); t.appendChild(tbody); wrap.appendChild(t); d.appendChild(wrap);
    return d;
  }

  /**
   * Bar horizontal. series: [{key, label, cls}] (1 seri = tanpa legenda).
   * data: [{label, values:{key:num}, total, tip:[[k,v]]}]
   */
  function hbars(container, { data, series, fmtValue, valueLabel, cols }) {
    container.innerHTML = '';
    if (series.length > 1) {
      container.appendChild(h('div', { class: 'legend' }, series.map(se => h('span', { class: 'k' }, [h('span', { class: 'sw ' + se.cls }), se.label]))));
    }
    const max = Math.max(...data.map(d => d.total), 1);
    const list = h('div', { class: 'hbars' });
    data.forEach(d => {
      const item = h('div', { class: 'hbar', tabindex: '0', 'aria-label': d.label + ': ' + (valueLabel ? valueLabel(d) : fmtValue(d.total)) });
      item.appendChild(h('div', { class: 'row' }, [h('span', { class: 'name', text: d.label }), h('span', { class: 'v', text: valueLabel ? valueLabel(d) : fmtValue(d.total) })]));
      const track = h('div', { class: 'track' });
      series.forEach(se => {
        const v = d.values[se.key] || 0;
        if (v <= 0) return;
        const seg = h('div', { class: 'seg ' + se.cls });
        seg.style.width = (v / max * 100) + '%';
        track.appendChild(seg);
      });
      item.appendChild(track);
      bindTip(item, d.label, d.tip || series.map(se => [se.label, fmtValue(d.values[se.key] || 0)]));
      list.appendChild(item);
    });
    container.appendChild(list);
    if (cols) container.appendChild(tableView(cols, data));
  }

  /** Bar bertumpuk 100% untuk kondisi (warna status + ikon + label). */
  function kondisiStack(container, items, total, fmtN, fmtPct) {
    container.innerHTML = '';
    const bar = h('div', { class: 'stack', role: 'img', 'aria-label': items.map(i => `${i.label} ${fmtN(i.n)}`).join(', ') });
    items.forEach(i => {
      if (!i.n) return;
      const seg = h('div', { class: 'seg ' + i.cls, tabindex: '0' });
      seg.style.flex = String(i.n);
      bindTip(seg, i.label, [['Jumlah', fmtN(i.n) + ' unit'], ['Porsi', fmtPct(i.n / total)]]);
      bar.appendChild(seg);
    });
    container.appendChild(bar);
    const lg = h('div', { class: 'kondisi-legend' });
    items.forEach(i => {
      const d = h('div');
      d.innerHTML = `<span class="sev sev-${i.sev}">${icon(i.sev)}${i.label}</span><b>${fmtN(i.n)}</b>${fmtPct(i.n / total)}`;
      lg.appendChild(d);
    });
    container.appendChild(lg);
  }

  function niceMax(v) {
    if (v <= 0) return 1;
    const p = Math.pow(10, Math.floor(Math.log10(v)));
    const m = v / p;
    const step = m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10;
    return step * p;
  }

  /** Grafik kolom SVG, satu seri. data: [{x, y, tip:[[k,v]]}] */
  function columns(container, { data, fmtTick, labelEvery, highlightMax, title, cols }) {
    container.innerHTML = '';
    const W = Math.max(280, container.clientWidth || 600);
    const H = 220, padL = 40, padR = 8, padT = 18, padB = 26;
    const max = niceMax(Math.max(...data.map(d => d.y), 1));
    const svg = s('svg', { class: 'colchart', viewBox: `0 0 ${W} ${H}`, width: W, height: H, role: 'img', 'aria-label': title || 'Grafik kolom' });
    const ph = H - padT - padB, pw = W - padL - padR;
    const ticks = 4;
    for (let i = 0; i <= ticks; i++) {
      const v = max * i / ticks;
      const y = padT + ph - (v / max) * ph;
      svg.appendChild(s('line', { class: i === 0 ? 'baseline' : 'gridline', x1: padL, x2: W - padR, y1: y, y2: y }));
      const t = s('text', { class: 'tick', x: padL - 6, y: y + 4, 'text-anchor': 'end' });
      t.textContent = fmtTick(v);
      svg.appendChild(t);
    }
    const band = pw / data.length;
    const bw = Math.max(2, Math.min(24, band * 0.7));
    const every = labelEvery || Math.ceil(data.length / Math.max(1, Math.floor(pw / 44)));
    let maxIdx = 0;
    data.forEach((d, i) => { if (d.y > data[maxIdx].y) maxIdx = i; });
    data.forEach((d, i) => {
      const cx = padL + band * i + band / 2;
      const bh = (d.y / max) * ph;
      const x = cx - bw / 2, y = padT + ph - bh;
      const hit = s('rect', { class: 'hit', x: padL + band * i, y: padT, width: band, height: ph, tabindex: '0', 'aria-label': `${d.x}: ${d.y}` });
      svg.appendChild(hit);
      if (bh > 0) {
        const r = Math.min(4, bw / 2, bh);
        const path = `M${x},${y + bh} L${x},${y + r} Q${x},${y} ${x + r},${y} L${x + bw - r},${y} Q${x + bw},${y} ${x + bw},${y + r} L${x + bw},${y + bh} Z`;
        svg.appendChild(s('path', { class: 'bar', d: path }));
      } else {
        svg.appendChild(s('g', {}));
      }
      const isLast = i === data.length - 1;
      if (i % every === 0 || (isLast && (i % every) >= every * 0.6)) {
        const t = s('text', { class: 'tick', x: cx, y: H - 8, 'text-anchor': 'middle' });
        t.textContent = d.x;
        svg.appendChild(t);
      }
      if (highlightMax && i === maxIdx && d.y > 0) {
        const t = s('text', { class: 'lbl', x: cx, y: y - 5, 'text-anchor': 'middle' });
        t.textContent = d.y;
        svg.appendChild(t);
      }
      bindTip(hit, String(d.x), d.tip || [['Nilai', String(d.y)]]);
    });
    container.appendChild(svg);
    if (cols) container.appendChild(tableView(cols, data));
  }

  root.Charts = { hbars, kondisiStack, columns, icon, sevChip, SEV_LABEL, tableView, hideTip };
})(window);
