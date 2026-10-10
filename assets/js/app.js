/* Pengendali tampilan dashboard. */
(function () {
  'use strict';
  const cfg = window.ASET_CONFIG;
  const A = window.AsetBMN;
  const C = window.Charts;
  const fmt = A.fmt;
  const esc = A.util.esc;
  const $ = id => document.getElementById(id);

  const state = {
    assets: [], stats: null, headers: [], fileName: '', sheet: '', acuan: null,
    filter: { q: '', jenis: '', kondisi: '', flag: '' },
    sort: { key: 'no', dir: 1 }, page: 0, pageSize: 50, tab: 'ringkasan'
  };

  // ---------------- Tema ----------------
  function initTheme() {
    let t = null;
    try { t = localStorage.getItem('tema-dashboard-aset'); } catch (e) { /* abaikan */ }
    if (t === 'light' || t === 'dark') document.documentElement.dataset.theme = t;
    $('btnTheme').addEventListener('click', () => {
      const cur = document.documentElement.dataset.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
      const nx = cur === 'dark' ? 'light' : 'dark';
      document.documentElement.dataset.theme = nx;
      try { localStorage.setItem('tema-dashboard-aset', nx); } catch (e) { /* abaikan */ }
    });
  }

  // ---------------- Muat file ----------------
  function setStatus(msg, isErr) {
    const el = $('status');
    el.textContent = msg || '';
    el.className = isErr ? 'err' : 'loading';
  }

  async function loadFile(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) { setStatus('File harus berformat .xlsx (Excel).', true); return; }
    setStatus('Membaca ' + file.name + '…');
    try {
      const buf = await file.arrayBuffer();
      const book = await window.XlsxReader.readWorkbook(buf, cfg.requiredHeaders);
      const norm = A.normalize(book, cfg);
      if (!norm.assets.length) throw new Error('Tidak ada baris aset di dalam file.');
      const stats = A.analyze(norm.assets, cfg, norm.headers);
      norm.assets.forEach(a => {
        a._hay = A.util.normText([a.nama, a.merk, a.tipe, a.kodeBarang, a.nup, a.lokasi, a.noPolisi, a.jenis, a.raw['Nama'], a.raw['Kode Register']].join(' '));
        a._hayRapat = a._hay.replace(/ /g, '');
      });
      Object.assign(state, { assets: norm.assets, headers: norm.headers, acuan: norm.acuan, stats, fileName: file.name, sheet: book.sheetName });
      state.filter = { q: '', jenis: '', kondisi: '', flag: '', ids: null, idsLabel: '' };
      state.qs = '';
      state.page = 0;
      setStatus('');
      renderAll();
    } catch (e) {
      console.error(e);
      setStatus(e.message || String(e), true);
    }
  }

  function initLoader() {
    const input = $('file');
    input.addEventListener('change', () => loadFile(input.files[0]));
    const drop = $('drop');
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('drag'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('drag'); }));
    drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) loadFile(f); });
    $('btnGanti').addEventListener('click', () => {
      $('dash').classList.add('hidden');
      $('empty').classList.remove('hidden');
      $('btnGanti').classList.add('hidden');
      input.value = '';
      $('meta').textContent = 'Ringkasan, kesimpulan, dan hal yang perlu diperhatikan dari Daftar Aset';
      state.assets = []; state.stats = null;
    });
  }

  // ---------------- Tabs ----------------
  const tabRender = {};
  function showTab(name) {
    state.tab = name;
    document.querySelectorAll('.tab').forEach(t => t.setAttribute('aria-selected', String(t.dataset.tab === name)));
    document.querySelectorAll('[role="tabpanel"]').forEach(p => p.classList.toggle('hidden', p.id !== 'p-' + name));
    if (tabRender[name]) tabRender[name]();
    C.hideTip();
  }
  function initTabs() {
    document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => showTab(t.dataset.tab)));
    let to = null;
    window.addEventListener('resize', () => {
      clearTimeout(to);
      to = setTimeout(() => { if (state.stats && tabRender[state.tab]) tabRender[state.tab](); }, 150);
    });
  }

  // ---------------- Komponen kecil ----------------
  function tile(label, value, note, hero) {
    return `<div class="tile${hero ? ' hero' : ''}"><div class="lbl">${esc(label)}</div><div class="val">${esc(value)}</div>${note ? `<div class="note">${esc(note)}</div>` : ''}</div>`;
  }
  function checkCard(c) {
    if (c.unavailable) {
      return `<div class="check ok"><div class="head"><span class="title">${esc(c.label)}</span>${C.sevChip('info')}</div>
        <div class="desc">Tidak dijalankan: kolom ${esc(c.need.join(', '))} tidak ada di file.</div></div>`;
    }
    const ok = c.count === 0;
    const sev = ok ? 'good' : c.sev;
    return `<div class="check${ok ? ' ok' : ''}">
      <div class="head"><span class="title">${esc(c.label)}</span>${C.sevChip(sev)}</div>
      <div class="big">${fmt.n(c.count)} <small>dari ${fmt.n(c.applicable)} aset relevan</small></div>
      <div class="desc">${esc(c.desc)}</div>
      ${ok ? '' : `<div><button class="linkbtn" type="button" data-flag="${c.id}">Lihat daftar</button></div>`}
    </div>`;
  }
  function bindFlagLinks(root) {
    root.querySelectorAll('[data-flag]').forEach(b => b.addEventListener('click', () => goDetail(b.dataset.flag)));
  }
  function kondisiHtml(k) {
    if (k === 'Rusak Berat') return `<span class="kond sev sev-critical">${C.icon('critical', 12)}Rusak Berat</span>`;
    if (k === 'Rusak Ringan') return `<span class="kond sev sev-warning">${C.icon('warning', 12)}Rusak Ringan</span>`;
    return esc(k);
  }
  function flagBadges(a, skip) {
    return a.flags.filter(f => !(skip || []).includes(f)).map(f => `<span class="badge ${A.FLAG[f].sev}">${esc(A.FLAG[f].label)}</span>`).join('');
  }
  const th = (label, num) => `<th${num ? ' class="num"' : ''}>${esc(label)}</th>`;
  const shortGuna = g => g.replace(/^Digunakan sendiri untuk /i, 'Sendiri: ').replace(/^Digunakan oleh satker lain.*$/i, 'Satker lain (1 K/L)');
  const merkTipe = (m, t) => (m && t && A.util.normText(m) === A.util.normText(t) ? m : [m, t].filter(Boolean).join(' · '));
  const umurStr = u => (u == null ? '–' : new Intl.NumberFormat('id-ID', { maximumFractionDigits: 1 }).format(u));

  // ---------------- Render semua ----------------
  function renderAll() {
    $('empty').classList.add('hidden');
    $('dash').classList.remove('hidden');
    $('btnGanti').classList.remove('hidden');
    const satker = [...new Set(state.assets.map(a => A.util.str(a.raw['Nama Satker'])).filter(Boolean))];
    $('meta').textContent = [satker.length === 1 ? satker[0] : satker.length + ' satker', fmt.n(state.assets.length) + ' aset', state.fileName].join(' · ');
    $('foot').textContent = `Sheet "${state.sheet}" · umur aset dihitung per ${fmt.tgl(state.acuan)} · data hanya diproses di browser ini · aturan pemeriksaan dapat diubah di assets/js/config.js`;

    const nPerhatian = state.stats.perhatian.filter(p => p.sev === 'critical' || p.sev === 'serious').length;
    const cnt = $('perhatianCount');
    cnt.textContent = nPerhatian;
    cnt.classList.toggle('hidden', !nPerhatian);

    tabRender.ringkasan = renderRingkasan;
    tabRender.perhatian = renderPerhatian;
    tabRender.umur = renderUmur;
    tabRender.detail = renderDetail;
    initDetailControls();
    initQuick();
    showTab('ringkasan');
  }

  // ---------------- Ringkasan ----------------
  function renderRingkasan() {
    const s = state.stats, t = s.total;
    const dataAssets = state.assets.filter(a => a.flags.some(f => A.FLAG[f].group === 'data')).length;
    const dataChecks = s.checks.filter(c => c.group === 'data' && c.count > 0).length;
    $('kpi').innerHTML = [
      tile('Jumlah aset', fmt.n(t.n), `${fmt.n(t.intra)} intrakomptabel · ${fmt.n(t.n - t.intra)} ekstrakomptabel`),
      tile('Nilai perolehan', fmt.rpc(t.perolehan), fmt.rp(t.perolehan)),
      tile('Nilai buku', fmt.rpc(t.buku), `${fmt.pct(t.buku / t.perolehan)} dari nilai perolehan`, true),
      tile('Kandidat penghapusan', fmt.n(s.hapus.n) + ' unit', 'Nilai perolehan ' + fmt.rpc(s.hapus.perolehan)),
      tile('Data perlu dilengkapi', fmt.n(dataAssets) + ' aset', `${dataChecks} jenis pemeriksaan bermasalah`)
    ].join('');

    renderQuick();

    $('kesimpulan').innerHTML = s.kesimpulan.map(k => `<li>${k}</li>`).join('');

    const al = $('alerts');
    if (!s.perhatian.length) {
      al.innerHTML = `<li><span>${C.sevChip('good')}</span><div><div class="t">Tidak ada temuan</div><div class="d">Semua pemeriksaan lolos.</div></div><div></div></li>`;
    } else {
      al.innerHTML = s.perhatian.map(p => `<li>
        <span>${C.sevChip(p.sev)}</span>
        <div><div class="t">${esc(p.label)}</div><div class="d">${esc(p.desc)}</div></div>
        <div class="n">${fmt.n(p.count)}<button class="linkbtn" type="button" data-flag="${p.id}">Lihat</button></div>
      </li>`).join('');
      bindFlagLinks(al);
    }

    C.kondisiStack($('chKondisi'), [
      { label: 'Baik', n: s.kondisi['Baik'] || 0, cls: 'good', sev: 'good' },
      { label: 'Rusak Ringan', n: s.kondisi['Rusak Ringan'] || 0, cls: 'warning', sev: 'warning' },
      { label: 'Rusak Berat', n: s.kondisi['Rusak Berat'] || 0, cls: 'critical', sev: 'critical' }
    ], t.n, fmt.n, fmt.pct);

    const byN = s.byJenis.slice().sort((a, b) => b.n - a.n);
    C.hbars($('chJumlah'), {
      data: byN.map(j => ({ label: j.label, total: j.n, values: { n: j.n }, tip: [['Unit', fmt.n(j.n)], ['Rusak Berat', fmt.n(j.rb)], ['Nilai perolehan', fmt.rpc(j.perolehan)]] })),
      series: [{ key: 'n', label: 'Unit', cls: 's1' }],
      fmtValue: fmt.n,
      cols: [{ label: 'Jenis BMN', get: d => d.label }, { label: 'Unit', num: true, get: d => fmt.n(d.total) }]
    });

    C.hbars($('chNilai'), {
      data: s.byJenis.map(j => {
        const susut = Math.max(0, j.perolehan - j.buku);
        return {
          label: j.label, total: j.perolehan, values: { buku: j.buku, susut },
          tip: [['Nilai perolehan', fmt.rpc(j.perolehan)], ['Nilai buku', fmt.rpc(j.buku)], ['Sudah disusutkan', fmt.rpc(susut)], ['Sisa nilai', j.perolehan ? fmt.pct(j.buku / j.perolehan) : '–']]
        };
      }),
      series: [{ key: 'buku', label: 'Nilai buku', cls: 's1' }, { key: 'susut', label: 'Sudah disusutkan', cls: 's2' }],
      fmtValue: fmt.rpc,
      valueLabel: d => `${fmt.rpc(d.total)} · sisa ${d.total ? fmt.pct(d.values.buku / d.total, 0) : '–'}`,
      cols: [
        { label: 'Jenis BMN', get: d => d.label },
        { label: 'Nilai perolehan', num: true, get: d => fmt.rp(d.total) },
        { label: 'Nilai buku', num: true, get: d => fmt.rp(d.values.buku) },
        { label: 'Sudah disusutkan', num: true, get: d => fmt.rp(d.values.susut) }
      ]
    });
  }


  // ---------------- Cari cepat (Ringkasan) ----------------
  function initQuick() {
    const inp = $('qCari');
    // Saran ketik: nama barang yang ada di file, urut dari yang paling banyak.
    const cnt = new Map();
    state.assets.forEach(a => cnt.set(a.nama, (cnt.get(a.nama) || 0) + 1));
    const top = [...cnt].sort((x, y) => y[1] - x[1]);
    $('qSaran').innerHTML = top.map(([n, c]) => `<option value="${esc(n)}">${fmt.n(c)} unit</option>`).join('');
    $('qContoh').innerHTML = '<span>Contoh:</span>' + top.slice(0, 4).map(([n]) => `<button class="chip" type="button" data-q="${esc(n)}">${esc(n)}</button>`).join('');
    $('qContoh').querySelectorAll('[data-q]').forEach(b => b.addEventListener('click', () => { inp.value = b.dataset.q; state.qs = inp.value; renderQuick(); inp.focus(); }));
    inp.value = state.qs || '';
    if (!initQuick.bound) {
      initQuick.bound = true;
      let to = null;
      inp.addEventListener('input', () => { clearTimeout(to); to = setTimeout(() => { state.qs = inp.value; renderQuick(); }, 200); });
      inp.addEventListener('keydown', e => { if (e.key === 'Enter') { clearTimeout(to); state.qs = inp.value; renderQuick(); } if (e.key === 'Escape') { inp.value = ''; state.qs = ''; renderQuick(); } });
    }
  }

  function miniList(title, items, max) {
    if (!items.length) return '';
    const rest = items.length - max;
    return `<div class="qs-box"><h3>${esc(title)}</h3><ul class="mini">${items.slice(0, max).map(i => `<li><span>${esc(i.label)}</span><b>${fmt.n(i.n)}</b></li>`).join('')}${rest > 0 ? `<li class="muted"><span>${rest} lainnya</span><b></b></li>` : ''}</ul></div>`;
  }

  function renderQuick() {
    const out = $('qHasil');
    const q = (state.qs || '').trim();
    if (A.util.normText(q).length < 2) { out.innerHTML = ''; return; }
    const r = A.ringkasCari(state.assets, q, cfg);
    if (!r.n) {
      out.innerHTML = `<div class="qs-empty">Tidak ada aset yang cocok dengan “${esc(q)}”. Coba kata yang lebih pendek (mis. hanya “pompa”) atau pilih dari saran saat mengetik.</div>`;
      return;
    }
    const k = r.kondisi;
    const tahun = r.tahunMin == null ? '–' : (r.tahunMin === r.tahunMax ? String(r.tahunMin) : `${r.tahunMin}–${r.tahunMax}`);
    const maxG = 10;
    out.innerHTML = `
      <div class="qs-head">
        <div><span class="qs-n">${fmt.n(r.n)} unit</span> cocok dengan “${esc(q)}”</div>
        <div class="qs-act"><button class="btn sm primary" type="button" id="qLihat">Lihat ${fmt.n(r.n)} aset di Detail</button><button class="btn sm" type="button" id="qBersih">Bersihkan</button></div>
      </div>
      <div class="tiles qs-tiles">
        ${tile('Jumlah', fmt.n(r.n) + ' unit', `${r.grup.length} variasi nama/merk/tipe`)}
        ${tile('Kondisi Baik', fmt.n(k['Baik'] || 0) + ' unit', `Rusak Ringan ${fmt.n(k['Rusak Ringan'] || 0)} · Rusak Berat ${fmt.n(k['Rusak Berat'] || 0)}`)}
        ${tile('Nilai perolehan', fmt.rpc(r.perolehan), fmt.rp(r.perolehan))}
        ${tile('Nilai buku', fmt.rpc(r.buku), `${fmt.n(r.habis)} unit sudah Rp0`)}
        ${tile('Tahun perolehan', tahun, r.umurRata == null ? '' : `Rata-rata umur ${umurStr(r.umurRata)} tahun`)}
        ${tile('Kandidat penghapusan', fmt.n(r.hapus) + ' unit', r.hapus ? 'Rusak Berat / Henti Guna / Usul Hapus' : 'Tidak ada')}
      </div>
      <div id="qKondisi"></div>
      <h3>Rincian barang yang cocok</h3>
      <div class="tbl-wrap"><table class="resp">
        <thead><tr>${th('Nama barang')}${th('Merk / tipe')}${th('Unit', 1)}${th('Baik', 1)}${th('Rusak Ringan', 1)}${th('Rusak Berat', 1)}${th('Nilai perolehan', 1)}</tr></thead>
        <tbody>${r.grup.slice(0, maxG).map(g => `<tr>
          <td data-label="Nama barang"><div>${esc(g.nama)}</div></td>
          <td data-label="Merk / tipe"><div>${esc(merkTipe(g.merk, g.tipe)) || '–'}</div></td>
          <td data-label="Unit" class="num"><div><b>${fmt.n(g.n)}</b></div></td>
          <td data-label="Baik" class="num"><div>${fmt.n(g.baik)}</div></td>
          <td data-label="Rusak Ringan" class="num"><div>${fmt.n(g.rr)}</div></td>
          <td data-label="Rusak Berat" class="num"><div>${fmt.n(g.rb)}</div></td>
          <td data-label="Nilai perolehan" class="num"><div>${fmt.rp(g.perolehan)}</div></td></tr>`).join('')}</tbody>
      </table></div>
      ${r.grup.length > maxG ? `<p class="note-box">Ditampilkan ${maxG} dari ${r.grup.length} variasi. Perjelas kata kunci atau buka di Detail untuk daftar lengkap.</p>` : ''}
      <div class="qs-grid">
        ${miniList('Lokasi ruang', r.lokasi, 5)}
        ${miniList('Status penggunaan', r.guna, 4)}
        ${r.penanda.length ? `<div class="qs-box"><h3>Penanda</h3><ul class="mini">${r.penanda.map(p => `<li><span>${C.sevChip(p.sev)} ${esc(p.label)}</span><b>${fmt.n(p.n)}</b></li>`).join('')}</ul></div>` : miniList('Jenis BMN', r.jenis, 4)}
      </div>`;

    C.kondisiStack($('qKondisi'), [
      { label: 'Baik', n: k['Baik'] || 0, cls: 'good', sev: 'good' },
      { label: 'Rusak Ringan', n: k['Rusak Ringan'] || 0, cls: 'warning', sev: 'warning' },
      { label: 'Rusak Berat', n: k['Rusak Berat'] || 0, cls: 'critical', sev: 'critical' }
    ], r.n, fmt.n, fmt.pct);

    $('qLihat').addEventListener('click', () => {
      resetFilter();
      state.filter.ids = new Set(r.ids);
      state.filter.idsLabel = q;
      showTab('detail');
      window.scrollTo({ top: 0 });
    });
    $('qBersih').addEventListener('click', () => { $('qCari').value = ''; state.qs = ''; renderQuick(); $('qCari').focus(); });
  }

  // ---------------- Perlu perhatian ----------------
  function renderPerhatian() {
    const s = state.stats, h = s.hapus;
    $('hapusTiles').innerHTML = [
      tile('Jumlah unit', fmt.n(h.n), `${fmt.pct(h.n / s.total.n)} dari seluruh aset`),
      tile('Nilai perolehan', fmt.rpc(h.perolehan), fmt.rp(h.perolehan)),
      tile('Nilai buku', fmt.rpc(h.buku), h.buku === 0 ? 'Seluruhnya sudah habis disusutkan' : fmt.rp(h.buku))
    ].join('') + `<div class="tile"><div class="lbl">Daftar lengkap</div><div class="val"><button class="btn sm" type="button" data-flag="hapus">Lihat ${fmt.n(h.n)} aset</button></div><div class="note">Bisa diunduh sebagai CSV</div></div>`;
    bindFlagLinks($('hapusTiles'));

    const ch = s.check;
    $('hapusChecks').innerHTML = ['rb_belum_usul', 'henti_belum_usul', 'usul_tapi_baik', 'rusak_ringan'].map(id => checkCard(ch[id])).join('');
    bindFlagLinks($('hapusChecks'));

    const top = h.byNama.slice(0, 10);
    C.hbars($('chHapusNama'), {
      data: top.map(x => ({ label: x.label, total: x.n, values: { n: x.n }, tip: [['Unit', fmt.n(x.n)], ['Nilai perolehan', fmt.rpc(x.perolehan)]] })),
      series: [{ key: 'n', label: 'Unit', cls: 's1' }], fmtValue: fmt.n,
      cols: [{ label: 'Nama barang', get: d => d.label }, { label: 'Unit', num: true, get: d => fmt.n(d.total) }]
    });
    C.hbars($('chHapusJenis'), {
      data: h.byJenis.map(x => ({ label: x.label, total: x.n, values: { n: x.n }, tip: [['Unit', fmt.n(x.n)], ['Nilai perolehan', fmt.rpc(x.perolehan)]] })),
      series: [{ key: 'n', label: 'Unit', cls: 's1' }], fmtValue: fmt.n,
      cols: [{ label: 'Jenis BMN', get: d => d.label }, { label: 'Unit', num: true, get: d => fmt.n(d.total) }]
    });

    const sevOrder = c => (c.count ? A.SEV_RANK[c.sev] : 9);
    $('dataChecks').innerHTML = s.checks.filter(c => c.group === 'data').sort((a, b) => sevOrder(a) - sevOrder(b) || b.count - a.count).map(checkCard).join('');
    bindFlagLinks($('dataChecks'));
    $('gunaChecks').innerHTML = s.checks.filter(c => c.group === 'guna').map(checkCard).join('');
    bindFlagLinks($('gunaChecks'));
  }

  // ---------------- Kendaraan & umur ----------------
  function renderUmur() {
    const s = state.stats, k = s.kendaraan;
    const ku = k.filter(a => a.umur != null).map(a => a.umur);
    const kHabis = k.filter(a => a.flags.includes('habis_susut')).length;
    const kStnk = k.filter(a => a.flags.includes('kend_stnk')).length;
    $('umurTiles').innerHTML = [
      tile('Kendaraan bermotor', fmt.n(k.length) + ' unit', `${fmt.n(k.filter(a => a.kondisi === 'Baik').length)} kondisi Baik`),
      tile('Rata-rata umur kendaraan', ku.length ? umurStr(ku.reduce((x, y) => x + y, 0) / ku.length) + ' tahun' : '–', ku.length ? `Tertua ${umurStr(Math.max(...ku))} tahun` : ''),
      tile('Kendaraan nilai buku Rp0', fmt.n(kHabis) + ' unit', `${k.length ? fmt.pct(kHabis / k.length, 0) : '–'} dari kendaraan`),
      tile(`Aset tua (≥ ${cfg.umurTuaTahun} tahun)`, fmt.n(s.check.tua.count) + ' unit', 'Peralatan & kendaraan'),
      tile('Kendaraan tanpa STNK', fmt.n(kStnk) + ' unit', 'Kolom No STNK kosong')
    ].join('');

    C.columns($('chTahun'), {
      data: s.byTahun.map(y => ({ x: y.tahun, y: y.n, tip: [['Unit', fmt.n(y.n)], ['Nilai perolehan', fmt.rpc(y.perolehan)]] })),
      fmtTick: v => fmt.n(v), highlightMax: true, title: 'Jumlah aset per tahun perolehan',
      cols: [{ label: 'Tahun', get: d => String(d.x) }, { label: 'Unit', num: true, get: d => fmt.n(d.y) }]
    });

    $('tblUmur').innerHTML = `<div class="tbl-wrap"><table>
      <thead><tr>${th('Jenis BMN')}${th('Unit', 1)}${th('Rata-rata umur (th)', 1)}${th('Tertua (th)', 1)}${th('Nilai buku Rp0 (intra)', 1)}${th('Aset tua', 1)}</tr></thead>
      <tbody>${s.umurPerJenis.map(u => `<tr>
        <td>${esc(u.label)}</td><td class="num">${fmt.n(u.n)}</td><td class="num">${umurStr(u.rata)}</td><td class="num">${umurStr(u.tertua)}</td>
        <td class="num">${fmt.n(u.habis)} dari ${fmt.n(u.intra)}${u.intra ? ' (' + fmt.pct(u.habis / u.intra, 0) + ')' : ''}</td><td class="num">${fmt.n(u.tua)}</td></tr>`).join('')}</tbody>
    </table></div>
    <div class="note-box">"Nilai buku Rp0" berarti masa manfaat akuntansi sudah habis, bukan berarti barang rusak. "Aset tua" memakai ambang kerja ${cfg.umurTuaTahun} tahun sejak perolehan (bukan masa manfaat resmi); ubah di config.js bila perlu.</div>`;

    $('kendSub').textContent = `${fmt.n(k.length)} kendaraan, urut dari yang paling lama. Klik baris untuk melihat detail lengkap.`;
    $('tblKend').innerHTML = `<div class="tbl-wrap tall"><table class="resp">
      <thead><tr>${th('Nama barang')}${th('Merk / tipe')}${th('No Polisi')}${th('Tahun', 1)}${th('Umur (th)', 1)}${th('Kondisi')}${th('Nilai perolehan', 1)}${th('Nilai buku', 1)}${th('Penggunaan')}${th('Penanda')}</tr></thead>
      <tbody>${k.map(a => `<tr class="click" data-id="${a.id}" tabindex="0">
        <td data-label="Nama barang"><div>${esc(a.nama)}</div></td><td data-label="Merk / tipe"><div>${esc(merkTipe(a.merk, a.tipe))}</div></td><td data-label="No Polisi" class="nowrap"><div>${esc(a.noPolisi) || '–'}</div></td>
        <td data-label="Tahun" class="num"><div>${a.tahun || '–'}</div></td><td data-label="Umur (th)" class="num"><div>${umurStr(a.umur)}</div></td><td data-label="Kondisi"><div>${kondisiHtml(a.kondisi)}</div></td>
        <td data-label="Nilai perolehan" class="num"><div>${fmt.rp(a.nilaiPerolehan)}</div></td><td data-label="Nilai buku" class="num"><div>${fmt.rp(a.nilaiBuku)}</div></td><td data-label="Penggunaan"><div>${esc(shortGuna(a.statusGuna))}</div></td>
        <td data-label="Penanda"><div>${flagBadges(a, ['tua', 'habis_susut']) || '–'}</div></td></tr>`).join('')}</tbody>
    </table></div>`;
    bindRowClicks($('tblKend'));
  }

  // ---------------- Detail ----------------
  const COLS = [
    { key: 'no', label: 'No', num: true, get: a => Number(a.no) || a.id + 1 },
    { key: 'jenis', label: 'Jenis BMN', get: a => a.jenis, show: a => esc(fmt.jenis(a.jenis)) },
    { key: 'kode', label: 'Kode / NUP', get: a => a.kodeBarang + '.' + a.nup.padStart(5, '0'), show: a => `${esc(a.kodeBarang)}<br>NUP ${esc(a.nup)}` },
    { key: 'nama', label: 'Nama barang', get: a => a.nama.toLowerCase(), show: a => `${esc(a.nama)}${a.merk || a.tipe ? `<br><span class="note">${esc(merkTipe(a.merk, a.tipe))}</span>` : ''}` },
    { key: 'kondisi', label: 'Kondisi', get: a => a.kondisi, show: a => kondisiHtml(a.kondisi) },
    { key: 'tahun', label: 'Tahun', num: true, get: a => a.tahun || 0, show: a => a.tahun || '–' },
    { key: 'perolehan', label: 'Nilai perolehan', num: true, get: a => a.nilaiPerolehan, show: a => fmt.rp(a.nilaiPerolehan) },
    { key: 'buku', label: 'Nilai buku', num: true, get: a => a.nilaiBuku, show: a => fmt.rp(a.nilaiBuku) },
    { key: 'lokasi', label: 'Lokasi ruang', get: a => a.lokasi.toLowerCase(), show: a => esc(a.lokasi) },
    { key: 'flags', label: 'Penanda', get: a => a.flags.length, show: a => flagBadges(a) }
  ];

  function initDetailControls() {
    const s = state.stats;
    const jenisOpts = s.byJenis.slice().sort((a, b) => b.n - a.n).map(j => `<option value="${esc(j.jenis)}">${esc(j.label)} (${fmt.n(j.n)})</option>`).join('');
    $('fJenis').innerHTML = `<option value="">Semua jenis BMN</option>${jenisOpts}`;
    $('fKondisi').innerHTML = `<option value="">Semua kondisi</option>` + Object.keys(s.kondisi).filter(k => s.kondisi[k]).map(k => `<option value="${esc(k)}">${esc(k)} (${fmt.n(s.kondisi[k])})</option>`).join('');
    const groups = [['hapus', 'Penghapusan & kondisi'], ['data', 'Kelengkapan data'], ['umur', 'Umur & penyusutan'], ['guna', 'Penggunaan']];
    $('fFlag').innerHTML = `<option value="">Semua penanda</option><option value="__any">Punya penanda apa pun</option>` + groups.map(([g, gl]) => {
      const opts = s.checks.filter(c => c.group === g && c.count > 0).map(c => `<option value="${c.id}">${esc(c.label)} (${fmt.n(c.count)})</option>`).join('');
      return opts ? `<optgroup label="${gl}">${opts}</optgroup>` : '';
    }).join('');

    if (!initDetailControls.bound) {
      initDetailControls.bound = true;
      let to = null;
      $('fCari').addEventListener('input', e => { clearTimeout(to); to = setTimeout(() => { state.filter.q = A.util.normText(e.target.value); state.page = 0; renderDetail(); }, 150); });
      $('fJenis').addEventListener('change', e => { state.filter.jenis = e.target.value; state.page = 0; renderDetail(); });
      $('fKondisi').addEventListener('change', e => { state.filter.kondisi = e.target.value; state.page = 0; renderDetail(); });
      $('fFlag').addEventListener('change', e => { state.filter.flag = e.target.value; state.page = 0; renderDetail(); });
      $('btnReset').addEventListener('click', () => { resetFilter(); renderDetail(); });
      $('btnCsv').addEventListener('click', exportCsv);
    }
  }
  function resetFilter() {
    state.filter = { q: '', jenis: '', kondisi: '', flag: '', ids: null, idsLabel: '' };
    state.page = 0;
    $('fCari').value = ''; $('fJenis').value = ''; $('fKondisi').value = ''; $('fFlag').value = '';
  }
  function goDetail(flag) {
    resetFilter();
    state.filter.flag = flag;
    $('fFlag').value = flag;
    showTab('detail');
    window.scrollTo({ top: 0 });
  }

  function filtered() {
    const f = state.filter;
    const words = f.q ? f.q.split(' ') : [];
    let list = state.assets.filter(a =>
      (!f.ids || f.ids.has(a.id)) &&
      words.every(w => a._hay.includes(w) || a._hayRapat.includes(w)) &&
      (!f.jenis || a.jenis === f.jenis) &&
      (!f.kondisi || a.kondisi === f.kondisi) &&
      (!f.flag || (f.flag === '__any' ? a.flags.length > 0 : a.flags.includes(f.flag))));
    const col = COLS.find(c => c.key === state.sort.key) || COLS[0];
    const d = state.sort.dir;
    list = list.slice().sort((x, y) => {
      const a = col.get(x), b = col.get(y);
      return (a < b ? -1 : a > b ? 1 : 0) * d || (x.id - y.id);
    });
    return list;
  }

  function renderDetail() {
    const list = filtered();
    const pages = Math.max(1, Math.ceil(list.length / state.pageSize));
    state.page = Math.min(state.page, pages - 1);
    const start = state.page * state.pageSize;
    const rows = list.slice(start, start + state.pageSize);
    const f = state.filter;
    const flagInfo = f.flag && f.flag !== '__any' ? A.FLAG[f.flag] : null;
    const totalNilai = list.reduce((s, a) => s + a.nilaiPerolehan, 0);

    $('tblDetail').innerHTML = `
      ${f.ids ? `<div class="active-filter"><b>Hasil cari cepat:</b><span>“${esc(f.idsLabel)}” (${fmt.n(f.ids.size)} aset)</span><button class="linkbtn" type="button" id="clrIds">Hapus filter</button></div>` : ''}
      ${flagInfo ? `<div class="active-filter">${C.sevChip(flagInfo.sev)}<b>${esc(flagInfo.label)}</b><span>${esc(flagInfo.desc)}</span><button class="linkbtn" type="button" id="clrFlag">Hapus filter</button></div>` : ''}
      <div class="filters"><span class="info">${fmt.n(list.length)} aset · nilai perolehan ${fmt.rpc(totalNilai)}</span></div>
      <div class="tbl-wrap"><table class="resp">
        <thead><tr>${COLS.map(c => `<th class="sortable${c.num ? ' num' : ''}" data-key="${c.key}" tabindex="0" aria-sort="${state.sort.key === c.key ? (state.sort.dir > 0 ? 'ascending' : 'descending') : 'none'}">${esc(c.label)}${state.sort.key === c.key ? ` <span class="arrow">${state.sort.dir > 0 ? '▲' : '▼'}</span>` : ''}</th>`).join('')}</tr></thead>
        <tbody>${rows.length ? rows.map(a => `<tr class="click" data-id="${a.id}" tabindex="0">${COLS.map(c => `<td data-label="${esc(c.label)}"${c.num ? ' class="num"' : ''}><div>${c.show ? c.show(a) : esc(c.get(a))}</div></td>`).join('')}</tr>`).join('') : `<tr><td colspan="${COLS.length}">Tidak ada aset yang cocok dengan filter.</td></tr>`}</tbody>
      </table></div>
      <div class="pager">
        <span>Menampilkan ${list.length ? fmt.n(start + 1) : 0}–${fmt.n(start + rows.length)} dari ${fmt.n(list.length)}</span>
        <button class="btn sm" type="button" id="pgPrev" ${state.page === 0 ? 'disabled' : ''}>‹ Sebelumnya</button>
        <span>Hal. ${state.page + 1} / ${pages}</span>
        <button class="btn sm" type="button" id="pgNext" ${state.page >= pages - 1 ? 'disabled' : ''}>Berikutnya ›</button>
      </div>`;

    const clrIds = $('clrIds');
    if (clrIds) clrIds.addEventListener('click', () => { state.filter.ids = null; state.filter.idsLabel = ''; renderDetail(); });
    const clr = $('clrFlag');
    if (clr) clr.addEventListener('click', () => { state.filter.flag = ''; $('fFlag').value = ''; renderDetail(); });
    $('pgPrev').addEventListener('click', () => { state.page--; renderDetail(); });
    $('pgNext').addEventListener('click', () => { state.page++; renderDetail(); });
    $('tblDetail').querySelectorAll('th.sortable').forEach(thEl => {
      const act = () => {
        const k = thEl.dataset.key;
        state.sort = { key: k, dir: state.sort.key === k ? -state.sort.dir : (COLS.find(c => c.key === k).num ? -1 : 1) };
        renderDetail();
      };
      thEl.addEventListener('click', act);
      thEl.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); act(); } });
    });
    bindRowClicks($('tblDetail'));
  }

  // ---------------- Dialog detail aset ----------------
  function bindRowClicks(root) {
    root.querySelectorAll('tr.click').forEach(tr => {
      const open = () => openAsset(state.assets[Number(tr.dataset.id)]);
      tr.addEventListener('click', open);
      tr.addEventListener('keydown', e => { if (e.key === 'Enter') open(); });
    });
  }
  function openAsset(a) {
    if (!a) return;
    $('dlgTitle').textContent = a.nama;
    $('dlgSub').textContent = `${fmt.jenis(a.jenis)} · ${a.kodeBarang} · NUP ${a.nup}`;
    const filled = state.headers.filter(h => h && A.util.str(a.raw[h]) !== '');
    const empty = state.headers.filter(h => h && A.util.str(a.raw[h]) === '');
    const flags = a.flags.length ? `<h3>Penanda</h3><ul class="alerts">${a.flags.map(f => `<li><span>${C.sevChip(A.FLAG[f].sev)}</span><div><div class="t">${esc(A.FLAG[f].label)}</div><div class="d">${esc(A.FLAG[f].desc)}</div></div><div></div></li>`).join('')}</ul>` : '';
    $('dlgBody').innerHTML = `${flags}<h3>Data lengkap</h3><dl class="kv">${filled.map(h => {
      const v = a.raw[h];
      const isMoney = /^nilai/i.test(h) && typeof v === 'number';
      return `<dt>${esc(h)}</dt><dd>${isMoney ? fmt.rp(v) : esc(v)}</dd>`;
    }).join('')}</dl>${empty.length ? `<p class="note-box">${empty.length} kolom kosong tidak ditampilkan: ${esc(empty.join(', '))}.</p>` : ''}`;
    const dlg = $('dlg');
    if (typeof dlg.showModal === 'function') dlg.showModal(); else dlg.setAttribute('open', '');
  }
  function initDialog() {
    const dlg = $('dlg');
    $('dlgClose').addEventListener('click', () => dlg.close());
    dlg.addEventListener('click', e => { if (e.target === dlg) dlg.close(); });
  }

  // ---------------- Ekspor CSV ----------------
  function exportCsv() {
    const list = filtered();
    const headers = state.headers.filter(Boolean);
    const sep = ';';
    const q = v => {
      const s = v == null ? '' : String(v);
      return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const lines = [headers.concat(['Penanda']).map(q).join(sep)];
    list.forEach(a => lines.push(headers.map(h => q(a.raw[h])).concat([q(a.flags.map(f => A.FLAG[f].label).join(', '))]).join(sep)));
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const d = new Date();
    const p = n => String(n).padStart(2, '0');
    const tag = state.filter.flag && state.filter.flag !== '__any' ? '-' + state.filter.flag : '';
    const link = document.createElement('a');
    link.href = url;
    link.download = `aset${tag}-${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ---------------- Mulai ----------------
  initTheme();
  initLoader();
  initTabs();
  initDialog();
})();
