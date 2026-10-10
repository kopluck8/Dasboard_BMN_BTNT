/*
 * Normalisasi data Daftar Aset + aturan kesimpulan & peringatan.
 * Fungsi murni (tanpa DOM) supaya bisa diuji di Node: lihat tests/cek-data.js
 */
(function (root) {
  'use strict';

  // ---------- util ----------
  const str = v => (v == null ? '' : String(v).replace(/\s+/g, ' ').trim());
  const up = v => str(v).toUpperCase();
  const yes = v => /^ya$/i.test(str(v));
  function num(v) {
    if (typeof v === 'number') return v;
    const s = str(v);
    if (!s || s === '-') return 0;
    let n = Number(s);
    if (Number.isFinite(n)) return n;
    n = Number(s.replace(/\./g, '').replace(',', '.'));
    return Number.isFinite(n) ? n : 0;
  }
  const esc = v => str(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const kosong = v => { const s = str(v); return s === '' || s === '-' || s === '0'; };
  const normKey = v => up(v).replace(/[\s.\-]/g, '');

  const nfInt = new Intl.NumberFormat('id-ID', { maximumFractionDigits: 0 });
  const fmt = {
    n: x => nfInt.format(x || 0),
    rp: x => 'Rp' + nfInt.format(Math.round(x || 0)),
    rpc(x) {
      const a = Math.abs(x || 0);
      const f = (v, d) => new Intl.NumberFormat('id-ID', { maximumFractionDigits: d, minimumFractionDigits: 0 }).format(v);
      if (a >= 1e12) return 'Rp' + f(x / 1e12, 2) + ' T';
      if (a >= 1e9) return 'Rp' + f(x / 1e9, 2) + ' M';
      if (a >= 1e6) return 'Rp' + f(x / 1e6, 1) + ' jt';
      if (a >= 1e3) return 'Rp' + f(x / 1e3, 0) + ' rb';
      return 'Rp' + f(x, 0);
    },
    pct: (x, d = 1) => new Intl.NumberFormat('id-ID', { maximumFractionDigits: d, minimumFractionDigits: d }).format((x || 0) * 100) + '%',
    tgl(d) {
      if (!d) return '';
      const p = n => String(n).padStart(2, '0');
      return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
    },
    jenis(j) {
      // "MESIN PERALATAN NON TIK" -> "Mesin Peralatan Non TIK"
      return str(j).toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).replace(/\bTik\b/g, 'TIK').replace(/\bDan\b/g, 'dan');
    }
  };

  // ---------- normalisasi ----------
  function normalize(book, cfg) {
    const rows = book.rows;
    const headers = rows[book.headerIndex].map(str);
    const toDate = root.XlsxReader.excelDate;
    const acuan = cfg.tanggalAcuan ? new Date(cfg.tanggalAcuan) : new Date();
    const assets = [];
    for (let r = book.headerIndex + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.length) continue;
      const raw = {};
      headers.forEach((h, i) => { if (h) raw[h] = row[i] == null ? '' : row[i]; });
      if (!str(raw['Jenis BMN'])) continue;

      const dates = {};
      headers.forEach(h => {
        if (/^tanggal/i.test(h)) {
          const v = raw[h];
          let d = typeof v === 'number' ? toDate(v) : (str(v) ? new Date(str(v)) : null);
          if (d && (isNaN(d) || d.getUTCFullYear() >= cfg.tanggalKosongTahun)) d = null;
          dates[h] = d;
          raw[h] = d ? fmt.tgl(d) : '';
        }
      });
      const tgl = dates['Tanggal Perolehan'] || null;
      const umur = tgl ? (acuan - tgl) / (365.25 * 86400000) : null;

      assets.push({
        id: assets.length,
        raw,
        no: raw['No'],
        jenis: up(raw['Jenis BMN']),
        kodeBarang: str(raw['Kode Barang']),
        nup: str(raw['NUP']),
        nama: str(raw['Nama Barang']),
        merk: str(raw['Merk']),
        tipe: str(raw['Tipe']),
        kondisi: str(raw['Kondisi']) || '(kosong)',
        intra: /^intra/i.test(str(raw['Intra / Extra'])),
        hentiGuna: yes(raw['Henti Guna']),
        usulHapus: yes(raw['Usul Hapus']),
        statusGuna: str(raw['Status Penggunaan']),
        lokasi: str(raw['Lokasi Ruang']),
        noPSP: str(raw['No PSP']),
        foto: num(raw['Jumlah Foto']),
        sertifikasi: str(raw['Status Sertifikasi']),
        noPolisi: str(raw['No Polisi']),
        noDokumen: str(raw['No Dokumen']),
        noSTNK: str(raw['No STNK']),
        kabkota: str(raw['Kab/Kota']),
        alamat: str(raw['Alamat']),
        tglPerolehan: tgl,
        tahun: tgl ? tgl.getUTCFullYear() : null,
        umur,
        nilaiPerolehan: num(raw['Nilai Perolehan']),
        nilaiPenyusutan: num(raw['Nilai Penyusutan']),
        nilaiBuku: num(raw['Nilai Buku'])
      });
    }
    return { headers, assets, acuan };
  }

  // ---------- aturan / penanda ----------
  // sev: critical | serious | warning | info
  const inList = (list, a) => list.includes(a.jenis);
  const FLAGS = [
    // Kandidat penghapusan
    { id: 'hapus', need: ['Kondisi'], group: 'hapus', sev: 'serious', label: 'Kandidat penghapusan',
      desc: 'Rusak Berat, Henti Guna, atau sudah Usul Hapus.',
      test: a => a.kondisi === 'Rusak Berat' || a.hentiGuna || a.usulHapus },
    { id: 'rb_belum_usul', need: ['Kondisi', 'Usul Hapus'], group: 'hapus', sev: 'critical', label: 'Rusak Berat, belum Usul Hapus',
      desc: 'Kondisi Rusak Berat tetapi belum diusulkan penghapusan.',
      test: a => a.kondisi === 'Rusak Berat' && !a.usulHapus },
    { id: 'henti_belum_usul', need: ['Henti Guna', 'Usul Hapus'], group: 'hapus', sev: 'warning', label: 'Henti Guna, belum Usul Hapus',
      desc: 'Sudah dihentikan penggunaannya tetapi belum diusulkan penghapusan.',
      test: a => a.hentiGuna && !a.usulHapus },
    { id: 'usul_tapi_baik', need: ['Kondisi', 'Usul Hapus'], group: 'hapus', sev: 'warning', label: 'Usul Hapus, kondisi Baik',
      desc: 'Diusulkan hapus padahal kondisi tercatat Baik. Periksa konsistensi data kondisi.',
      test: a => a.usulHapus && a.kondisi === 'Baik' },
    { id: 'rusak_ringan', need: ['Kondisi'], group: 'hapus', sev: 'warning', label: 'Rusak Ringan',
      desc: 'Perlu pemeliharaan agar tidak menjadi Rusak Berat.',
      test: a => a.kondisi === 'Rusak Ringan' },

    // Kelengkapan data
    { id: 'tanpa_psp', need: ['No PSP'], group: 'data', sev: 'serious', label: 'Belum ada No PSP',
      desc: 'Penetapan Status Penggunaan belum tercatat.',
      test: a => kosong(a.noPSP) },
    { id: 'tanpa_foto', need: ['Jumlah Foto'], group: 'data', sev: 'warning', label: 'Tanpa foto',
      desc: 'Jumlah foto = 0.',
      test: a => !(a.foto > 0) },
    { id: 'belum_ruang', need: ['Lokasi Ruang'], group: 'data', sev: 'warning', label: 'Peralatan belum berlokasi ruang',
      desc: 'Hanya dicek untuk Mesin Peralatan (TIK & non TIK).',
      scope: (a, c) => inList(c.jenisWajibRuang, a),
      test: a => !a.lokasi || /belum berlokasi/i.test(a.lokasi) },
    { id: 'tanah_sertifikat', need: ['Status Sertifikasi'], group: 'data', sev: 'serious', label: 'Tanah belum bersertipikat penuh',
      desc: 'Hanya dicek untuk jenis Tanah.',
      scope: (a, c) => inList(c.jenisTanah, a),
      test: a => !/bersertipikat seluruh/i.test(a.sertifikasi) },
    { id: 'alamat_kosong', need: ['Kab/Kota'], group: 'data', sev: 'warning', label: 'Tanah/bangunan tanpa Kab/Kota',
      desc: 'Hanya dicek untuk Tanah, Bangunan, Rumah Negara.',
      scope: (a, c) => inList(c.jenisBeralamat, a),
      test: a => kosong(a.kabkota) },
    { id: 'kend_dokumen', need: ['No Dokumen'], group: 'data', sev: 'serious', label: 'Kendaraan tanpa No Dokumen (BPKB)',
      desc: 'Kolom "No Dokumen" kosong atau 0.',
      scope: (a, c) => inList(c.jenisKendaraan, a),
      test: a => kosong(a.noDokumen) },
    { id: 'kend_stnk', need: ['No STNK'], group: 'data', sev: 'warning', label: 'Kendaraan tanpa No STNK',
      desc: 'Kolom "No STNK" kosong.',
      scope: (a, c) => inList(c.jenisKendaraan, a),
      test: a => kosong(a.noSTNK) },
    { id: 'kend_nopol', need: ['No Polisi'], group: 'data', sev: 'warning', label: 'Kendaraan tanpa No Polisi',
      desc: 'Kolom "No Polisi" kosong.',
      scope: (a, c) => inList(c.jenisKendaraan, a),
      test: a => kosong(a.noPolisi) },
    { id: 'kend_duplikat', need: [], group: 'data', sev: 'serious', label: 'Nomor kendaraan ganda',
      desc: 'No Polisi, No STNK, atau No Dokumen yang sama dipakai lebih dari satu kendaraan.',
      scope: (a, c) => inList(c.jenisKendaraan, a),
      test: (a, ctx) => ctx.dupKend.has(a.id) },

    // Umur & penyusutan
    { id: 'habis_susut', need: ['Intra / Extra'], group: 'umur', sev: 'info', label: 'Nilai buku habis (intrakomptabel)',
      desc: 'Intrakomptabel dengan nilai buku Rp0: masa manfaat akuntansi sudah habis.',
      scope: (a, c) => a.intra && !inList(c.jenisTanah, a),
      test: a => a.nilaiPerolehan > 0 && a.nilaiBuku <= 0 },
    { id: 'tua', need: ['Tanggal Perolehan'], group: 'umur', sev: 'info', label: 'Aset tua',
      desc: 'Peralatan/kendaraan berumur ≥ ambang (lihat config.js).',
      scope: (a, c) => !inList(c.jenisTanpaPeremajaan, a),
      test: (a, ctx, c) => a.umur != null && a.umur >= c.umurTuaTahun },

    // Penggunaan
    { id: 'tidak_digunakan', need: ['Status Penggunaan'], group: 'guna', sev: 'serious', label: 'BMN tidak digunakan',
      desc: 'Status Penggunaan "BMN Tidak Digunakan" (idle).',
      test: a => /tidak digunakan/i.test(a.statusGuna) },
    { id: 'satker_lain', need: ['Status Penggunaan'], group: 'guna', sev: 'info', label: 'Digunakan satker lain',
      desc: 'Dipakai satker lain; pastikan ada dasar penggunaan.',
      test: a => /satker lain/i.test(a.statusGuna) }
  ];
  const FLAG = Object.fromEntries(FLAGS.map(f => [f.id, f]));
  const SEV_RANK = { critical: 0, serious: 1, warning: 2, info: 3 };

  function findDuplicateVehicles(assets, cfg) {
    const dup = new Set();
    const keys = ['noPolisi', 'noSTNK', 'noDokumen'];
    keys.forEach(k => {
      const m = new Map();
      assets.filter(a => cfg.jenisKendaraan.includes(a.jenis) && !kosong(a[k])).forEach(a => {
        const key = normKey(a[k]);
        if (!m.has(key)) m.set(key, []);
        m.get(key).push(a.id);
      });
      m.forEach(ids => { if (ids.length > 1) ids.forEach(i => dup.add(i)); });
    });
    return dup;
  }

  const sum = (arr, k) => arr.reduce((s, a) => s + (a[k] || 0), 0);
  function groupBy(arr, keyFn) {
    const m = new Map();
    arr.forEach(a => { const k = keyFn(a); if (!m.has(k)) m.set(k, []); m.get(k).push(a); });
    return m;
  }

  // ---------- analisis ----------
  function analyze(assets, cfg, headers) {
    const ctx = { dupKend: findDuplicateVehicles(assets, cfg) };
    const hs = new Set(headers || []);
    // Aturan yang kolomnya tidak ada di file tidak dijalankan (hindari alarm palsu).
    const avail = f => !headers || (f.need || []).every(c => hs.has(c));
    const active = FLAGS.filter(avail);
    assets.forEach(a => {
      a.flags = active.filter(f => (!f.scope || f.scope(a, cfg)) && f.test(a, ctx, cfg)).map(f => f.id);
    });

    const checks = FLAGS.map(f => {
      const scoped = assets.filter(a => !f.scope || f.scope(a, cfg));
      const hit = scoped.filter(a => a.flags.includes(f.id));
      return { id: f.id, group: f.group, sev: f.sev, label: f.label, desc: f.desc, applicable: scoped.length, count: hit.length, nilai: sum(hit, 'nilaiPerolehan'), unavailable: !avail(f), need: f.need || [] };
    });
    const check = Object.fromEntries(checks.map(c => [c.id, c]));

    const total = {
      n: assets.length,
      perolehan: sum(assets, 'nilaiPerolehan'),
      penyusutan: sum(assets, 'nilaiPenyusutan'),
      buku: sum(assets, 'nilaiBuku'),
      intra: assets.filter(a => a.intra).length
    };

    const byJenis = [...groupBy(assets, a => a.jenis)].map(([jenis, list]) => ({
      jenis, label: fmt.jenis(jenis), n: list.length,
      perolehan: sum(list, 'nilaiPerolehan'), buku: sum(list, 'nilaiBuku'), penyusutan: sum(list, 'nilaiPenyusutan'),
      rb: list.filter(a => a.kondisi === 'Rusak Berat').length,
      rr: list.filter(a => a.kondisi === 'Rusak Ringan').length
    })).sort((x, y) => y.perolehan - x.perolehan);

    const kondisi = { 'Baik': 0, 'Rusak Ringan': 0, 'Rusak Berat': 0 };
    assets.forEach(a => { kondisi[a.kondisi] = (kondisi[a.kondisi] || 0) + 1; });

    const withYear = assets.filter(a => a.tahun);
    const years = withYear.map(a => a.tahun);
    const byTahun = [];
    if (years.length) {
      const y0 = Math.min(...years), y1 = Math.max(...years);
      const g = groupBy(withYear, a => a.tahun);
      for (let y = y0; y <= y1; y++) {
        const list = g.get(y) || [];
        byTahun.push({ tahun: y, n: list.length, perolehan: sum(list, 'nilaiPerolehan') });
      }
    }

    const hapusList = assets.filter(a => a.flags.includes('hapus'));
    const hapus = {
      n: hapusList.length,
      perolehan: sum(hapusList, 'nilaiPerolehan'),
      buku: sum(hapusList, 'nilaiBuku'),
      byJenis: [...groupBy(hapusList, a => a.jenis)].map(([j, l]) => ({ label: fmt.jenis(j), n: l.length, perolehan: sum(l, 'nilaiPerolehan') })).sort((x, y) => y.n - x.n),
      byNama: [...groupBy(hapusList, a => a.nama)].map(([j, l]) => ({ label: j, n: l.length, perolehan: sum(l, 'nilaiPerolehan') })).sort((x, y) => y.n - x.n),
      byTahun: [...groupBy(hapusList, a => a.tahun || 0)].map(([t, l]) => ({ tahun: t, n: l.length })).sort((x, y) => y.n - x.n)
    };

    const kendaraan = assets.filter(a => cfg.jenisKendaraan.includes(a.jenis))
      .sort((x, y) => (x.tahun || 0) - (y.tahun || 0));

    const umurPerJenis = [...groupBy(assets.filter(a => !cfg.jenisTanpaPeremajaan.includes(a.jenis)), a => a.jenis)]
      .map(([j, l]) => {
        const u = l.filter(a => a.umur != null).map(a => a.umur);
        return {
          jenis: j, label: fmt.jenis(j), n: l.length,
          rata: u.length ? u.reduce((s, x) => s + x, 0) / u.length : null,
          tertua: u.length ? Math.max(...u) : null,
          habis: l.filter(a => a.flags.includes('habis_susut')).length,
          intra: l.filter(a => a.intra).length,
          tua: l.filter(a => a.flags.includes('tua')).length
        };
      }).sort((x, y) => y.n - x.n);

    const stats = { total, byJenis, kondisi, byTahun, hapus, kendaraan, umurPerJenis, checks, check };
    stats.kesimpulan = buildKesimpulan(assets, stats, cfg);
    stats.perhatian = checks
      .filter(c => c.count > 0 && c.sev !== 'info' && c.id !== 'hapus')
      .concat(check.hapus.count ? [check.hapus] : [])
      .sort((x, y) => SEV_RANK[x.sev] - SEV_RANK[y.sev] || y.count - x.count);
    return stats;
  }

  function buildKesimpulan(assets, s, cfg) {
    const out = [];
    const t = s.total;
    if (!t.n) return out;
    const top = s.byJenis.slice().sort((a, b) => b.buku - a.buku)[0];
    out.push(`Terdapat <b>${fmt.n(t.n)} aset</b> dengan nilai perolehan <b>${fmt.rpc(t.perolehan)}</b> dan nilai buku <b>${fmt.rpc(t.buku)}</b> (${fmt.pct(t.buku / t.perolehan)} dari nilai perolehan). ` +
      `Nilai buku terbesar ada pada <b>${esc(top.label)}</b> (${fmt.pct(top.buku / t.buku)} dari total nilai buku).`);

    const rb = s.kondisi['Rusak Berat'] || 0, rr = s.kondisi['Rusak Ringan'] || 0;
    const baikPct = (s.kondisi['Baik'] || 0) / t.n;
    let k = `Kondisi Baik ${fmt.pct(baikPct)}. `;
    if (rb) {
      const belum = s.check.rb_belum_usul.count;
      k += `<b>${fmt.n(rb)} aset Rusak Berat</b>` + (belum ? `, ${fmt.n(belum)} di antaranya <b>belum diusulkan hapus</b>.` : ', seluruhnya sudah diusulkan hapus.');
      const topNama = s.hapus.byNama.slice(0, 3).map(x => `${esc(x.label)} (${x.n})`).join(', ');
      if (topNama) k += ` Terbanyak: ${topNama}.`;
    }
    if (rr) k += ` ${fmt.n(rr)} aset Rusak Ringan perlu pemeliharaan.`;
    out.push(k);

    const hs = s.check.habis_susut;
    if (hs.count) {
      const aktif = assets.filter(a => a.flags.includes('habis_susut') && !a.flags.includes('hapus')).length;
      out.push(`<b>${fmt.n(hs.count)} dari ${fmt.n(hs.applicable)} aset intrakomptabel</b> (${fmt.pct(hs.count / hs.applicable)}) nilai bukunya sudah Rp0, dan ${fmt.n(aktif)} di antaranya bukan kandidat penghapusan (masih aktif dipakai). Ini sinyal kebutuhan peremajaan dan perencanaan anggaran pengganti.`);
    }

    if (s.byTahun.length) {
      const peak = s.byTahun.slice().sort((a, b) => b.n - a.n)[0];
      const rbPeak = s.hapus.byTahun[0];
      let p = `Pengadaan terbanyak terjadi pada tahun <b>${peak.tahun}</b> (${fmt.n(peak.n)} unit, ${fmt.pct(peak.n / t.n)}).`;
      if (rbPeak && rbPeak.tahun && s.hapus.n) {
        p += rbPeak.tahun === peak.tahun
          ? ` Kandidat penghapusan juga paling banyak berasal dari perolehan ${rbPeak.tahun} (${fmt.n(rbPeak.n)} unit), sehingga aset seangkatan yang masih Baik perlu dipantau.`
          : ` Kandidat penghapusan paling banyak berasal dari perolehan ${rbPeak.tahun} (${fmt.n(rbPeak.n)} unit).`;
      }
      out.push(p);
    }

    const td = s.check.tidak_digunakan;
    if (td.count) {
      const tanah = assets.filter(a => a.flags.includes('tidak_digunakan') && cfg.jenisTanah.includes(a.jenis));
      out.push(`<b>${fmt.n(td.count)} BMN tidak digunakan</b>` + (tanah.length ? `, termasuk ${tanah.length} bidang tanah senilai ${fmt.rpc(sum(tanah, 'nilaiBuku'))}` : '') + '. Perlu keputusan: digunakan kembali, dimanfaatkan, dialihkan, atau diserahkan.');
    }

    const dataIssues = s.checks.filter(c => c.group === 'data' && c.count > 0);
    const missing = s.checks.filter(c => c.unavailable);
    if (dataIssues.length) {
      out.push('Kelengkapan data yang perlu dilengkapi: ' + dataIssues.map(c => `${c.label} (${fmt.n(c.count)})`).join('; ') + '.');
    } else {
      out.push('Semua pemeriksaan kelengkapan data lolos.');
    }
    if (missing.length) out.push(`${missing.length} pemeriksaan tidak dijalankan karena kolomnya tidak ada di file: ${missing.map(c => esc(c.label)).join(', ')}.`);
    return out;
  }

  // ---------- cari cepat ----------
  // Huruf kecil, tanpa aksen, tanda baca jadi spasi: "Pompa-Punggung" == "pompa punggung".
  const normText = v => str(v).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();

  /** Teks yang dicari: identitas barang saja (bukan lokasi), supaya jumlah tidak tercampur. */
  function itemText(a) {
    if (a._item == null) {
      a._item = ' ' + normText([a.nama, a.merk, a.tipe, a.raw['Nama'], a.kodeBarang, a.jenis].join(' ')) + ' ';
      a._itemRapat = a._item.replace(/ /g, ''); // "Lap Top" juga cocok dengan "laptop"
    }
    return a._item;
  }

  /**
   * Semua kata pada kueri harus ada (urutan bebas).
   * Kata berupa angka dicocokkan sebagai awalan kata (kode barang, ukuran).
   * Kata huruf juga dicocokkan tanpa spasi, jadi "laptop" menemukan "Lap Top".
   */
  function cocok(assets, query) {
    const tokens = normText(query).split(' ').filter(Boolean);
    if (!tokens.length) return [];
    return assets.filter(a => {
      const t = itemText(a);
      return tokens.every(tok => (/^\d+$/.test(tok) ? t.includes(' ' + tok) : (t.includes(tok) || a._itemRapat.includes(tok))));
    });
  }

  function ringkasCari(assets, query, cfg) {
    const list = cocok(assets, query);
    const kondisi = { 'Baik': 0, 'Rusak Ringan': 0, 'Rusak Berat': 0 };
    list.forEach(a => { kondisi[a.kondisi] = (kondisi[a.kondisi] || 0) + 1; });
    const tahun = list.map(a => a.tahun).filter(Boolean);
    const umur = list.map(a => a.umur).filter(u => u != null);

    const grp = new Map();
    list.forEach(a => {
      const key = normText([a.nama, a.merk, a.tipe].join('|'));
      if (!grp.has(key)) grp.set(key, { nama: a.nama, merk: a.merk, tipe: a.tipe, n: 0, baik: 0, rr: 0, rb: 0, perolehan: 0 });
      const g = grp.get(key);
      g.n++; g.perolehan += a.nilaiPerolehan;
      if (a.kondisi === 'Baik') g.baik++; else if (a.kondisi === 'Rusak Ringan') g.rr++; else if (a.kondisi === 'Rusak Berat') g.rb++;
    });
    const countBy = fn => [...groupBy(list, fn)].map(([label, l]) => ({ label, n: l.length })).sort((x, y) => y.n - x.n);

    const flagCount = {};
    list.forEach(a => a.flags.forEach(f => { flagCount[f] = (flagCount[f] || 0) + 1; }));
    const penanda = Object.keys(flagCount)
      .filter(f => FLAG[f].sev !== 'info')
      .map(f => ({ id: f, label: FLAG[f].label, sev: FLAG[f].sev, n: flagCount[f] }))
      .sort((x, y) => SEV_RANK[x.sev] - SEV_RANK[y.sev] || y.n - x.n);

    return {
      query: str(query), ids: list.map(a => a.id), n: list.length,
      kondisi,
      perolehan: sum(list, 'nilaiPerolehan'), buku: sum(list, 'nilaiBuku'),
      tahunMin: tahun.length ? Math.min(...tahun) : null, tahunMax: tahun.length ? Math.max(...tahun) : null,
      umurRata: umur.length ? umur.reduce((s, x) => s + x, 0) / umur.length : null,
      habis: list.filter(a => a.flags.includes('habis_susut')).length,
      hapus: list.filter(a => a.flags.includes('hapus')).length,
      grup: [...grp.values()].sort((x, y) => y.n - x.n),
      jenis: countBy(a => fmt.jenis(a.jenis)),
      lokasi: countBy(a => a.lokasi || '(kosong)'),
      guna: countBy(a => a.statusGuna || '(kosong)'),
      penanda
    };
  }

  root.AsetBMN = { normalize, analyze, ringkasCari, cocok, FLAGS, FLAG, SEV_RANK, fmt, util: { str, num, kosong, esc, normText } };
})(typeof window !== 'undefined' ? window : globalThis);
