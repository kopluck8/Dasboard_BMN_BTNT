/*
 * Uji cepat di komputer sendiri (Node 18+), tanpa browser:
 *   node tests/cek-data.js path/ke/daftar-aset.xlsx
 * Mencetak ringkasan angka yang sama dengan yang tampil di dashboard.
 * File .xlsx TIDAK perlu (dan jangan) di-commit ke repo.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const base = path.join(__dirname, '..', 'assets', 'js');
['config.js', 'xlsx-reader.js', 'analisis.js'].forEach(f => vm.runInThisContext(fs.readFileSync(path.join(base, f), 'utf8'), { filename: f }));

const file = process.argv[2];
if (!file) { console.error('Pakai: node tests/cek-data.js <file.xlsx>'); process.exit(1); }

(async () => {
  const buf = fs.readFileSync(file);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const cfg = globalThis.ASET_CONFIG;
  const book = await globalThis.XlsxReader.readWorkbook(ab, cfg.requiredHeaders);
  const { assets, headers } = globalThis.AsetBMN.normalize(book, cfg);
  const s = globalThis.AsetBMN.analyze(assets, cfg, headers);
  const out = {
    sheet: book.sheetName,
    total: s.total,
    kondisi: s.kondisi,
    hapus: { n: s.hapus.n, perolehan: s.hapus.perolehan, buku: s.hapus.buku },
    checks: Object.fromEntries(s.checks.map(c => [c.id, `${c.count}/${c.applicable}`])),
    tahunPuncak: s.byTahun.slice().sort((a, b) => b.n - a.n)[0],
    kendaraan: s.kendaraan.length
  };
  console.log(JSON.stringify(out, null, 2));
  console.log('\nKesimpulan:');
  s.kesimpulan.forEach((k, i) => console.log(`${i + 1}. ${k.replace(/<[^>]+>/g, '')}`));
})().catch(e => { console.error(e); process.exit(1); });
