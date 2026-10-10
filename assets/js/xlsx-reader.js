/*
 * Pembaca .xlsx minimal tanpa library eksternal.
 * - Membuka arsip ZIP lewat central directory.
 * - Dekompresi memakai DecompressionStream('deflate-raw') bawaan browser.
 * - Mengurai sharedStrings + worksheet dengan regex (cukup untuk ekspor SIMAN/SAKTI).
 * Semua proses terjadi di memori browser; tidak ada data yang dikirim ke mana pun.
 */
(function (root) {
  'use strict';

  const td = new TextDecoder('utf-8');

  function readU16(v, o) { return v.getUint16(o, true); }
  function readU32(v, o) { return v.getUint32(o, true); }

  /** Baca daftar entri ZIP dari End Of Central Directory. */
  function listZipEntries(buf) {
    const v = new DataView(buf);
    const bytes = new Uint8Array(buf);
    let eocd = -1;
    const min = Math.max(0, buf.byteLength - 65557);
    for (let i = buf.byteLength - 22; i >= min; i--) {
      if (readU32(v, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('File bukan .xlsx yang valid (struktur ZIP tidak ditemukan).');
    const count = readU16(v, eocd + 10);
    let p = readU32(v, eocd + 16);
    const entries = {};
    for (let n = 0; n < count; n++) {
      if (readU32(v, p) !== 0x02014b50) throw new Error('Central directory ZIP rusak.');
      const method = readU16(v, p + 10);
      const csize = readU32(v, p + 20);
      const nameLen = readU16(v, p + 28);
      const extraLen = readU16(v, p + 30);
      const commentLen = readU16(v, p + 32);
      const localOffset = readU32(v, p + 42);
      const name = td.decode(bytes.subarray(p + 46, p + 46 + nameLen));
      entries[name] = { method, csize, localOffset };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  async function inflateRaw(u8) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('Browser ini belum mendukung DecompressionStream. Gunakan Chrome/Edge/Firefox/Safari versi terbaru.');
    }
    const stream = new Blob([u8]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function readEntry(buf, entry) {
    const v = new DataView(buf);
    const o = entry.localOffset;
    if (readU32(v, o) !== 0x04034b50) throw new Error('Header file ZIP rusak.');
    const start = o + 30 + readU16(v, o + 26) + readU16(v, o + 28);
    const data = new Uint8Array(buf, start, entry.csize);
    if (entry.method === 0) return td.decode(data);
    if (entry.method === 8) return td.decode(await inflateRaw(data));
    throw new Error('Metode kompresi ZIP tidak didukung: ' + entry.method);
  }

  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function decodeXml(s) {
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
      if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return ENT[e];
    }).replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  }

  function textOf(xml) {
    // Gabungkan semua <t> (termasuk rich-text run), abaikan teks fonetik <rPh>.
    const clean = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let out = '';
    const re = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g;
    let m;
    while ((m = re.exec(clean))) out += m[1] ? decodeXml(m[1]) : '';
    return out;
  }

  function parseSharedStrings(xml) {
    const out = [];
    const re = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
    let m;
    while ((m = re.exec(xml))) out.push(m[1] ? textOf(m[1]) : '');
    return out;
  }

  function colIndex(ref) {
    let n = 0;
    for (let i = 0; i < ref.length; i++) {
      const c = ref.charCodeAt(i);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  /** Mengubah worksheet XML menjadi array baris (array nilai). */
  function parseSheet(xml, sst) {
    const rows = [];
    const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let rm;
    while ((rm = rowRe.exec(xml))) {
      const rAttr = /\br="(\d+)"/.exec(rm[1]);
      const rIdx = rAttr ? parseInt(rAttr[1], 10) - 1 : rows.length;
      const row = [];
      let cm, auto = 0;
      cellRe.lastIndex = 0;
      while ((cm = cellRe.exec(rm[2]))) {
        const attrs = cm[1];
        const inner = cm[2] || '';
        const refM = /\br="([A-Z]+)\d*"/.exec(attrs);
        const ci = refM ? colIndex(refM[1]) : auto;
        auto = ci + 1;
        const tM = /\bt="(\w+)"/.exec(attrs);
        const t = tM ? tM[1] : 'n';
        let val = null;
        if (t === 'inlineStr') {
          val = textOf(inner);
        } else {
          const vM = /<v>([\s\S]*?)<\/v>/.exec(inner);
          if (vM) {
            const raw = vM[1];
            if (t === 's') val = sst[parseInt(raw, 10)] ?? '';
            else if (t === 'b') val = raw === '1';
            else if (t === 'str' || t === 'e') val = decodeXml(raw);
            else { const num = Number(raw); val = Number.isFinite(num) ? num : decodeXml(raw); }
          }
        }
        row[ci] = val;
      }
      rows[rIdx] = row;
    }
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  /**
   * Baca workbook dan kembalikan { sheetName, rows } untuk sheet pertama
   * yang barisnya memuat semua kolom wajib (requiredHeaders).
   */
  async function readWorkbook(arrayBuffer, requiredHeaders) {
    const entries = listZipEntries(arrayBuffer);
    const sst = entries['xl/sharedStrings.xml'] ? parseSharedStrings(await readEntry(arrayBuffer, entries['xl/sharedStrings.xml'])) : [];

    // Urutan & nama sheet dari workbook.xml + rels
    const sheets = [];
    if (entries['xl/workbook.xml'] && entries['xl/_rels/workbook.xml.rels']) {
      const wb = await readEntry(arrayBuffer, entries['xl/workbook.xml']);
      const rels = await readEntry(arrayBuffer, entries['xl/_rels/workbook.xml.rels']);
      const relMap = {};
      rels.replace(/<Relationship\b([^>]*)\/?>/g, (m, a) => {
        const id = /\bId="([^"]+)"/.exec(a), tg = /\bTarget="([^"]+)"/.exec(a);
        if (id && tg) relMap[id[1]] = tg[1];
        return m;
      });
      wb.replace(/<sheet\b([^>]*)\/?>/g, (m, a) => {
        const nm = /\bname="([^"]*)"/.exec(a), rid = /\br:id="([^"]+)"/.exec(a);
        if (nm && rid && relMap[rid[1]]) {
          let target = relMap[rid[1]].replace(/^\//, '');
          if (!target.startsWith('xl/')) target = 'xl/' + target;
          sheets.push({ name: decodeXml(nm[1]), path: target });
        }
        return m;
      });
    }
    if (!sheets.length) {
      Object.keys(entries).filter(k => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)).sort()
        .forEach(k => sheets.push({ name: k, path: k }));
    }

    for (const sh of sheets) {
      if (!entries[sh.path]) continue;
      const rows = parseSheet(await readEntry(arrayBuffer, entries[sh.path]), sst);
      const hIdx = rows.findIndex(r => {
        const set = new Set(r.map(x => (x == null ? '' : String(x).trim())));
        return requiredHeaders.every(h => set.has(h));
      });
      if (hIdx >= 0) return { sheetName: sh.name, headerIndex: hIdx, rows };
    }
    throw new Error('Tidak menemukan sheet dengan kolom: ' + requiredHeaders.join(', ') + '. Pastikan file adalah ekspor Daftar Aset (Master Aset).');
  }

  /** Serial tanggal Excel (sistem 1900) ke Date (UTC). */
  function excelDate(serial) {
    if (typeof serial !== 'number' || !Number.isFinite(serial)) return null;
    const ms = Math.round((serial - 25569) * 86400000);
    return new Date(ms);
  }

  root.XlsxReader = { readWorkbook, excelDate, _internal: { parseSheet, parseSharedStrings, listZipEntries } };
})(typeof window !== 'undefined' ? window : globalThis);
