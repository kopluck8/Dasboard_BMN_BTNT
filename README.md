# Dashboard Aset BMN

Dashboard statis untuk membaca file **Daftar Aset BMN (.xlsx)**: ringkasan angka, kesimpulan otomatis, dan daftar hal yang perlu ditindaklanjuti. Dashboard ini dibuat untuk pimpinan (tab *Ringkasan*) sekaligus pengelola BMN (tab *Detail aset*, dengan filter dan unduhan CSV).

## Privasi: data tidak pernah meninggalkan browser

- File `.xlsx` **dibaca langsung di browser** pengguna. File tidak diunggah ke GitHub atau ke server mana pun.
- Halaman memasang *Content-Security-Policy* `connect-src 'none'`, sehingga browser memblokir semua permintaan jaringan dari halaman ini.
- Tidak ada library dari CDN. Semua kode ada di repo ini, termasuk pembaca .xlsx dan grafik.
- `.gitignore` menolak `*.xlsx` dan `*.csv` supaya data tidak ter-*commit* secara tidak sengaja.

Karena itu, repo **boleh publik**: yang dipublikasikan hanya kode, bukan data aset.

## Cara pakai

1. Buka halaman GitHub Pages repo ini, atau buka `index.html` langsung di browser.
2. Klik **Pilih file…** atau tarik file ekspor Daftar Aset (sheet *Master Aset*).
3. Baca hasilnya di empat tab:

| Tab | Isi |
|---|---|
| **Ringkasan** | Angka utama, kesimpulan otomatis, daftar hal yang perlu diperhatikan (urut menurut tingkat kepentingan), kondisi aset, jumlah dan nilai per jenis BMN |
| **Perlu perhatian** | Kandidat penghapusan beserta pemeriksaan konsistensinya, kelengkapan data, dan penggunaan BMN (idle / satker lain) |
| **Kendaraan & umur** | Tahun perolehan, umur dan penyusutan per jenis, daftar kendaraan |
| **Detail aset** | Tabel lengkap: cari, filter (jenis, kondisi, penanda), urutkan, unduh CSV. Klik baris untuk melihat semua kolom |

Setiap tombol **Lihat / Lihat daftar** membuka tab Detail yang sudah tersaring ke aset terkait, dan daftarnya bisa langsung diunduh sebagai CSV. CSV memakai pemisah `;` supaya langsung terbaca di Excel berbahasa Indonesia.

## Publikasi ke GitHub Pages

```bash
git init
git add .
git commit -m "Dashboard Aset BMN"
git branch -M main
git remote add origin https://github.com/<akun>/<nama-repo>.git
git push -u origin main
```

Lalu di GitHub buka **Settings → Pages → Build and deployment**, pilih *Source: Deploy from a branch*, *Branch: `main` / `(root)`*, kemudian **Save**. Setelah satu sampai dua menit, halaman tersedia di `https://<akun>.github.io/<nama-repo>/`.

## Aturan pemeriksaan

Semua aturan ada di `assets/js/analisis.js` (`FLAGS`). Daftar jenis aset dan ambang yang dipakai aturan diatur di `assets/js/config.js`.

| Penanda | Berlaku untuk | Aturan |
|---|---|---|
| Kandidat penghapusan | semua | Kondisi Rusak Berat **atau** Henti Guna = Ya **atau** Usul Hapus = Ya |
| Rusak Berat, belum Usul Hapus | semua | Rusak Berat dan Usul Hapus ≠ Ya |
| Henti Guna, belum Usul Hapus | semua | Henti Guna = Ya dan Usul Hapus ≠ Ya |
| Usul Hapus, kondisi Baik | semua | inkonsistensi data kondisi |
| Belum ada No PSP | semua | kolom No PSP kosong |
| Tanpa foto | semua | Jumlah Foto = 0 |
| Peralatan belum berlokasi ruang | Mesin Peralatan (TIK & non TIK) | Lokasi Ruang = "Belum berlokasi" |
| Tanah belum bersertipikat penuh | Tanah | Status Sertifikasi ≠ "Bersertipikat Seluruh Bidang" |
| Tanah/bangunan tanpa Kab/Kota | Tanah, Bangunan, Rumah Negara | Kab/Kota kosong |
| Kendaraan tanpa No Dokumen / STNK / No Polisi | Alat Angkutan Bermotor | kolom terkait kosong atau 0 |
| Nomor kendaraan ganda | Alat Angkutan Bermotor | No Polisi, No STNK, atau No Dokumen sama pada lebih dari satu kendaraan |
| Nilai buku habis | intrakomptabel, kecuali tanah | Nilai Buku = 0 |
| Aset tua | peralatan & kendaraan | umur sejak tanggal perolehan ≥ `umurTuaTahun` (default 10) |
| BMN tidak digunakan / satker lain | semua | dari kolom Status Penggunaan |

### Catatan penafsiran

- **"Belum Bersertipikat" pada peralatan dan "Kab/Kota kosong" pada barang bergerak itu wajar.** Karena itu, pemeriksaan hanya dijalankan pada jenis aset yang relevan, supaya tidak muncul alarm palsu.
- **Nilai buku Rp0 ≠ rusak.** Artinya masa manfaat akuntansi sudah habis. Aset ekstrakomptabel tidak dihitung dalam pemeriksaan ini karena memang tidak dikapitalisasi.
- **"Aset tua" memakai ambang kerja, bukan masa manfaat resmi.** Sesuaikan `umurTuaTahun` di `config.js` bila perlu.
- Kolom **"Umur Aset"** dari file tidak dipakai karena maknanya tidak jelas (nilainya tidak cocok dengan selisih tanggal perolehan). Umur dihitung dari *Tanggal Perolehan* sampai tanggal dashboard dibuka. Tanggal acuan bisa dikunci di `config.js` (`tanggalAcuan: '2026-06-30'`).
- Untuk kendaraan, nomor BPKB terbaca di kolom **No Dokumen**, bukan kolom *No BPKP* yang kosong.

## Struktur repo

```
index.html               halaman utama
assets/css/style.css     tampilan (tema terang & gelap)
assets/js/config.js      pengaturan aturan (ubah di sini)
assets/js/xlsx-reader.js pembaca .xlsx tanpa library
assets/js/analisis.js    normalisasi, aturan, kesimpulan
assets/js/charts.js      grafik ringan (bar, kolom, tooltip)
assets/js/app.js         pengendali tampilan
tests/cek-data.js        uji cepat di Node tanpa browser
```

## Uji cepat tanpa browser

```bash
node tests/cek-data.js /lokasi/daftar-aset.xlsx
```

Perintah ini mencetak angka total, hasil tiap pemeriksaan, dan teks kesimpulan. Gunakan untuk mencocokkan angka dengan Excel setelah mengubah aturan.

## Kebutuhan browser

Chrome/Edge, Firefox, atau Safari versi terbaru. Pembaca .xlsx memakai `DecompressionStream`, yang tidak tersedia di browser lama.

## Format file yang didukung

File ekspor Daftar Aset dengan baris judul yang memuat minimal: `Jenis BMN`, `Nama Barang`, `Kondisi`, `Nilai Perolehan`, `Nilai Buku`. Baris judul dan sheet yang benar dicari otomatis. Kolom lain bersifat opsional; pemeriksaan yang kolomnya tidak ada akan menghasilkan 0.
