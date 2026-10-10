/*
 * Pengaturan aturan analisis. Ubah nilai di sini bila kebijakan internal berbeda;
 * tidak perlu menyentuh kode lain.
 */
(function (root) {
  'use strict';
  root.ASET_CONFIG = {
    // Kolom minimal yang harus ada agar file dikenali sebagai Daftar Aset.
    requiredHeaders: ['Jenis BMN', 'Nama Barang', 'Kondisi', 'Nilai Perolehan', 'Nilai Buku'],

    // Jenis BMN yang wajar ditempatkan di ruangan (dicek kolom "Lokasi Ruang").
    // Tanah, bangunan, kendaraan, dan alat besar memang umumnya "Belum berlokasi".
    jenisWajibRuang: ['MESIN PERALATAN NON TIK', 'MESIN PERALATAN KHUSUS TIK'],

    jenisKendaraan: ['ALAT ANGKUTAN BERMOTOR'],
    jenisTanah: ['TANAH'],
    // Jenis yang wajib memiliki alamat / Kab-Kota.
    jenisBeralamat: ['TANAH', 'BANGUNAN DAN GEDUNG', 'RUMAH NEGARA'],
    // Jenis yang tidak disusutkan / umurnya tidak relevan untuk peremajaan.
    jenisTanpaPeremajaan: ['TANAH', 'BANGUNAN DAN GEDUNG', 'RUMAH NEGARA', 'JALAN DAN JEMBATAN', 'INSTALASI DAN JARINGAN', 'ASET TAK BERWUJUD'],

    // Ambang "aset tua" (tahun sejak tanggal perolehan) untuk peralatan & kendaraan.
    // Ini ambang kerja internal, BUKAN masa manfaat resmi. Sesuaikan bila perlu.
    umurTuaTahun: 10,

    // Tanggal acuan untuk menghitung umur. null = tanggal hari ini saat dashboard dibuka.
    tanggalAcuan: null,

    // Nilai tanggal penghapusan "kosong" di SIMAN (31-12-9999).
    tanggalKosongTahun: 9999
  };
})(typeof window !== 'undefined' ? window : globalThis);
