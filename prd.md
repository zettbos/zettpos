# PRD: Sistem Aplikasi Kasir (POS) & Penjualan Terpadu (by Zettbos)
================================================================================
DOKUMEN SPESIFIKASI KEBUTUHAN PRODUK & LIVING ARCHITECTURE
Status Dokumen: DISETUJUI & SELESAI DIIMPLEMENTASIKAN
Versi: 1.0.0 (Production Ready)
Standar Rekayasa: ZettBOT Engineering Standard v4.1
================================================================================

================================================================================
ZONA ATAS: SPESIFIKASI KEBUTUHAN PRODUK (BLUEPRINT AWAL)
================================================================================

## 1. NAMA PROJEK & PROBLEM STATEMENT
- **Nama Projek Sementara:** ZettPOS - Aplikasi Kasir & Manajemen Penjualan Modern
- **Problem Statement:** 
  Banyak pemilik usaha mikro, kecil, dan menengah (UMKM) masih mencatat penjualan dan stok barang secara manual menggunakan buku kertas atau file spreadsheet yang rawan salah hitung, rentan kehilangan data, dan memakan waktu lama saat rekapitulasi laba harian.
- **Tujuan Solusi:** 
  Menyediakan aplikasi kasir web berbasis Google Sheets + Google Apps Script yang 100% bebas biaya server, dapat diakses dari HP maupun laptop, dengan pencatatan otomatis transaksi, pemotongan stok instan, penomoran nota sekuensial, cetak struk thermal, dan laporan laba bersih real-time.

---

## 2. PERAN PENGGUNA & HAK AKSES (USER ROLES)
Sistem memiliki 2 tingkatan hak akses berbasis proteksi PIN cepat:
1. **Kasir (Cashier):**
   - Mengakses layar Transaksi Kasir (POS).
   - Mencari produk, menambah ke keranjang belanja, menerapkan diskon per transaksi.
   - Memilih metode pembayaran (Tunai, QRIS, Transfer Bank).
   - Menghitung uang diterima & nilai kembalian secara otomatis.
   - Mencetak struk kasir (ukuran 58mm / 80mm) dan membagikan nota digital ke WhatsApp pelanggan.
   - Melihat riwayat transaksi kasir hari ini.
   - *Batasan:* Tidak dapat melihat harga modal (HPP), tidak dapat melihat laba bersih, dan tidak dapat mengubah master produk/pengaturan.
2. **Pemilik Toko / Admin (Owner):**
   - Hak akses penuh Kasir + seluruh modul manajerial.
   - Modul Master Produk: Tambah, edit, hapus, kelola harga beli (modal), harga jual, dan sisa stok.
   - Modul Kategori Produk.
   - Modul Laporan Penjualan & Analisis Laba Bersih (Harian, Bulanan, Rentang Tanggal).
   - Modul Pengaturan Toko: Profil toko (nama, alamat, no HP/WA, footer nota), ganti PIN Kasir & Admin.

---

## 3. ARSITEKTUR & LINGKUNGAN KERJA (TECH STACK)
- **Platform Host:** Google Apps Script (Web App) — Zero server cost (Gratis selamanya).
- **Database Engine:** Google Sheets (terintegrasi langsung di Google Drive pengguna).
- **Format Arsitektur:** Monolith Terstruktur (3 berkas utama agar mudah dideploy):
  - `code.gs`: Backend API controller, LockService mutasi transaksi, routing `doGet`, kalkulasi data dengan `SpreadsheetApp.flush()`.
  - `setup.gs`: Inisialisasi tab sheet, header kolom baku (anti-underscore), dan migrasi non-destruktif (*Safe Migration*).
  - `index.html`: Antarmuka modern Anti-AI (Clean Slate & Navy), responsive layout, client-side in-memory cache, dan optimasi modal/toast tanpa alert browser.
- **Kompatibilitas Tampilan (Viewport):** Responsif Penuh (Mobile HP 360px+ dan Desktop PC 1024px+).

---

## 4. ALUR KERJA APLIKASI (USER FLOW)
1. **Buka Aplikasi:**
   - Halaman utama otomatis memuat antarmuka Kasir dengan katalog produk cepat & keranjang di sebelah kanan (atau tab bawah pada layar HP).
   - Data produk dimuat sekali ke dalam *Client-Side In-Memory Cache* untuk navigasi instan tanpa loading berulang.
2. **Proses Kasir (Checkout):**
   - Kasir mencari produk via filter kategori atau kotak pencarian instan (nama/kode).
   - Klik kartu/baris produk untuk menambahkan ke keranjang belanja.
   - Input kuantitas (Qty), ubah harga khusus jika diizinkan, atau pasang diskon nominal/persen.
   - Klik tombol **Bayar (Proses Transaksi)**:
     - Muncul modal pembayaran: Total Belanja, Metode Bayar (Tunai / QRIS / Transfer).
     - Jika Tunai: Input nominal diterima -> Otomatis hitung kembalian dengan tombol pecahan uang cepat (Rp 10rb, 20rb, 50rb, 100rb, Uang Pas).
   - Klik **Konfirmasi Simpan**:
     - Backend mengunci eksekusi (`LockService`), menghasilkan Nomor Invoice unik `INV-YYYYMM-XXX`, mengurangi stok produk secara atomik, menyimpan baris transaksi & detail ke sheet, lalu memanggil `SpreadsheetApp.flush()`.
     - Tampilan memunculkan modal Cetak Struk / Kirim WA dan keranjang otomatis dikosongkan.
3. **Peralihan ke Mode Admin:**
   - Kasir/Owner klik tombol gembok / avatar di navbar -> Masukkan PIN Admin.
   - Jika valid, tab navigasi admin (Master Produk, Laporan Laba, Pengaturan) terbuka.

---

## 5. DESAIN BASIS DATA (GOOGLE SHEETS)
*Aturan Baku: Tanpa karakter underscore (_) pada nama tab sheet. Nama kolom konsisten.*

### Tab 1: `Produk`
| Indeks | Nama Kolom | Tipe Data | Keterangan & Contoh |
| :--- | :--- | :--- | :--- |
| 0 | IDProduk | String | Kode unik produk, misal: `PRD-001` |
| 1 | Barcode | String | Kode barcode/SKU opsional, misal: `8999901001` |
| 2 | NamaProduk | String | Nama barang, misal: `Kopi Susu Gula Aren` |
| 3 | Kategori | String | Nama kategori, misal: `Minuman` |
| 4 | HargaBeli | Number | Harga modal per unit, misal: `8000` |
| 5 | HargaJual | Number | Harga jual ke pembeli, misal: `15000` |
| 6 | Stok | Number | Sisa stok barang saat ini, misal: `45` |
| 7 | Satuan | String | Satuan barang, misal: `Cup`, `Pcs`, `Botol` |
| 8 | Status | String | `Aktif` / `Nonaktif` |

### Tab 2: `Transaksi`
| Indeks | Nama Kolom | Tipe Data | Keterangan & Contoh |
| :--- | :--- | :--- | :--- |
| 0 | NoInvoice | String | Kode transaksi unik, misal: `INV-202611-001` |
| 1 | TanggalWaktu | String | Format: `yyyy-MM-dd HH:mm:ss` (WIB) |
| 2 | Kasir | String | Nama atau peran kasir yang bertugas |
| 3 | TotalBelanja | Number | Total belanja kotor sebelum diskon |
| 4 | Diskon | Number | Nilai potongan harga |
| 5 | TotalAkhir | Number | Nominal wajib dibayar pembeli |
| 6 | MetodeBayar | String | `Tunai`, `QRIS`, `Transfer Bank` |
| 7 | JumlahBayar | Number | Nominal uang yang diserahkan pelanggan |
| 8 | Kembalian | Number | Nominal kembalian |
| 9 | TotalModal | Number | Total HPP untuk perhitungan laba bersih |
| 10 | LabaBersih | Number | `TotalAkhir - TotalModal` |
| 11 | Status | String | `Lunas` / `Batal` |

### Tab 3: `DetailTransaksi`
| Indeks | Nama Kolom | Tipe Data | Keterangan & Contoh |
| :--- | :--- | :--- | :--- |
| 0 | NoInvoice | String | Foreign key ke Tab Transaksi |
| 1 | IDProduk | String | Kode barang |
| 2 | NamaProduk | String | Nama barang saat transaksi |
| 3 | HargaBeli | Number | Harga modal saat transaksi terjadi |
| 4 | HargaJual | Number | Harga jual per unit |
| 5 | Qty | Number | Kuantitas yang dibeli |
| 6 | Subtotal | Number | `HargaJual * Qty` |

### Tab 4: `Kategori`
| Indeks | Nama Kolom | Tipe Data | Keterangan & Contoh |
| :--- | :--- | :--- | :--- |
| 0 | IDKategori | String | Kode kategori, misal: `KAT-01` |
| 1 | NamaKategori | String | Nama kategori, misal: `Makanan`, `Minuman` |

### Tab 5: `Pengaturan`
| Indeks | Nama Kolom | Tipe Data | Keterangan & Contoh |
| :--- | :--- | :--- | :--- |
| 0 | Kunci | String | Parameter kunci: `NAMA_TOKO`, `ALAMAT`, `NO_WA`, `FOOTER_STRUK`, `PIN_ADMIN`, `PIN_KASIR` |
| 1 | Nilai | String | Isi konfigurasi |

---

## 6. ATURAN VALIDASI & LOGIKA KHUSUS (ENGINEERING LOGIC)
1. **Generator No Invoice Anti-Tabrakan:**
   - Format: `INV-YYYYMM-XXX` (misal: `INV-202611-001`).
   - Dihitung dari nomor urut tertinggi transaksi pada bulan berjalan (`MAX + 1`), bukan sekadar jumlah baris.
   - Dilindungi `LockService.getScriptLock()` dengan timeout 30 detik untuk menjamin keamanan multi-kasir.
2. **Pengurangan Stok Otomatis:**
   - Saat transaksi berhasil disimpan, stok setiap produk yang dibeli otomatis berkurang di tab `Produk`.
   - Validasi peringatan jika stok yang tersedia kurang dari jumlah yang dimasukkan ke keranjang.
3. **Penyimpanan Nilai HPP Historis:**
   - Harga modal dicatat permanen pada saat transaksi dibuat ke `DetailTransaksi` dan `Transaksi` agar laporan laba masa lalu tidak berubah jika harga modal produk dinaikkan di masa depan.
4. **Format Tanggal Baku & Anti-Timezone Shift:**
   - Seluruh pencatatan waktu menggunakan `Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')`.
   - Pembacaan tanggal menggunakan `getDisplayValues()` agar bebas distorsi konversi waktu browser.
5. **Anti-Freeze UI & Error Handling:**
   - Form submit wajib menggunakan `event.preventDefault()`.
   - Tombol simpan memunculkan animasi loading dan dilindungi `.withFailureHandler()` agar tombol tidak mati jika koneksi terputus.

---

## 7. RENCANA ANTARMUKA & ESTETIKA (HUMAN-CRAFTED ANTI-AI)
- **Tema Visual:** Clean Slate & Deep Navy (#0f172a, #1e293b, aksen Emerald #059669 untuk konfirmasi kasir).
- **Tipografi & Ikon:** Ikon vektor SVG inline yang tajam dan seragam (tanpa emoji norak yang bertebaran), spasi lapang, dan kartu bersudut membulat halus (`rounded-xl`).
- **Komponen Utama:**
  - **Top Navigation Bar:** Logo toko, jam digital real-time, status koneksi/kasir, tombol ganti mode (Kasir/Admin), dan tombol Fullscreen.
  - **Layar POS (Kasir):**
    - Sisi Kiri / Atas: Filter kategori pil horizontal + Kotak pencarian produk instan + Grid kartu produk responsif.
    - Sisi Kanan / Bawah (Mobile Drawer): Keranjang belanja aktif dengan tombol tambah/kurang kuantitas, ringkasan subtotal, diskon, dan tombol "Bayar Sekarang".
  - **Modal Pembayaran:** Kalkulator pecahan cepat uang tunai, tombol QRIS / Transfer, hitung kembalian otomatis.
  - **Modal Struk:** Preview struk nota belanja siap cetak ke printer bluetooth thermal 58mm/80mm atau ekspor PDF, serta tombol kirim teks nota rapi ke nomor WhatsApp pelanggan.
  - **Layar Admin (Dashboard & Laporan):**
    - Kartu metrik 3 kolom: Total Omset Hari Ini, Total Laba Bersih, Jumlah Transaksi.
    - Tabel Master Produk dengan tombol Tambah, Edit, dan Kelola Stok.
    - Tabel Laporan Penjualan dengan filter rentang tanggal dan pencarian nomor nota.

---

## 8. KRITERIA PENGUJIAN (ACCEPTANCE CRITERIA)
- [ ] Berkas `setup.gs` sukses membuat seluruh tab sheet tanpa error saat dijalankan pertama kali.
- [ ] Transaksi baru berhasil mengenerate nomor `INV-YYYYMM-001` secara berurutan.
- [ ] Stok produk di sheet `Produk` otomatis berkurang sesuai kuantitas yang dibeli.
- [ ] Nilai kembalian terhitung dengan tepat saat nominal uang tunai diinput.
- [ ] Struk kasir dapat dicetak dengan format rapi tanpa terpotong.
- [ ] Menu Admin terlindungi oleh PIN Admin.
- [ ] Halaman berjalan lancar di layar ponsel (mobile-friendly) maupun desktop PC tanpa bug layout.

================================================================================
ZONA BAWAH: MASTER CONTEXT & LOGIKA SISTEM TERKINI (LIVING CONTEXT)
================================================================================
- **Status Arsitektur:** Sistem ZettPOS v1.0.0 selesai dibangun penuh & live di production.
- **Repositori Resmi GitHub:** https://github.com/zettbos/zettpos
- **URL Domain Produksi Vercel:** https://zettpos.vercel.app
- **Tech Stack Aktif:** Google Apps Script + Google Sheets (Backend) & Vercel Static Web App (Frontend).
- **Skema Database Aktif:** Tab `Produk`, `Transaksi`, `DetailTransaksi`, `Kategori`, `Pengaturan` (Semua nama tab bebas underscore).
- **Logika Penomoran:** `INV-YYYYMM-XXX` sekuensial aman via LockService & mutasi atomik stok.
- **Riwayat Perubahan:**
  - [2026-11-16] Pembuatan cetak biru spesifikasi produk awal (PRD).
  - [2026-11-16] Implementasi penuh berkas `setup.gs`, `code.gs`, dan `index.html`.
  - [2026-11-16] Inisialisasi Git, pembuatan repositori GitHub `zettbos/zettpos`, dan deployment ke domain Vercel `https://zettpos.vercel.app`.
