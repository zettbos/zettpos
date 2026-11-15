/**
 * ============================================================================
 * ZettPOS - Backend Controller & API (by Zettbos)
 * Standar Rekayasa: ZettBOT Engineering Standard v4.1
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// 1. PEMETAAN INDEKS KOLOM (INDEX-BASED MAPPING) - TAHAN PERUBAHAN HEADER
// ----------------------------------------------------------------------------
var COL_PRODUK = {
  ID: 0,
  BARCODE: 1,
  NAMA: 2,
  KATEGORI: 3,
  HARGA_BELI: 4,
  HARGA_JUAL: 5,
  STOK: 6,
  SATUAN: 7,
  STATUS: 8
};

var COL_TRANSAKSI = {
  INVOICE: 0,
  TANGGAL: 1,
  KASIR: 2,
  TOTAL_BELANJA: 3,
  DISKON: 4,
  TOTAL_AKHIR: 5,
  METODE_BAYAR: 6,
  JUMLAH_BAYAR: 7,
  KEMBALIAN: 8,
  TOTAL_MODAL: 9,
  LABA_BERSIH: 10,
  STATUS: 11
};

var COL_DETAIL = {
  INVOICE: 0,
  ID_PRODUK: 1,
  NAMA_PRODUK: 2,
  HARGA_BELI: 3,
  HARGA_JUAL: 4,
  QTY: 5,
  SUBTOTAL: 6
};

var COL_KATEGORI = {
  ID: 0,
  NAMA: 1
};

var COL_PENGATURAN = {
  KUNCI: 0,
  NILAI: 1
};

// ----------------------------------------------------------------------------
// 2. ROUTING WEB APP
// ----------------------------------------------------------------------------
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ZettPOS - Kasir & Penjualan Modern')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------------------------------
// 3. API DATA AWAL (INITIAL PAYLOAD DENGAN IN-MEMORY CACHING FRIENDLY)
// ----------------------------------------------------------------------------
function getInitialData() {
  try {
    var ss = getTargetSpreadsheet();
    var sheetProduk = ss.getSheetByName('Produk');
    var sheetKategori = ss.getSheetByName('Kategori');
    var sheetPengaturan = ss.getSheetByName('Pengaturan');
    var sheetTransaksi = ss.getSheetByName('Transaksi');

    // Pastikan database sudah terinisialisasi
    if (!sheetProduk || !sheetTransaksi) {
      initDatabase();
      sheetProduk = ss.getSheetByName('Produk');
      sheetKategori = ss.getSheetByName('Kategori');
      sheetPengaturan = ss.getSheetByName('Pengaturan');
      sheetTransaksi = ss.getSheetByName('Transaksi');
    }

    // 1. Ambil Pengaturan Toko
    var pengaturan = {};
    if (sheetPengaturan && sheetPengaturan.getLastRow() > 1) {
      var dataPengaturan = sheetPengaturan.getRange(2, 1, sheetPengaturan.getLastRow() - 1, 2).getDisplayValues();
      for (var p = 0; p < dataPengaturan.length; p++) {
        var key = dataPengaturan[p][COL_PENGATURAN.KUNCI];
        var val = dataPengaturan[p][COL_PENGATURAN.NILAI];
        if (key) pengaturan[key] = val;
      }
    }

    // Default Fallback Pengaturan
    pengaturan.NAMA_TOKO = pengaturan.NAMA_TOKO || 'ZETTBOS POS';
    pengaturan.ALAMAT_TOKO = pengaturan.ALAMAT_TOKO || 'Jl. Bisnis Sukses No. 1';
    pengaturan.NO_WA_TOKO = pengaturan.NO_WA_TOKO || '6281234567890';
    pengaturan.PESAN_STRUK = pengaturan.PESAN_STRUK || 'Terima kasih atas kunjungan Anda!';
    pengaturan.PIN_ADMIN = pengaturan.PIN_ADMIN || '1234';
    pengaturan.PIN_KASIR = pengaturan.PIN_KASIR || '0000';

    // 2. Ambil Kategori
    var kategoriList = [];
    if (sheetKategori && sheetKategori.getLastRow() > 1) {
      var dataKategori = sheetKategori.getRange(2, 1, sheetKategori.getLastRow() - 1, 2).getDisplayValues();
      for (var k = 0; k < dataKategori.length; k++) {
        var namaKat = dataKategori[k][COL_KATEGORI.NAMA];
        if (namaKat && kategoriList.indexOf(namaKat) === -1) {
          kategoriList.push(namaKat);
        }
      }
    }

    // 3. Ambil Master Produk
    var produkList = [];
    if (sheetProduk && sheetProduk.getLastRow() > 1) {
      var dataProduk = sheetProduk.getRange(2, 1, sheetProduk.getLastRow() - 1, 9).getDisplayValues();
      for (var i = 0; i < dataProduk.length; i++) {
        var row = dataProduk[i];
        var status = row[COL_PRODUK.STATUS] || 'Aktif';
        if (status === 'Aktif') {
          produkList.push({
            id: row[COL_PRODUK.ID],
            barcode: row[COL_PRODUK.BARCODE],
            nama: row[COL_PRODUK.NAMA],
            kategori: row[COL_PRODUK.KATEGORI],
            hargaBeli: parseAngka(row[COL_PRODUK.HARGA_BELI]),
            hargaJual: parseAngka(row[COL_PRODUK.HARGA_JUAL]),
            stok: parseAngka(row[COL_PRODUK.STOK]),
            satuan: row[COL_PRODUK.SATUAN] || 'Pcs',
            status: status
          });
        }
      }
    }

    // 4. Ringkasan Hari Ini (Omset & Transaksi)
    var hariIniStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
    var omsetHariIni = 0;
    var labaHariIni = 0;
    var totalTrxHariIni = 0;
    var transaksiHariIni = [];

    if (sheetTransaksi && sheetTransaksi.getLastRow() > 1) {
      var dataTrx = sheetTransaksi.getRange(2, 1, sheetTransaksi.getLastRow() - 1, 12).getDisplayValues();
      for (var t = dataTrx.length - 1; t >= 0; t--) {
        var rTrx = dataTrx[t];
        var tglWaktu = rTrx[COL_TRANSAKSI.TANGGAL] || '';
        var statusTrx = rTrx[COL_TRANSAKSI.STATUS] || 'Lunas';
        
        if (statusTrx === 'Lunas' && tglWaktu.indexOf(hariIniStr) === 0) {
          var nominal = parseAngka(rTrx[COL_TRANSAKSI.TOTAL_AKHIR]);
          var laba = parseAngka(rTrx[COL_TRANSAKSI.LABA_BERSIH]);
          omsetHariIni += nominal;
          labaHariIni += laba;
          totalTrxHariIni++;

          if (transaksiHariIni.length < 15) {
            transaksiHariIni.push({
              invoice: rTrx[COL_TRANSAKSI.INVOICE],
              tanggalWaktu: tglWaktu,
              kasir: rTrx[COL_TRANSAKSI.KASIR],
              totalAkhir: nominal,
              metodeBayar: rTrx[COL_TRANSAKSI.METODE_BAYAR],
              labaBersih: laba
            });
          }
        }
      }
    }

    return {
      success: true,
      data: {
        toko: {
          nama: pengaturan.NAMA_TOKO,
          alamat: pengaturan.ALAMAT_TOKO,
          telepon: pengaturan.NO_WA_TOKO,
          pesanStruk: pengaturan.PESAN_STRUK
        },
        produk: produkList,
        kategori: kategoriList,
        ringkasan: {
          omsetHariIni: omsetHariIni,
          labaHariIni: labaHariIni,
          totalTransaksi: totalTrxHariIni
        },
        transaksiTerbaru: transaksiHariIni,
        serverTime: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')
      }
    };
  } catch (error) {
    Logger.log('Error getInitialData: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ----------------------------------------------------------------------------
// 4. PEMROSESAN TRANSAKSI KASIR (MUTEX LOCK + SEQUENTIAL INVOICE + AUTO-STOK)
// ----------------------------------------------------------------------------
function simpanTransaksi(payload) {
  // LockService untuk mencegah nomor invoice ganda dan race condition stok
  var lock = LockService.getScriptLock();
  var successLock = lock.tryLock(30000);
  if (!successLock) {
    return { success: false, error: 'Sistem sedang sibuk memproses transaksi lain. Silakan coba lagi.' };
  }

  try {
    var ss = getTargetSpreadsheet();
    var sheetTransaksi = ss.getSheetByName('Transaksi');
    var sheetDetail = ss.getSheetByName('DetailTransaksi');
    var sheetProduk = ss.getSheetByName('Produk');

    if (!payload || !payload.items || payload.items.length === 0) {
      throw new Error('Keranjang belanja kosong!');
    }

    var now = new Date();
    var tanggalWaktu = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var periodeBulan = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMM'); // e.g. 202611

    // 1. Generate No Invoice Sekuensial (MAX + 1 pada bulan berjalan)
    var maxUrut = 0;
    var prefix = 'INV-' + periodeBulan + '-';

    if (sheetTransaksi.getLastRow() > 1) {
      var allInvoices = sheetTransaksi.getRange(2, 1, sheetTransaksi.getLastRow() - 1, 1).getDisplayValues();
      for (var i = 0; i < allInvoices.length; i++) {
        var inv = allInvoices[i][0] || '';
        if (inv.indexOf(prefix) === 0) {
          var nomorBagian = parseInt(inv.substring(prefix.length), 10);
          if (!isNaN(nomorBagian) && nomorBagian > maxUrut) {
            maxUrut = nomorBagian;
          }
        }
      }
    }

    var nextUrut = maxUrut + 1;
    var urutPadded = ('000' + nextUrut).slice(-3); // e.g. 001, 002
    var noInvoice = prefix + urutPadded;

    // 2. Baca Data Master Produk untuk Validasi Stok & HPP
    var dataProdukRange = sheetProduk.getRange(2, 1, Math.max(sheetProduk.getLastRow() - 1, 1), 9);
    var produkRows = dataProdukRange.getDisplayValues();
    var produkMap = {};

    for (var p = 0; p < produkRows.length; p++) {
      var rowP = produkRows[p];
      var idP = rowP[COL_PRODUK.ID];
      if (idP) {
        produkMap[idP] = {
          rowIndex: p + 2, // 1-based index di spreadsheet
          nama: rowP[COL_PRODUK.NAMA],
          hargaBeli: parseAngka(rowP[COL_PRODUK.HARGA_BELI]),
          hargaJual: parseAngka(rowP[COL_PRODUK.HARGA_JUAL]),
          stokSekarang: parseAngka(rowP[COL_PRODUK.STOK])
        };
      }
    }

    // 3. Validasi & Hitung Total Belanja, Total Modal (HPP), dan Detail
    var totalBelanja = 0;
    var totalModal = 0;
    var detailRowsToAppend = [];
    var stokUpdateList = [];

    for (var it = 0; it < payload.items.length; it++) {
      var item = payload.items[it];
      var master = produkMap[item.id];
      
      var qty = parseInt(item.qty, 10) || 1;
      var hargaJual = parseAngka(item.harga);
      var hargaBeli = master ? master.hargaBeli : 0;
      var subtotal = hargaJual * qty;

      totalBelanja += subtotal;
      totalModal += (hargaBeli * qty);

      // Detail baris untuk Tab DetailTransaksi
      detailRowsToAppend.push([
        noInvoice,
        item.id,
        item.nama || (master ? master.nama : 'Produk'),
        hargaBeli,
        hargaJual,
        qty,
        subtotal
      ]);

      // Kurangi stok jika produk terdaftar
      if (master) {
        var sisaStokBaru = Math.max(0, master.stokSekarang - qty);
        stokUpdateList.push({
          row: master.rowIndex,
          stokBaru: sisaStokBaru
        });
      }
    }

    var diskon = parseAngka(payload.diskon) || 0;
    var totalAkhir = Math.max(0, totalBelanja - diskon);
    var jumlahBayar = parseAngka(payload.jumlahBayar);
    var metodeBayar = payload.metodeBayar || 'Tunai';

    if (metodeBayar === 'Tunai' && jumlahBayar < totalAkhir) {
      throw new Error('Jumlah pembayaran tunai kurang dari total tagihan!');
    }

    var kembalian = (metodeBayar === 'Tunai') ? (jumlahBayar - totalAkhir) : 0;
    var labaBersih = totalAkhir - totalModal;
    var kasirName = payload.kasir || 'Kasir';

    // 4. Simpan ke Tab Transaksi
    var trxRow = [
      noInvoice,
      tanggalWaktu,
      kasirName,
      totalBelanja,
      diskon,
      totalAkhir,
      metodeBayar,
      jumlahBayar,
      kembalian,
      totalModal,
      labaBersih,
      'Lunas'
    ];
    sheetTransaksi.appendRow(trxRow);

    // 5. Simpan ke Tab DetailTransaksi secara Bulk
    if (detailRowsToAppend.length > 0) {
      var startDetailRow = sheetDetail.getLastRow() + 1;
      sheetDetail.getRange(startDetailRow, 1, detailRowsToAppend.length, 7).setValues(detailRowsToAppend);
    }

    // 6. Update Sisa Stok Produk di Sheet Produk
    for (var u = 0; u < stokUpdateList.length; u++) {
      var up = stokUpdateList[u];
      sheetProduk.getRange(up.row, COL_PRODUK.STOK + 1).setValue(up.stokBaru);
    }

    // 7. Flush Buffer untuk Menjamin Ketahanan Data
    SpreadsheetApp.flush();

    // Data siap cetak untuk Thermal Printer / WA
    return {
      success: true,
      message: 'Transaksi berhasil disimpan!',
      invoice: noInvoice,
      data: {
        invoice: noInvoice,
        tanggalWaktu: tanggalWaktu,
        kasir: kasirName,
        totalBelanja: totalBelanja,
        diskon: diskon,
        totalAkhir: totalAkhir,
        metodeBayar: metodeBayar,
        jumlahBayar: jumlahBayar,
        kembalian: kembalian,
        items: payload.items
      }
    };
  } catch (err) {
    Logger.log('Error simpanTransaksi: ' + err.message);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------------------------------------------
// 5. MANAJEMEN MASTER PRODUK (TAMBAH / EDIT / HAPUS)
// ----------------------------------------------------------------------------
function simpanProduk(data) {
  try {
    var ss = getTargetSpreadsheet();
    var sheetProduk = ss.getSheetByName('Produk');
    if (!sheetProduk) {
      initDatabase();
      sheetProduk = ss.getSheetByName('Produk');
    }

    var idProduk = data.id;
    var isNew = false;

    // Jika ID kosong, generate ID produk baru (PRD-XXX)
    if (!idProduk) {
      isNew = true;
      var maxID = 0;
      if (sheetProduk.getLastRow() > 1) {
        var existingIDs = sheetProduk.getRange(2, 1, sheetProduk.getLastRow() - 1, 1).getDisplayValues();
        for (var i = 0; i < existingIDs.length; i++) {
          var idStr = existingIDs[i][0] || '';
          if (idStr.indexOf('PRD-') === 0) {
            var num = parseInt(idStr.replace('PRD-', ''), 10);
            if (!isNaN(num) && num > maxID) maxID = num;
          }
        }
      }
      idProduk = 'PRD-' + ('000' + (maxID + 1)).slice(-3);
    }

    var rowValues = [
      idProduk,
      data.barcode || '',
      data.nama || 'Produk Baru',
      data.kategori || 'Umum',
      parseAngka(data.hargaBeli),
      parseAngka(data.hargaJual),
      parseAngka(data.stok),
      data.satuan || 'Pcs',
      data.status || 'Aktif'
    ];

    if (isNew) {
      sheetProduk.appendRow(rowValues);
    } else {
      // Cari baris produk yang akan diupdate
      var allRows = sheetProduk.getRange(2, 1, sheetProduk.getLastRow() - 1, 1).getDisplayValues();
      var targetRow = -1;
      for (var r = 0; r < allRows.length; r++) {
        if (allRows[r][0] === idProduk) {
          targetRow = r + 2;
          break;
        }
      }

      if (targetRow !== -1) {
        sheetProduk.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheetProduk.appendRow(rowValues);
      }
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Produk berhasil disimpan!',
      produk: {
        id: idProduk,
        barcode: data.barcode,
        nama: data.nama,
        kategori: data.kategori,
        hargaBeli: parseAngka(data.hargaBeli),
        hargaJual: parseAngka(data.hargaJual),
        stok: parseAngka(data.stok),
        satuan: data.satuan,
        status: data.status || 'Aktif'
      }
    };
  } catch (e) {
    Logger.log('Error simpanProduk: ' + e.message);
    return { success: false, error: e.message };
  }
}

function hapusProduk(idProduk) {
  try {
    var ss = getTargetSpreadsheet();
    var sheetProduk = ss.getSheetByName('Produk');
    if (!sheetProduk || sheetProduk.getLastRow() <= 1) {
      throw new Error('Data produk tidak ditemukan!');
    }

    var allIDs = sheetProduk.getRange(2, 1, sheetProduk.getLastRow() - 1, 1).getDisplayValues();
    var targetRow = -1;
    for (var i = 0; i < allIDs.length; i++) {
      if (allIDs[i][0] === idProduk) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow !== -1) {
      // Soft Delete: Ganti status menjadi Nonaktif agar riwayat transaksi masa lalu tetap utuh
      sheetProduk.getRange(targetRow, COL_PRODUK.STATUS + 1).setValue('Nonaktif');
      SpreadsheetApp.flush();
      return { success: true, message: 'Produk berhasil dinonaktifkan!' };
    } else {
      return { success: false, error: 'ID Produk tidak ditemukan.' };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ----------------------------------------------------------------------------
// 6. LAPORAN PENJUALAN & LABA BERSIH
// ----------------------------------------------------------------------------
function getLaporanPenjualan(startDateStr, endDateStr) {
  try {
    var ss = getTargetSpreadsheet();
    var sheetTransaksi = ss.getSheetByName('Transaksi');
    var sheetDetail = ss.getSheetByName('DetailTransaksi');

    if (!sheetTransaksi || sheetTransaksi.getLastRow() <= 1) {
      return {
        success: true,
        data: {
          totalOmset: 0,
          totalLaba: 0,
          totalTransaksi: 0,
          list: []
        }
      };
    }

    var dataTrx = sheetTransaksi.getRange(2, 1, sheetTransaksi.getLastRow() - 1, 12).getDisplayValues();
    var listHasil = [];
    var totalOmset = 0;
    var totalLaba = 0;

    for (var i = dataTrx.length - 1; i >= 0; i--) {
      var row = dataTrx[i];
      var tglWaktu = row[COL_TRANSAKSI.TANGGAL] || '';
      var tglOnly = tglWaktu.substring(0, 10);
      var status = row[COL_TRANSAKSI.STATUS] || 'Lunas';

      // Filter tanggal jika ada
      if (startDateStr && tglOnly < startDateStr) continue;
      if (endDateStr && tglOnly > endDateStr) continue;

      var akhir = parseAngka(row[COL_TRANSAKSI.TOTAL_AKHIR]);
      var laba = parseAngka(row[COL_TRANSAKSI.LABA_BERSIH]);

      if (status === 'Lunas') {
        totalOmset += akhir;
        totalLaba += laba;
      }

      listHasil.push({
        invoice: row[COL_TRANSAKSI.INVOICE],
        tanggal: tglWaktu,
        kasir: row[COL_TRANSAKSI.KASIR],
        totalBelanja: parseAngka(row[COL_TRANSAKSI.TOTAL_BELANJA]),
        diskon: parseAngka(row[COL_TRANSAKSI.DISKON]),
        totalAkhir: akhir,
        metodeBayar: row[COL_TRANSAKSI.METODE_BAYAR],
        labaBersih: laba,
        status: status
      });
    }

    return {
      success: true,
      data: {
        totalOmset: totalOmset,
        totalLaba: totalLaba,
        totalTransaksi: listHasil.length,
        list: listHasil
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ----------------------------------------------------------------------------
// 7. PENGATURAN TOKO & VALIDASI PIN ADMIN
// ----------------------------------------------------------------------------
function simpanPengaturan(settingsObj) {
  try {
    var ss = getTargetSpreadsheet();
    var sheet = ss.getSheetByName('Pengaturan');
    if (!sheet) {
      initDatabase();
      sheet = ss.getSheetByName('Pengaturan');
    }

    var existingData = [];
    if (sheet.getLastRow() > 1) {
      existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
    }

    var mapRow = {};
    for (var r = 0; r < existingData.length; r++) {
      var k = existingData[r][COL_PENGATURAN.KUNCI];
      if (k) mapRow[k] = r + 2;
    }

    for (var key in settingsObj) {
      var val = String(settingsObj[key]);
      if (mapRow[key]) {
        sheet.getRange(mapRow[key], 2).setValue(val);
      } else {
        sheet.appendRow([key, val]);
        mapRow[key] = sheet.getLastRow();
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Pengaturan berhasil diperbarui!' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function verifikasiPinAdmin(pinInput) {
  try {
    var ss = getTargetSpreadsheet();
    var sheet = ss.getSheetByName('Pengaturan');
    var pinAdminTersimpan = '1234';

    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === 'PIN_ADMIN') {
          pinAdminTersimpan = data[i][1];
          break;
        }
      }
    }

    if (String(pinInput).trim() === String(pinAdminTersimpan).trim()) {
      return { success: true };
    } else {
      return { success: false, error: 'PIN Admin salah!' };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ----------------------------------------------------------------------------
// 8. HELPER UTILITY
// ----------------------------------------------------------------------------
function parseAngka(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  // Hapus karakter non-digit selain minus atau titik
  var clean = String(val).replace(/[^0-9.-]/g, '');
  var num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}
