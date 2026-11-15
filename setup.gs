/**
 * ============================================================================
 * ZettPOS - Setup Database & Safe Migration (by Zettbos)
 * Standar Rekayasa: ZettBOT Engineering Standard v4.1
 * ============================================================================
 * Petunjuk:
 * 1. Buka Google Sheets baru di Google Drive Anda.
 * 2. Klik menu 'Ekstensi' > 'Apps Script'.
 * 3. Salin berkas setup.gs, code.gs, dan index.html ke editor Apps Script.
 * 4. Pilih fungsi 'initDatabase' lalu klik 'Jalankan' (Run) untuk membuat database.
 * ============================================================================
 */

function initDatabase() {
  var ss = getTargetSpreadsheet();
  
  // Definisi skema tabel baku (DILARANG MENGGUNAKAN KARAKTER UNDERSCORE PADA NAMA TAB)
  var schemas = {
    'Produk': [
      'IDProduk', 'Barcode', 'NamaProduk', 'Kategori', 
      'HargaBeli', 'HargaJual', 'Stok', 'Satuan', 'Status'
    ],
    'Transaksi': [
      'NoInvoice', 'TanggalWaktu', 'Kasir', 'TotalBelanja', 
      'Diskon', 'TotalAkhir', 'MetodeBayar', 'JumlahBayar', 
      'Kembalian', 'TotalModal', 'LabaBersih', 'Status'
    ],
    'DetailTransaksi': [
      'NoInvoice', 'IDProduk', 'NamaProduk', 'HargaBeli', 
      'HargaJual', 'Qty', 'Subtotal'
    ],
    'Kategori': [
      'IDKategori', 'NamaKategori'
    ],
    'Pengaturan': [
      'Kunci', 'Nilai'
    ]
  };

  var results = [];

  for (var sheetName in schemas) {
    var expectedHeaders = schemas[sheetName];
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // Buat tab baru jika belum ada
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(expectedHeaders);
      styleHeaderRow(sheet, expectedHeaders.length);
      results.push('Tab [' + sheetName + '] berhasil dibuat baru.');
      
      // Masukkan data awal (dummy/default) hanya untuk sheet yang baru dibuat
      seedInitialData(sheetName, sheet);
    } else {
      // Safe Migration: Sinkronisasi kolom baru tanpa menghapus data riil pengguna
      var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0] || [];
      var newHeadersToAdd = [];

      for (var i = 0; i < expectedHeaders.length; i++) {
        var colName = expectedHeaders[i];
        if (currentHeaders.indexOf(colName) === -1) {
          newHeadersToAdd.push(colName);
        }
      }

      if (newHeadersToAdd.length > 0) {
        var startCol = sheet.getLastColumn() + 1;
        for (var j = 0; j < newHeadersToAdd.length; j++) {
          sheet.getRange(1, startCol + j).setValue(newHeadersToAdd[j]);
        }
        styleHeaderRow(sheet, sheet.getLastColumn());
        results.push('Tab [' + sheetName + '] diperbarui dengan ' + newHeadersToAdd.length + ' kolom baru di ujung kanan.');
      } else {
        results.push('Tab [' + sheetName + '] sudah sesuai, data aman.');
      }
    }
  }

  // Hapus tab bawaan 'Sheet1' jika ada dan tab lain sudah siap
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {
      // Abaikan jika tidak diizinkan
    }
  }

  // Paksa simpan semua buffer ke disk Google Sheets
  SpreadsheetApp.flush();

  Logger.log('Inisialisasi selesai:\n' + results.join('\n'));
  return {
    success: true,
    message: 'Database ZettPOS berhasil diinisialisasi!',
    details: results,
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * Format visual header tabel agar rapi dan profesional
 */
function styleHeaderRow(sheet, numColumns) {
  if (numColumns < 1) return;
  var headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#0f172a'); // Slate 900
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

/**
 * Menghubungkan ke Spreadsheet aktif atau membuat baru jika dijalankan standalone
 */
function getTargetSpreadsheet() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    // Mode standalone
  }

  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SPREADSHEET_ID');

  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (err) {
      Logger.log('ID Spreadsheet lama tidak ditemukan, membuat baru...');
    }
  }

  var newSS = SpreadsheetApp.create('ZettPOS - Database Kasir & Penjualan');
  props.setProperty('SPREADSHEET_ID', newSS.getId());
  return newSS;
}

/**
 * Pengisian sampel data awal agar aplikasi langsung bisa dicoba setelah di-deploy
 */
function seedInitialData(sheetName, sheet) {
  if (sheetName === 'Kategori') {
    sheet.appendRow(['KAT-01', 'Makanan']);
    sheet.appendRow(['KAT-02', 'Minuman']);
    sheet.appendRow(['KAT-03', 'Snack & Camilan']);
  } 
  else if (sheetName === 'Produk') {
    sheet.appendRow(['PRD-001', '8991001', 'Kopi Susu Gula Aren', 'Minuman', 8000, 15000, 50, 'Cup', 'Aktif']);
    sheet.appendRow(['PRD-002', '8991002', 'Americano Dingin', 'Minuman', 6000, 12000, 40, 'Cup', 'Aktif']);
    sheet.appendRow(['PRD-003', '8991003', 'Croissant Butter', 'Makanan', 10000, 18000, 25, 'Pcs', 'Aktif']);
    sheet.appendRow(['PRD-004', '8991004', 'Roti Bakar Coklat Keju', 'Makanan', 9000, 16000, 30, 'Porsi', 'Aktif']);
    sheet.appendRow(['PRD-005', '8991005', 'Kentang Goreng Krispi', 'Snack & Camilan', 7000, 14000, 35, 'Porsi', 'Aktif']);
  } 
  else if (sheetName === 'Pengaturan') {
    sheet.appendRow(['NAMA_TOKO', 'ZETTBOS COFFEE & BAKERY']);
    sheet.appendRow(['ALAMAT_TOKO', 'Jl. Merdeka No. 45, Bandung']);
    sheet.appendRow(['NO_WA_TOKO', '6281234567890']);
    sheet.appendRow(['PESAN_STRUK', 'Terima kasih atas kunjungan Anda!\nFollow Instagram kami: @zettbos']);
    sheet.appendRow(['PIN_ADMIN', '1234']);
    sheet.appendRow(['PIN_KASIR', '0000']);
  }
}
