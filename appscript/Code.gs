// ============================================================
//  MONIC Laundry POS — Google Apps Script Backend
//  File    : Code.gs
//  Repo    : https://github.com/icaluwu/demo-laundry-pos-spreadsheet
//  Sheet   : Demo Laundry
//  Version : 2.2 — Fix Ghost Update (camelCase ↔ Proper_Case resolver)
// ============================================================

// ── GANTI dengan ID Spreadsheet Anda (dari URL) ──
// URL Contoh: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
var SPREADSHEET_ID = '1GZqwOJ9miDVrysECdyHq9nkbTFzEh2XGqkc2yDZMrXM';

// ── Timezone ──
var TIMEZONE = 'Asia/Jakarta';

// ============================================================
//  ENTRY POINT — POST (dipanggil dari index.html)
//  Payload: { sheetName, action, data }
//  action : 'create' (default) | 'update' | 'delete'
// ============================================================
function doPost(e) {
  try {
    var raw     = e.postData ? e.postData.contents : '{}';
    var payload = JSON.parse(raw);

    var action    = payload.action    || 'create';
    var sheetName = payload.sheetName || '';
    var data      = payload.data      || {};

    var result;

    // ── Route berdasarkan action ──
    if (action === 'update') {
      result = updateSheetRow(sheetName, data._id, data);
      return buildResponse(result);
    }

    if (action === 'delete') {
      result = deleteSheetRow(sheetName, data._id);
      return buildResponse(result);
    }

    // ── action === 'create' → route ke masing-masing handler ──
    switch (sheetName) {
      case 'Orders':
        result = appendDataTransaksi(data);
        break;
      case 'Customer':
        result = appendDataCustomer(data);
        break;
      case 'Deposit':
        result = appendDeposit(data);
        break;
      case 'Pengeluaran':
      case 'Pengeluaran_Outlet':
        result = appendPengeluaranOutlet(data);
        break;
      case 'Pengeluaran_Management':
        result = appendPengeluaranManagement(data);
        break;
      case 'Antar_Jemput':
        result = appendAntarJemput(data);
        break;
      case 'Laporan_Kerja_Karyawan':
        result = appendLaporanKerja(data);
        break;
      case 'Data_Karyawan':
        result = appendDataKaryawan(data);
        break;
      case 'Master_Layanan_Harga':
        result = appendMasterLayanan(data);
        break;
      default:
        result = { status: 'error', message: 'sheetName tidak dikenal: ' + sheetName };
    }

    return buildResponse(result);

  } catch (err) {
    return buildResponse({ status: 'error', message: err.message });
  }
}

// ============================================================
//  ENTRY POINT — GET (untuk testing via browser / READ data)
// ============================================================
function doGet(e) {
  var action = e.parameter.action || '';

  switch (action) {
    case 'getMasterLayanan':
      return buildResponse(getMasterLayanan());
    case 'getTransaksi':
      return buildResponse(getSheetData('Data_Transaksi'));
    case 'getPengeluaran':
      return buildResponse(getSheetData('Pengeluaran_Outlet'));
    case 'getCustomer':
      return buildResponse(getSheetData('Data_Customer'));
    case 'getKaryawan':
      return buildResponse(getSheetData('Data_Karyawan'));
    case 'getLaporanKerja':
      return buildResponse(getSheetData('Laporan_Kerja_Karyawan'));
    default:
      return buildResponse({ status: 'ok', message: 'MONIC Laundry API aktif ✅', version: '2.1.0' });
  }
}

// ============================================================
//  UNIVERSAL CRUD HELPERS (update & delete by _ID column)
// ============================================================

/**
 * Cari posisi baris dengan nilai _id di kolom bertanda '_ID'.
 * Pencarian dilakukan di SELURUH kolom sehingga posisi _ID bebas.
 * @return {Object|null} { rowIndex (1-based), colIndex, headers, rowData }
 */
function findRowById(sheet, id) {
  if (!id) return null;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  // Cari kolom '_ID'
  var idCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).toUpperCase() === '_ID') { idCol = c; break; }
  }
  if (idCol < 0) return null; // Sheet belum punya kolom _ID

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idCol]) === String(id)) {
      return { rowIndex: r + 1, colIndex: idCol, headers: headers, rowData: data[r] };
    }
  }
  return null;
}

/**
 * Normalize a string for fuzzy header matching:
 * lowercase + remove underscores, spaces, hyphens.
 * Example: 'Nama_Layanan' → 'namalayanan', 'namaLayanan' → 'namalayanan'
 */
function normalizeKey(str) {
  return String(str).toLowerCase().replace(/[_\s\-]/g, '');
}

/**
 * Per-sheet FIELD_MAP: maps camelCase frontend key → exact sheet header name.
 * Takes precedence over fuzzy matching. Add entries here for new sheets/fields.
 */
var FIELD_MAP = {
  'Data_Transaksi': {
    notaId: 'No_Nota', customer: 'Nama_Customer', phone: 'No_HP',
    layanan: 'Layanan', berat: 'Berat_Kg', total: 'Total_Bayar',
    metodeBayar: 'Metode_Bayar', statusBayar: 'Status_Bayar', status: 'Status', kasir: 'Kasir'
  },
  'Data_Customer': {
    nama: 'Nama', phone: 'No_HP', alamat: 'Alamat', email: 'Email'
  },
  'Data_Deposit': {
    customer: 'Nama_Customer', phone: 'No_HP', jumlah: 'Jumlah_Topup',
    saldoSebelum: 'Saldo_Sebelum', saldoSesudah: 'Saldo_Sesudah', kasir: 'Kasir'
  },
  'Pengeluaran_Outlet': {
    nama: 'Keterangan', keterangan: 'Keterangan',
    kategori: 'Kategori', jumlah: 'Jumlah', kasir: 'Kasir'
  },
  'Pengeluaran_Management': {
    deskripsi: 'Deskripsi', kategori: 'Kategori', jumlah: 'Jumlah',
    jenisPembayaran: 'Jenis_Pembayaran', catatanTambahan: 'Catatan_Tambahan',
    tanggal: 'Tanggal', dicatatOleh: 'Dicatat_Oleh'
  },
  'Antar_Jemput': {
    customer: 'Nama_Customer', phone: 'No_HP', alamat: 'Alamat',
    jadwal: 'Jadwal', tipe: 'Tipe', status: 'Status', kasir: 'Kasir'
  },
  'Laporan_Kerja_Karyawan': {
    nama: 'Nama_Karyawan', tanggal: 'Tanggal', jamMasuk: 'Jam_Masuk',
    jamKeluar: 'Jam_Keluar', totalJam: 'Total_Jam', keterangan: 'Keterangan'
  },
  'Data_Karyawan': {
    nama: 'Nama', jabatan: 'Jabatan', phone: 'No_HP',
    tglMasuk: 'Tgl_Masuk', gajiPokok: 'Gaji_Pokok', status: 'Status'
  },
  'Master_Layanan_Harga': {
    namaLayanan: 'Nama_Layanan', hargaPerKg: 'Harga_Per_Satuan',
    harga: 'Harga_Per_Satuan', satuan: 'Satuan',
    deskripsi: 'Deskripsi', aktif: 'Aktif'
  }
};

/**
 * Resolve a frontend key to a column index using three strategies:
 *   1. Exact string match          → fastest, no transform
 *   2. Per-sheet FIELD_MAP lookup  → explicit override
 *   3. Fuzzy normalised match      → strips _ and lowercases both sides
 * Returns column index (0-based), or -1 if not found.
 */
function resolveColIndex(headers, key, sheetName) {
  var i;
  // 1. Exact match
  for (i = 0; i < headers.length; i++) {
    if (String(headers[i]) === key) return i;
  }
  // 2. Explicit FIELD_MAP
  var mapped = (FIELD_MAP[sheetName] || {})[key];
  if (mapped) {
    for (i = 0; i < headers.length; i++) {
      if (String(headers[i]) === mapped) return i;
    }
  }
  // 3. Fuzzy normalised
  var normKey = normalizeKey(key);
  for (i = 0; i < headers.length; i++) {
    if (normalizeKey(headers[i]) === normKey) return i;
  }
  return -1;
}

/**
 * Update baris yang cocok dengan _id.
 * Menggunakan resolveColIndex() agar camelCase frontend
 * cocok dengan nama header Proper_Case di spreadsheet.
 */
function updateSheetRow(sheetName, id, newData) {
  if (!id) return { status: 'error', message: '_id diperlukan untuk update.' };

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan: ' + sheetName };

  var found = findRowById(sheet, id);
  if (!found) return { status: 'error', message: 'ID tidak ditemukan di sheet ' + sheetName + ': ' + id };

  // Salin baris lama lalu update field yang berhasil dipetakan
  var row = found.rowData.slice();
  Object.keys(newData).forEach(function(key) {
    if (key === '_id') return; // jangan timpa kolom _ID
    var colIdx = resolveColIndex(found.headers, key, sheetName);
    if (colIdx >= 0) row[colIdx] = newData[key];
  });

  // Auto-update Timestamp jika kolom ada
  var tsIdx = resolveColIndex(found.headers, 'Timestamp', sheetName);
  if (tsIdx >= 0) row[tsIdx] = formatTimestamp();

  sheet.getRange(found.rowIndex, 1, 1, row.length).setValues([row]);
  return { status: 'ok', message: 'Data diperbarui di ' + sheetName };
}

/**
 * Hapus baris yang cocok dengan _id.
 */
function deleteSheetRow(sheetName, id) {
  if (!id) return { status: 'error', message: '_id diperlukan untuk delete.' };

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan: ' + sheetName };

  var found = findRowById(sheet, id);
  if (!found) return { status: 'error', message: 'ID tidak ditemukan di sheet ' + sheetName + ': ' + id };

  sheet.deleteRow(found.rowIndex);
  return { status: 'ok', message: 'Data dihapus dari ' + sheetName };
}

// ============================================================
//  SHEET HANDLERS — CREATE
//  Catatan: _id dikirim dari frontend, disimpan di kolom _ID
//           (kolom terakhir) agar backward compatible dengan
//           data lama yang tidak memiliki kolom ini.
// ============================================================

// 1. Data_Transaksi
function appendDataTransaksi(data) {
  var headers = [
    'No_Nota', 'Nama_Customer', 'No_HP', 'Layanan',
    'Berat_Kg', 'Total_Bayar', 'Metode_Bayar',
    'Status', 'Kasir', 'Timestamp', '_ID'
  ];
  var row = [
    data.notaId      || autoNotaId(),
    data.customer    || '',
    data.phone       || '',
    data.layanan     || '',
    data.berat       || 0,
    data.total       || 0,
    data.metodeBayar || 'Tunai',
    data.status      || 'Proses',
    data.kasir       || 'Kasir Demo',
    formatTimestamp(),
    data._id         || ''
  ];
  return appendRow('Data_Transaksi', headers, row);
}

// 2. Data_Customer
function appendDataCustomer(data) {
  var headers = ['Nama', 'No_HP', 'Alamat', 'Email', 'Tgl_Daftar', '_ID'];
  var row = [
    data.nama   || '',
    data.phone  || '',
    data.alamat || '',
    data.email  || '',
    formatTimestamp(),
    data._id    || ''
  ];
  return appendRow('Data_Customer', headers, row);
}

// 3. Data_Deposit
function appendDeposit(data) {
  var headers = [
    'Nama_Customer', 'No_HP', 'Jumlah_Topup',
    'Saldo_Sebelum', 'Saldo_Sesudah', 'Kasir', 'Timestamp', '_ID'
  ];
  var row = [
    data.customer     || '',
    data.phone        || '',
    data.jumlah       || 0,
    data.saldoSebelum || 0,
    data.saldoSesudah || 0,
    data.kasir        || 'Kasir Demo',
    formatTimestamp(),
    data._id          || ''
  ];
  return appendRow('Data_Deposit', headers, row);
}

// 4. Pengeluaran_Outlet
function appendPengeluaranOutlet(data) {
  var headers = ['Keterangan', 'Kategori', 'Jumlah', 'Kasir', 'Timestamp', '_ID'];
  var row = [
    data.keterangan || data.nama || '',
    data.kategori   || 'Operasional',
    data.jumlah     || 0,
    data.kasir      || 'Kasir Demo',
    formatTimestamp(),
    data._id        || ''
  ];
  return appendRow('Pengeluaran_Outlet', headers, row);
}

// 5. Pengeluaran_Management
function appendPengeluaranManagement(data) {
  var headers = ['Deskripsi', 'Kategori', 'Jumlah', 'Tanggal', 'Dicatat_Oleh', 'Timestamp', '_ID'];
  var row = [
    data.deskripsi   || '',
    data.kategori    || '',
    data.jumlah      || 0,
    data.tanggal     || formatDate(),
    data.dicatatOleh || 'Admin',
    formatTimestamp(),
    data._id         || ''
  ];
  return appendRow('Pengeluaran_Management', headers, row);
}

// 6. Antar_Jemput
function appendAntarJemput(data) {
  var headers = [
    'Nama_Customer', 'No_HP', 'Alamat', 'Jadwal',
    'Tipe', 'Status', 'Kasir', 'Timestamp', '_ID'
  ];
  var row = [
    data.customer || '',
    data.phone    || '',
    data.alamat   || '',
    data.jadwal   || '',
    data.tipe     || 'Jemput',
    data.status   || 'Menunggu',
    data.kasir    || 'Kasir Demo',
    formatTimestamp(),
    data._id      || ''
  ];
  return appendRow('Antar_Jemput', headers, row);
}

// 7. Laporan_Kerja_Karyawan
function appendLaporanKerja(data) {
  var headers = [
    'Nama_Karyawan', 'Tanggal', 'Jam_Masuk',
    'Jam_Keluar', 'Total_Jam', 'Keterangan', 'Timestamp', '_ID'
  ];
  var row = [
    data.nama       || '',
    data.tanggal    || formatDate(),
    data.jamMasuk   || '',
    data.jamKeluar  || '',
    data.totalJam   || 0,
    data.keterangan || '',
    formatTimestamp(),
    data._id        || ''
  ];
  return appendRow('Laporan_Kerja_Karyawan', headers, row);
}

// 8. Data_Karyawan
function appendDataKaryawan(data) {
  var headers = [
    'Nama', 'Jabatan', 'No_HP', 'Tgl_Masuk',
    'Gaji_Pokok', 'Status', 'Timestamp', '_ID'
  ];
  var row = [
    data.nama      || '',
    data.jabatan   || '',
    data.phone     || '',
    data.tglMasuk  || formatDate(),
    data.gajiPokok || 0,
    data.status    || 'Aktif',
    formatTimestamp(),
    data._id       || ''
  ];
  return appendRow('Data_Karyawan', headers, row);
}

// 9. Master_Layanan_Harga
function appendMasterLayanan(data) {
  var headers = [
    'Nama_Layanan', 'Harga_Per_Satuan', 'Satuan',
    'Deskripsi', 'Aktif', 'Timestamp', '_ID'
  ];
  var row = [
    data.namaLayanan || '',
    data.hargaPerKg  || 0,
    data.satuan      || 'kg',
    data.deskripsi   || '',
    data.aktif !== undefined ? data.aktif : true,
    formatTimestamp(),
    data._id         || ''
  ];
  return appendRow('Master_Layanan_Harga', headers, row);
}

// ============================================================
//  READ HELPERS
// ============================================================

function getMasterLayanan() {
  return getSheetData('Master_Layanan_Harga');
}

function getSheetData(sheetName) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan: ' + sheetName };

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { status: 'ok', data: [] };

  var headers = values[0];
  var rows    = values.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  return { status: 'ok', data: rows };
}

// ============================================================
//  CORE UTILITY — appendRow
//  Auto-init header + style jika sheet masih kosong.
//  Jika sheet sudah ada data, tambah kolom _ID ke header
//  yang lama (migrasi backward-compatible).
// ============================================================
function appendRow(sheetName, headers, row) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return {
      status : 'error',
      message: 'Sheet "' + sheetName + '" tidak ditemukan di spreadsheet.'
    };
  }

  // Jika sheet masih kosong → tulis header baru
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    styleHeader(sheet, headers.length);

  } else {
    // Sheet sudah punya data — cek apakah kolom _ID sudah ada
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasIdCol = existingHeaders.some(function(h) {
      return String(h).toUpperCase() === '_ID';
    });

    // Jika belum ada kolom _ID, tambahkan ke kanan (migrasi non-destructive)
    if (!hasIdCol) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue('_ID');
      styleHeader(sheet, nextCol);
      // Tidak perlu isi nilai _ID untuk baris lama (biarkan kosong)
    }
  }

  sheet.appendRow(row);

  return {
    status : 'ok',
    message: 'Data berhasil ditambahkan ke ' + sheetName,
    sheet  : sheetName,
    id     : row[row.length - 1] // nilai _ID
  };
}

// ============================================================
//  UTILITIES
// ============================================================

function styleHeader(sheet, colCount) {
  var range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground('#1a56db');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  range.setFontSize(11);
  sheet.setFrozenRows(1);
}

function formatTimestamp() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
}

function formatDate() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy');
}

function autoNotaId() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Data_Transaksi');
  var count = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  var date  = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd');
  return 'MNC-' + date + '-' + String(count + 1).padStart(3, '0');
}

function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  SETUP — Jalankan SEKALI untuk inisialisasi / migrasi header
//  Aman dijalankan berulang: tidak menghapus data yang ada.
// ============================================================
function setupAllHeaders() {
  // Kolom _ID ditambahkan di akhir agar backward compatible
  var configs = [
    {
      name   : 'Data_Transaksi',
      headers: ['No_Nota','Nama_Customer','No_HP','Layanan','Berat_Kg','Total_Bayar','Metode_Bayar','Status','Kasir','Timestamp','_ID']
    },
    {
      name   : 'Data_Customer',
      headers: ['Nama','No_HP','Alamat','Email','Tgl_Daftar','_ID']
    },
    {
      name   : 'Data_Deposit',
      headers: ['Nama_Customer','No_HP','Jumlah_Topup','Saldo_Sebelum','Saldo_Sesudah','Kasir','Timestamp','_ID']
    },
    {
      name   : 'Pengeluaran_Outlet',
      headers: ['Keterangan','Kategori','Jumlah','Kasir','Timestamp','_ID']
    },
    {
      name   : 'Pengeluaran_Management',
      headers: ['Deskripsi','Kategori','Jumlah','Tanggal','Dicatat_Oleh','Timestamp','_ID']
    },
    {
      name   : 'Antar_Jemput',
      headers: ['Nama_Customer','No_HP','Alamat','Jadwal','Tipe','Status','Kasir','Timestamp','_ID']
    },
    {
      name   : 'Laporan_Kerja_Karyawan',
      headers: ['Nama_Karyawan','Tanggal','Jam_Masuk','Jam_Keluar','Total_Jam','Keterangan','Timestamp','_ID']
    },
    {
      name   : 'Data_Karyawan',
      headers: ['Nama','Jabatan','No_HP','Tgl_Masuk','Gaji_Pokok','Status','Timestamp','_ID']
    },
    {
      name   : 'Master_Layanan_Harga',
      headers: ['Nama_Layanan','Harga_Per_Satuan','Satuan','Deskripsi','Aktif','Timestamp','_ID']
    }
  ];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  configs.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      Logger.log('⚠️  Sheet tidak ditemukan: ' + config.name);
      return;
    }

    if (sheet.getLastRow() === 0) {
      // Sheet kosong → tulis header penuh
      sheet.appendRow(config.headers);
    } else {
      // Sheet sudah ada data → cek apakah _ID sudah ada
      var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var hasIdCol = existingHeaders.some(function(h) {
        return String(h).toUpperCase() === '_ID';
      });

      if (!hasIdCol) {
        // Tambah kolom _ID di kolom berikutnya (tidak overwrite data)
        var nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue('_ID');
        Logger.log('   ↳ Kolom _ID ditambahkan di kolom ' + nextCol);
      }

      // Update header baris 1 (tidak menyentuh data)
      var fullLen = config.headers.length;
      sheet.getRange(1, 1, 1, fullLen).setValues([config.headers]);
    }

    // Style header
    var range = sheet.getRange(1, 1, 1, config.headers.length);
    range.setBackground('#1a56db');
    range.setFontColor('#ffffff');
    range.setFontWeight('bold');
    range.setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, config.headers.length, 155);

    Logger.log('✅ Header siap: ' + config.name + ' (' + config.headers.length + ' kolom)');
  });

  // Seed data layanan jika Master_Layanan_Harga masih kosong
  seedMasterLayanan(ss);

  SpreadsheetApp.flush();
  Logger.log('🎉 Setup selesai! Semua sheet siap digunakan.');
}

function seedMasterLayanan(ss) {
  var sheet = ss.getSheetByName('Master_Layanan_Harga');
  if (!sheet || sheet.getLastRow() > 1) return; // Skip jika sudah ada data

  var layanan = [
    ['Cuci Kering Setrika', 7000,  'kg',  'Cuci + kering + setrika standar', true, formatTimestamp(), ''],
    ['Cuci Setrika',        8000,  'kg',  'Cuci basah + setrika',             true, formatTimestamp(), ''],
    ['Setrika Saja',        5000,  'kg',  'Hanya setrika',                    true, formatTimestamp(), ''],
    ['Cuci Kering',         6000,  'kg',  'Cuci + kering tanpa setrika',      true, formatTimestamp(), ''],
    ['Express 6 Jam',       15000, 'kg',  'Layanan kilat selesai 6 jam',      true, formatTimestamp(), ''],
    ['Bed Cover',           35000, 'pcs', 'Cuci bed cover / selimut tebal',   true, formatTimestamp(), ''],
    ['Karpet',              25000, 'pcs', 'Cuci karpet per meter',            true, formatTimestamp(), ''],
  ];

  layanan.forEach(function(row) { sheet.appendRow(row); });
  Logger.log('✅ Master_Layanan_Harga di-seed dengan ' + layanan.length + ' layanan.');
}

// ── Test — jalankan dari Apps Script Editor untuk cek koneksi ──
function testKoneksi() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('✅ Terhubung ke: ' + ss.getName());
  Logger.log('📋 Sheets tersedia:');
  ss.getSheets().forEach(function(s) {
    Logger.log('   • ' + s.getName() + ' (' + s.getLastRow() + ' baris)');
  });
}

// ── Test CRUD — jalankan untuk verifikasi update/delete berjalan ──
function testCRUD() {
  // --- Test UPDATE (gunakan camelCase persis seperti yang dikirim frontend) ---
  Logger.log('--- TEST UPDATE (camelCase keys) ---');
  var updateResult = updateSheetRow('Data_Karyawan', 'TEST-ID-001', {
    nama: 'Ahmad Diperbarui',
    status: 'Nonaktif',
    gajiPokok: 3000000
  });
  Logger.log(JSON.stringify(updateResult));

  Logger.log('--- TEST UPDATE Master_Layanan_Harga ---');
  var mlResult = updateSheetRow('Master_Layanan_Harga', 'TEST-ML-001', {
    namaLayanan: 'Express 3 Jam',
    hargaPerKg: 20000,
    satuan: 'kg',
    aktif: true
  });
  Logger.log(JSON.stringify(mlResult));

  // --- Test DELETE ---
  Logger.log('--- TEST DELETE ---');
  var deleteResult = deleteSheetRow('Data_Karyawan', 'TEST-ID-002');
  Logger.log(JSON.stringify(deleteResult));
}
