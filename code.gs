/**
 * INA BASIC STORE - Warehouse V5.1 (HPP, Profit, QC Jahit & Auto-Cleanup)
 * Pembaruan V5.1:
 * - Menghapus otomatis tab/sheet lama yang sudah tidak terpakai saat setup() dijalankan
 *   (MASTER_BAHAN, PENGIRIMAN_BAHAN, STOK_BAHAN, LOG_BAHAN, STOK_LOKASI, MIGRASI_ALOKASI, dan Sheet1 kosong).
 * - Memastikan seluruh sheet operasional (WO_JAHIT, WO_JAHIT_ITEM, LAPORAN_JAHIT, OPNAME_SESI,
 *   OPNAME_DETAIL, MASTER_LOKASI, ATRIBUT_PRODUK) otomatis tersedia sebelum transaksi berjalan.
 * - Validasi kuantitas QC (Qty_Lolos + Qty_Gagal <= Qty_Terima) & idempotensi update harga/HPP.
 * - Mendukung debug parameter action di doGet(e) untuk pengecekan cepat lewat browser.
 */

var IBS = {
  SHEET_MASTER: 'MASTER_PRODUK',
  SHEET_ATTRIBUTES: 'ATRIBUT_PRODUK',
  SHEET_VARIANTS: 'VARIASI_PRODUK',
  SHEET_STOCK: 'STOK_VARIASI',
  SHEET_LOG: 'OPERATION_LOG',
  SHEET_LOCATIONS: 'MASTER_LOKASI',
  SHEET_OPNAME_SESI: 'OPNAME_SESI',
  SHEET_OPNAME_DETAIL: 'OPNAME_DETAIL',
  SHEET_WO_JAHIT: 'WO_JAHIT',
  SHEET_WO_JAHIT_ITEM: 'WO_JAHIT_ITEM',
  SHEET_LAPORAN_JAHIT: 'LAPORAN_JAHIT',
  DEPRECATED_SHEETS: [
    'MASTER_BAHAN',
    'PENGIRIMAN_BAHAN',
    'STOK_BAHAN',
    'LOG_BAHAN',
    'STOK_LOKASI',
    'MIGRASI_ALOKASI'
  ],
  LOCK_TIMEOUT_MS: 15000,
  HEADERS: {
    MASTER_PRODUK: ['Barcode', 'Nama', 'Status_Aktif', 'Created_At', 'Updated_At'],
    ATRIBUT_PRODUK: ['Barcode_Induk', 'Nama_Atribut', 'Tipe_Atribut', 'Nilai', 'Kode_Warna', 'Urutan_Atribut', 'Urutan_Nilai', 'Status_Aktif'],
    VARIASI_PRODUK: ['ID_Variasi', 'Barcode_Induk', 'Label_Variasi', 'Atribut_JSON', 'Status_Aktif', 'Created_At', 'Updated_At', 'HPP', 'Harga_Jual', 'Margin_Rp', 'Margin_Persen'],
    STOK_VARIASI: ['ID_Variasi', 'Barcode_Induk', 'Lokasi_Rak', 'Kuantitas', 'Updated_At'],
    OPERATION_LOG: ['Timestamp', 'ID_Transaksi', 'Aktivitas', 'Barcode', 'Lokasi_Asal', 'Lokasi_Tujuan', 'Qty', 'Status', 'ID_Variasi', 'Nama_Produk', 'Variasi_JSON', 'Label_Variasi', 'Produk_Tampilan', 'Nama_Pelapor', 'HPP_Satuan', 'Harga_Jual_Satuan', 'Total_Omzet', 'Total_Profit'],
    MASTER_LOKASI: ['Kode_Lokasi', 'Nama_Lokasi', 'Jenis_Lokasi', 'Status_Aktif', 'Created_At'],
    OPNAME_SESI: ['Session_ID', 'Nama_Sesi', 'Lokasi', 'Status', 'Dibuat_Oleh', 'Waktu_Mulai', 'Waktu_Selesai'],
    OPNAME_DETAIL: ['Detail_ID', 'Session_ID', 'Barcode', 'ID_Variasi', 'Label_Variasi', 'Qty_Sistem_Snapshot', 'Qty_Fisik', 'Nama_Pelapor', 'Waktu', 'Status_Hitung'],
    WO_JAHIT: ['WO_ID', 'Nomor_WO', 'Status', 'Tim_Label', 'Lokasi_Asal_Bahan', 'Lokasi_Tujuan', 'Tanggal_Target', 'Catatan', 'Dibuat_Oleh', 'Created_At'],
    WO_JAHIT_ITEM: ['Item_ID', 'WO_ID', 'Model_Warna_Ukuran', 'Target_Qty', 'Catatan'],
    LAPORAN_JAHIT: ['Report_ID', 'Transaction_ID', 'WO_ID', 'Nomor_WO', 'Nama_Penjahit', 'Qty_Bagus', 'Qty_Cacat', 'Status', 'Waktu_Lapor', 'Qty_Terima', 'Qty_Lolos', 'Qty_Gagal', 'Alasan_Gagal', 'Barcode_WIP', 'ID_Variasi_WIP', 'Lokasi_Tujuan', 'Nama_Penerima', 'Waktu_QC']
  }
};

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllOperationalSheets_(ss);

  // Hapus tab/sheet lama yang sudah tidak terpakai
  var deletedSheets = cleanupUnusedSheets_(ss);

  var sheets = [
    IBS.SHEET_MASTER, IBS.SHEET_ATTRIBUTES, IBS.SHEET_VARIANTS, IBS.SHEET_STOCK,
    IBS.SHEET_LOG, IBS.SHEET_LOCATIONS, IBS.SHEET_OPNAME_SESI,
    IBS.SHEET_OPNAME_DETAIL, IBS.SHEET_WO_JAHIT, IBS.SHEET_WO_JAHIT_ITEM,
    IBS.SHEET_LAPORAN_JAHIT
  ];

  sheets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      sh.setFrozenRows(1);
      var width = Math.max(sh.getLastColumn(), 1);
      sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d9eaf7');
    }
  });

  // Format teks (@) agar kode/barcode berawalan nol tidak berubah jadi angka
  ss.getSheetByName(IBS.SHEET_MASTER).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_ATTRIBUTES).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_VARIANTS).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_STOCK).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('D:F').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('I:I').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL).getRange('B:D').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT).getRange('A:D').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT).getRange('N:O').setNumberFormat('@');

  // Format mata uang & persentase pada VARIASI_PRODUK dan OPERATION_LOG
  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  shVar.getRange('H2:J').setNumberFormat('#,##0');
  shVar.getRange('K2:K').setNumberFormat('0.0%');
  syncVariantMarginFormulas_(shVar);

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  shLog.getRange('O2:R').setNumberFormat('#,##0');

  var locSheet = ss.getSheetByName(IBS.SHEET_LOCATIONS);

  // Bersihkan checkbox kosong di bawah data terakhir agar getLastRow() akurat
  cleanAndApplyCheckboxV2_(ss.getSheetByName(IBS.SHEET_MASTER), 3);
  cleanAndApplyCheckboxV2_(ss.getSheetByName(IBS.SHEET_ATTRIBUTES), 8);
  cleanAndApplyCheckboxV2_(shVar, 5);
  cleanAndApplyCheckboxV2_(locSheet, 4);

  var msg = 'Setup V5.1 selesai.';
  if (deletedSheets.length > 0) {
    msg += ' Tab dihapus: ' + deletedSheets.join(', ');
  }
  Logger.log(msg);
  try {
    ss.toast(msg, 'Ina Basic Store V5.1', 6);
  } catch (e) {}

  return { success: true, message: msg, deleted_sheets: deletedSheets };
}

function hapusTabTidakTerpakai() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deleted = cleanupUnusedSheets_(ss);
  var msg = deleted.length
    ? 'Berhasil menghapus tab tidak terpakai: ' + deleted.join(', ')
    : 'Tidak ada tab lama yang perlu dihapus.';
  Logger.log(msg);
  try {
    ss.toast(msg, 'Pembersihan Tab', 6);
  } catch (e) {}
  return deleted;
}

function cleanupUnusedSheets_(ss) {
  var deleted = [];
  IBS.DEPRECATED_SHEETS.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
      deleted.push(name);
    }
  });

  // Hapus tab bawaan 'Sheet1' jika masih kosong dan sudah ada sheet lain
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
    deleted.push('Sheet1');
  }

  return deleted;
}

function ensureAllOperationalSheets_(ss) {
  ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER, IBS.HEADERS.MASTER_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_ATTRIBUTES, IBS.HEADERS.ATRIBUT_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_VARIANTS, IBS.HEADERS.VARIASI_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_STOCK, IBS.HEADERS.STOK_VARIASI);
  ensureSheetHeadersV2_(ss, IBS.SHEET_LOG, IBS.HEADERS.OPERATION_LOG);

  var locSheet = ensureSheetHeadersV2_(ss, IBS.SHEET_LOCATIONS, IBS.HEADERS.MASTER_LOKASI);
  if (locSheet.getLastRow() === 1) {
    locSheet.appendRow(['RAK-UTAMA', 'Rak Utama', 'warehouse', true, new Date()]);
  }

  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT, IBS.HEADERS.WO_JAHIT);
  ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT_ITEM, IBS.HEADERS.WO_JAHIT_ITEM);
  ensureSheetHeadersV2_(ss, IBS.SHEET_LAPORAN_JAHIT, IBS.HEADERS.LAPORAN_JAHIT);
}

function syncVariantMarginFormulas_(shVar) {
  var lastRow = shVar.getLastRow();
  if (lastRow < 2) return;
  var count = lastRow - 1;
  var formulasJ = [];
  var formulasK = [];
  for (var i = 0; i < count; i++) {
    var r = i + 2;
    formulasJ.push(['=IF(I' + r + '>0, I' + r + '-H' + r + ', 0)']);
    formulasK.push(['=IF(I' + r + '>0, (I' + r + '-H' + r + ')/I' + r + ', 0)']);
  }
  shVar.getRange(2, 10, count, 1).setFormulas(formulasJ);
  shVar.getRange(2, 11, count, 1).setFormulas(formulasK);
}

function cleanAndApplyCheckboxV2_(sheet, checkboxCol) {
  if (!sheet) return;
  var maxRows = sheet.getMaxRows();
  var values = sheet.getRange(1, 1, maxRows, 1).getDisplayValues();
  var lastRealRow = 1;
  for (var i = values.length - 1; i >= 1; i--) {
    if (textV2_(values[i][0]) !== '') {
      lastRealRow = i + 1;
      break;
    }
  }
  if (maxRows > lastRealRow) {
    var emptyRange = sheet.getRange(lastRealRow + 1, 1, maxRows - lastRealRow, Math.max(sheet.getLastColumn(), 1));
    emptyRange.clearDataValidations();
    emptyRange.clearContent();
  }
  if (lastRealRow >= 2) {
    var dataRange = sheet.getRange(2, checkboxCol, lastRealRow - 1, 1);
    dataRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
  }
}

function applyRowCheckboxV2_(sheet, startRow, numRows, checkboxCol) {
  if (!sheet || numRows < 1) return;
  var range = sheet.getRange(startRow, checkboxCol, numRows, 1);
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
}

function ensureSheetHeadersV2_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  var currentWidth = Math.max(sh.getLastColumn(), 1);
  var current = sh.getRange(1, 1, 1, currentWidth).getDisplayValues()[0].map(function (v) { return String(v).trim(); });
  headers.forEach(function (header) {
    if (current.indexOf(header) === -1) {
      current.push(header);
      sh.getRange(1, current.length).setValue(header);
    }
  });
  return sh;
}

function assertV2SheetsReady_(ss) {
  var required = [
    IBS.SHEET_MASTER, IBS.SHEET_ATTRIBUTES, IBS.SHEET_VARIANTS, IBS.SHEET_STOCK,
    IBS.SHEET_LOG, IBS.SHEET_LOCATIONS, IBS.SHEET_OPNAME_SESI,
    IBS.SHEET_OPNAME_DETAIL, IBS.SHEET_WO_JAHIT, IBS.SHEET_WO_JAHIT_ITEM,
    IBS.SHEET_LAPORAN_JAHIT
  ];
  var missing = required.some(function (name) { return !ss.getSheetByName(name); });
  if (missing) {
    ensureAllOperationalSheets_(ss);
  }
}

function getActiveLocationsV2_(ss) {
  var sh = ss.getSheetByName(IBS.SHEET_LOCATIONS);
  if (!sh || sh.getLastRow() < 2) return [{ value: 'Rak Utama', label: 'Rak Utama', kind: 'warehouse' }];
  var rows = readObjectsV2_(sh);
  var active = [];
  rows.forEach(function (r) {
    if (booleanV2_(r.Status_Aktif, true)) {
      active.push({ value: textV2_(r.Nama_Lokasi), label: textV2_(r.Nama_Lokasi), kind: textV2_(r.Jenis_Lokasi) || 'warehouse' });
    }
  });
  return active.length ? active : [{ value: 'Rak Utama', label: 'Rak Utama', kind: 'warehouse' }];
}

function isValidLocationV2_(ss, value) {
  var locs = getActiveLocationsV2_(ss);
  return locs.some(function (l) { return l.value === textV2_(value); });
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? upperV2_(e.parameter.action) : '';
  if (action === 'GET_PENDING_QC') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    assertV2SheetsReady_(ss);
    return jsonResponseV2_(true, 'Daftar laporan pending.', getPendingQcReports_(ss));
  }
  if (action === 'GET_ACTIVE_WO') {
    var ssWo = SpreadsheetApp.getActiveSpreadsheet();
    assertV2SheetsReady_(ssWo);
    return jsonResponseV2_(true, 'Daftar WO aktif.', getActiveWoJahit_(ssWo));
  }
  return jsonResponseV2_(true, 'Ina Basic Store Warehouse API V5.1 (HPP, Profit & QC Jahit) aktif.', {
    version: '5.1',
    supported_actions: [
      'GET_PRODUCT', 'GET_LOCATIONS', 'GET_OPNAME_SESSIONS', 'GET_ACTIVE_WO', 'GET_PENDING_QC',
      'CREATE_PRODUCT', 'ADD_VARIANTS', 'UPDATE_PRODUCT', 'SET_VARIANT_STATUS',
      'UPDATE_VARIANT_PRICING', 'SAVE_TRANSACTION', 'DEACTIVATE_PRODUCT',
      'CREATE_OPNAME_SESSION', 'SUBMIT_OPNAME_COUNT', 'APPROVE_OPNAME_SESSION',
      'CREATE_WO_JAHIT', 'SUBMIT_LAPORAN_JAHIT', 'SUBMIT_QC_RECEIPT'
    ]
  });
}

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (parseError) {
    return jsonResponseV2_(false, 'Format permintaan tidak valid.');
  }

  var action = textV2_(payload.action).toUpperCase();
  if (!action && payload.activity) {
    var legacy = textV2_(payload.activity).toLowerCase();
    if (legacy === 'cekproduk') action = 'GET_PRODUCT';
    else if (legacy === 'editproduk') action = 'UPDATE_PRODUCT';
    else if (legacy === 'hapusproduk') action = 'DEACTIVATE_PRODUCT';
    else return jsonResponseV2_(false, 'Aplikasi lama tidak kompatibel. Perbarui aplikasi ke versi terbaru.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'GET_PRODUCT') {
    try {
      assertV2SheetsReady_(ss);
      var readResult = getProductActionV2_(ss, payload);
      return jsonResponseV2_(readResult.success, readResult.message, readResult.data);
    } catch (readError) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + (readError.message || String(readError)));
    }
  }

  if (action === 'GET_LOCATIONS') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Daftar lokasi aktif.', { locations: getActiveLocationsV2_(ss) });
    } catch (locError) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + (locError.message || String(locError)));
    }
  }

  if (action === 'GET_OPNAME_SESSIONS') {
    try {
      assertV2SheetsReady_(ss);
      var opnameRes = getOpnameSessions_(ss);
      return jsonResponseV2_(opnameRes.success, opnameRes.message || 'Daftar sesi opname.', opnameRes.data);
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  if (action === 'GET_ACTIVE_WO') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Daftar WO aktif.', getActiveWoJahit_(ss));
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  if (action === 'GET_PENDING_QC') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Daftar laporan pending.', getPendingQcReports_(ss));
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(IBS.LOCK_TIMEOUT_MS)) return jsonResponseV2_(false, 'Server sedang memproses transaksi lain. Coba lagi.');
  try {
    assertV2SheetsReady_(ss);
    var result;
    switch (action) {
      case 'CREATE_PRODUCT':
        result = createProductActionV2_(ss, payload);
        break;
      case 'ADD_VARIANTS':
        result = addVariantsActionV2_(ss, payload);
        break;
      case 'UPDATE_PRODUCT':
        result = updateProductActionV2_(ss, payload);
        break;
      case 'SET_VARIANT_STATUS':
        result = setVariantStatusActionV2_(ss, payload);
        break;
      case 'UPDATE_VARIANT_PRICING':
        result = updateVariantPricingActionV5_(ss, payload);
        break;
      case 'SAVE_TRANSACTION':
        result = saveStockTransactionV2_(ss, payload);
        break;
      case 'DEACTIVATE_PRODUCT':
        result = updateProductActionV2_(ss, { barcode: payload.barcode, status_aktif: false, transaction_id: payload.transaction_id });
        break;
      case 'CREATE_OPNAME_SESSION':
        result = createOpnameSession_(ss, payload);
        break;
      case 'SUBMIT_OPNAME_COUNT':
        result = submitOpnameCount_(ss, payload);
        break;
      case 'APPROVE_OPNAME_SESSION':
        result = approveOpnameSession_(ss, payload);
        break;
      case 'CREATE_WO_JAHIT':
        result = createWoJahit_(ss, payload);
        break;
      case 'SUBMIT_LAPORAN_JAHIT':
        result = submitLaporanJahit_(ss, payload);
        break;
      case 'SUBMIT_QC_RECEIPT':
        result = submitQcReceipt_(ss, payload);
        break;
      default:
        result = { success: false, message: 'Aksi tidak dikenal: ' + action };
    }
    return jsonResponseV2_(result.success, result.message, result.data);
  } catch (err) {
    console.error(err.stack || err);
    return jsonResponseV2_(false, 'Kesalahan server: ' + (err.message || String(err)));
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MODUL INTI: PRODUK, VARIASI, HPP, HARGA JUAL & TRANSAKSI
// ==========================================
function getProductActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  var includeInactive = payload.include_inactive === true;
  var data = composeProductDataV2_(ss, product, includeInactive);
  if (!includeInactive && !product.active) return { success: false, message: 'Produk sedang nonaktif.' };
  if (!data.variants.length && !includeInactive) return { success: false, message: 'Produk belum memiliki variasi. Buka menu Kelola Data untuk menambahkan variasi.' };
  return { success: true, message: 'Data produk ditemukan.', data: data };
}

function composeProductDataV2_(ss, product, includeInactive) {
  var attrRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_ATTRIBUTES));
  var variantRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_VARIANTS));
  var stockRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_STOCK));

  var definitions = {};
  attrRows.forEach(function (r) {
    if (upperV2_(r.Barcode_Induk) !== upperV2_(product.barcode)) return;
    var name = textV2_(r.Nama_Atribut);
    if (!name) return;
    if (!definitions[name]) definitions[name] = { name: name, type: textV2_(r.Tipe_Atribut) || 'text', order: Number(r.Urutan_Atribut) || 0, values: [] };
    definitions[name].values.push({
      value: textV2_(r.Nilai),
      hex: textV2_(r.Kode_Warna),
      order: Number(r.Urutan_Nilai) || 0,
      active: booleanV2_(r.Status_Aktif, true)
    });
  });

  var variantList = [];
  variantRows.forEach(function (r) {
    if (upperV2_(r.Barcode_Induk) !== upperV2_(product.barcode)) return;
    var active = booleanV2_(r.Status_Aktif, false);
    if (!includeInactive && !active) return;
    var attrs = parseJsonObjectV2_(r.Atribut_JSON);
    var stocks = {};
    stockRows.forEach(function (s) {
      if (textV2_(s.ID_Variasi) === textV2_(r.ID_Variasi)) stocks[textV2_(s.Lokasi_Rak)] = Number(s.Kuantitas) || 0;
    });
    var hpp = Number(r.HPP) || 0;
    var hargaJual = Number(r.Harga_Jual) || 0;
    var marginRp = hargaJual > 0 ? (hargaJual - hpp) : 0;
    var marginPersen = hargaJual > 0 ? ((hargaJual - hpp) / hargaJual) : 0;

    variantList.push({
      id: textV2_(r.ID_Variasi),
      label: textV2_(r.Label_Variasi),
      attributes: attrs,
      active: active,
      stocks: stocks,
      hpp: hpp,
      harga_jual: hargaJual,
      margin_rp: marginRp,
      margin_persen: marginPersen
    });
  });

  var fullDefinitions = Object.keys(definitions).map(function (name) {
    var def = definitions[name];
    def.values.sort(function (a, b) { return a.order - b.order; });
    return { name: def.name, type: def.type, order: def.order, values: def.values.slice() };
  }).sort(function (a, b) { return a.order - b.order; });

  var attrList = fullDefinitions.map(function (def) {
    var allowed = {};
    variantList.forEach(function (v) {
      if (!v.active) return;
      if (Object.prototype.hasOwnProperty.call(v.attributes, def.name)) allowed[textV2_(v.attributes[def.name])] = true;
    });
    return {
      name: def.name,
      type: def.type,
      order: def.order,
      values: def.values.filter(function (v) { return allowed[v.value] && v.active; })
    };
  });

  return {
    barcode: product.barcode,
    nama: product.name,
    aktif: product.active,
    attributes: attrList,
    attributeDefinitions: fullDefinitions,
    variants: variantList,
    locations: getActiveLocationsV2_(ss)
  };
}

function createProductActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var name = textV2_(payload.nama);
  var trxId = textV2_(payload.transaction_id);
  var attributes = normalizeAttributesV2_(payload.attributes || []);
  var variants = normalizeVariantsV2_(payload.variants || [], attributes);
  var productActive = booleanV2_(payload.status_aktif, true);
  var defaultHpp = Math.max(0, Number(payload.default_hpp) || 0);
  var defaultHargaJual = Math.max(0, Number(payload.default_harga_jual) || 0);

  if (!trxId) return { success: false, message: 'ID transaksi wajib diisi.' };
  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  if (!name || name.length < 2) return { success: false, message: 'Nama produk minimal 2 karakter.' };
  if (findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode)) return { success: false, message: 'Barcode sudah terdaftar.' };
  if (trxExistsV2_(ss.getSheetByName(IBS.SHEET_LOG), trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  if (!variants.length) return { success: false, message: 'Minimal satu kombinasi variasi harus dipilih.' };
  if (!variants.some(function (v) { return v.active; })) return { success: false, message: 'Minimal satu variasi harus berstatus aktif.' };

  var shMaster = ss.getSheetByName(IBS.SHEET_MASTER);
  var shAttrs = ss.getSheetByName(IBS.SHEET_ATTRIBUTES);
  var shVariants = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var activeLocs = getActiveLocationsV2_(ss);
  var now = new Date();
  var starts = [];
  try {
    var mStart = appendRowsV2_(shMaster, [[barcode, name, productActive, now, now]], starts);
    applyRowCheckboxV2_(shMaster, mStart, 1, 3);

    var attrRows = [];
    attributes.forEach(function (a, ai) {
      a.values.forEach(function (v, vi) {
        attrRows.push([barcode, a.name, a.type, v.value, v.hex || '', ai + 1, vi + 1, true]);
      });
    });
    if (attrRows.length) {
      var aStart = appendRowsV2_(shAttrs, attrRows, starts);
      applyRowCheckboxV2_(shAttrs, aStart, attrRows.length, 8);
    }

    var variantRows = [];
    var stockRows = [];
    var baseVarRow = shVariants.getLastRow() + 1;
    variants.forEach(function (v, idx) {
      var id = newVariantIdV2_();
      var label = makeVariantLabelV2_(v.attributes, attributes);
      var hppVal = v.hpp !== undefined ? Math.max(0, Number(v.hpp) || 0) : defaultHpp;
      var hjVal = v.harga_jual !== undefined ? Math.max(0, Number(v.harga_jual) || 0) : defaultHargaJual;
      var rNum = baseVarRow + idx;
      variantRows.push([
        id, barcode, label, JSON.stringify(v.attributes), v.active, now, now,
        hppVal, hjVal,
        '=IF(I' + rNum + '>0, I' + rNum + '-H' + rNum + ', 0)',
        '=IF(I' + rNum + '>0, (I' + rNum + '-H' + rNum + ')/I' + rNum + ', 0)'
      ]);
      activeLocs.forEach(function (loc) { stockRows.push([id, barcode, loc.value, 0, now]); });
    });
    var vStart = appendRowsV2_(shVariants, variantRows, starts);
    applyRowCheckboxV2_(shVariants, vStart, variantRows.length, 5);
    shVariants.getRange(vStart, 8, variantRows.length, 3).setNumberFormat('#,##0');
    shVariants.getRange(vStart, 11, variantRows.length, 1).setNumberFormat('0.0%');

    appendRowsV2_(shStock, stockRows, starts);
    appendRowsV2_(shLog, [[now, trxId, 'ProdukBaru', barcode, '', '', 0, 'SUCCESS', '', name, '', '', name, textV2_(payload.nama_pelapor), defaultHpp, defaultHargaJual, 0, 0]], starts);
    SpreadsheetApp.flush();
    return { success: true, message: 'Produk dan ' + variants.length + ' variasi berhasil didaftarkan.', data: { barcode: barcode, nama: name, jumlah_variasi: variants.length } };
  } catch (err) {
    rollbackAppendsV2_(starts);
    throw err;
  }
}

function addVariantsActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var trxId = textV2_(payload.transaction_id);
  var defaultHpp = Math.max(0, Number(payload.default_hpp) || 0);
  var defaultHargaJual = Math.max(0, Number(payload.default_harga_jual) || 0);

  if (!barcode || !trxId) return { success: false, message: 'Barcode dan ID transaksi wajib diisi.' };
  if (trxExistsV2_(ss.getSheetByName(IBS.SHEET_LOG), trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  if (!product.active) return { success: false, message: 'Aktifkan produk sebelum menambahkan variasi.' };

  var current = composeProductDataV2_(ss, product, true);
  var existingDefs = (current.attributeDefinitions || []).map(function (a) { return { name: a.name, type: a.type, values: a.values.map(function (v) { return { value: v.value, hex: v.hex }; }) }; });
  var incomingDefs = normalizeAttributesV2_(payload.attributes || []);
  var mergedDefs = [];
  if (!existingDefs.length) {
    mergedDefs = incomingDefs;
  } else {
    if (incomingDefs.length !== existingDefs.length) return { success: false, message: 'Jumlah atribut harus sama dengan definisi produk yang sudah ada.' };
    existingDefs.forEach(function (oldDef) {
      var incoming = incomingDefs.filter(function (d) { return d.name.toLowerCase() === oldDef.name.toLowerCase(); })[0];
      if (!incoming || incoming.type !== oldDef.type) throw new Error('Definisi atribut ' + oldDef.name + ' tidak cocok.');
      var values = oldDef.values.slice();
      incoming.values.forEach(function (v) {
        if (!values.some(function (x) { return x.value.toLowerCase() === v.value.toLowerCase(); })) values.push(v);
      });
      mergedDefs.push({ name: oldDef.name, type: oldDef.type, values: values });
    });
  }

  var variants = normalizeVariantsV2_(payload.variants || [], mergedDefs);
  if (!variants.length) return { success: false, message: 'Pilih minimal satu kombinasi variasi baru.' };
  var oldKeys = {};
  current.variants.forEach(function (v) { oldKeys[variantKeyV2_(v.attributes, mergedDefs)] = true; });
  variants.forEach(function (v) {
    var key = variantKeyV2_(v.attributes, mergedDefs);
    if (oldKeys[key]) throw new Error('Kombinasi sudah ada: ' + makeVariantLabelV2_(v.attributes, mergedDefs));
    oldKeys[key] = true;
  });

  var shAttrs = ss.getSheetByName(IBS.SHEET_ATTRIBUTES);
  var shVariants = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var activeLocs = getActiveLocationsV2_(ss);
  var starts = [];
  var changedVariantRows = [];
  var now = new Date();
  try {
    if (!existingDefs.length && mergedDefs.length) {
      current.variants.forEach(function (oldVariant) {
        if (oldVariant.active && Object.keys(oldVariant.attributes || {}).length === 0) {
          var oldRow = findVariantRowV2_(shVariants, oldVariant.id);
          if (oldRow) { changedVariantRows.push({ row: oldRow.row, oldActive: true, oldUpdated: shVariants.getRange(oldRow.row, 7).getValue() }); shVariants.getRange(oldRow.row, 5).setValue(false); shVariants.getRange(oldRow.row, 7).setValue(now); }
        }
      });
    }
    var newAttrRows = [];
    mergedDefs.forEach(function (def, ai) {
      var oldDef = existingDefs.filter(function (x) { return x.name === def.name; })[0];
      def.values.forEach(function (v, vi) {
        var alreadyExists = oldDef && oldDef.values.some(function (x) { return x.value.toLowerCase() === v.value.toLowerCase(); });
        if (!alreadyExists) newAttrRows.push([barcode, def.name, def.type, v.value, v.hex || '', ai + 1, vi + 1, true]);
      });
    });
    if (newAttrRows.length) {
      var aStart = appendRowsV2_(shAttrs, newAttrRows, starts);
      applyRowCheckboxV2_(shAttrs, aStart, newAttrRows.length, 8);
    }

    var vr = [], sr = [];
    var baseVarRow = shVariants.getLastRow() + 1;
    variants.forEach(function (v, idx) {
      var id = newVariantIdV2_();
      var rNum = baseVarRow + idx;
      vr.push([
        id, barcode, makeVariantLabelV2_(v.attributes, mergedDefs), JSON.stringify(v.attributes), v.active, now, now,
        defaultHpp, defaultHargaJual,
        '=IF(I' + rNum + '>0, I' + rNum + '-H' + rNum + ', 0)',
        '=IF(I' + rNum + '>0, (I' + rNum + '-H' + rNum + ')/I' + rNum + ', 0)'
      ]);
      activeLocs.forEach(function (loc) { sr.push([id, barcode, loc.value, 0, now]); });
    });
    var vStart = appendRowsV2_(shVariants, vr, starts);
    applyRowCheckboxV2_(shVariants, vStart, vr.length, 5);
    shVariants.getRange(vStart, 8, vr.length, 3).setNumberFormat('#,##0');
    shVariants.getRange(vStart, 11, vr.length, 1).setNumberFormat('0.0%');

    appendRowsV2_(shStock, sr, starts);
    appendRowsV2_(shLog, [[now, trxId, 'TambahVariasi', barcode, '', '', 0, 'SUCCESS', '', product.name, '', '', product.name, textV2_(payload.nama_pelapor), defaultHpp, defaultHargaJual, 0, 0]], starts);
    SpreadsheetApp.flush();
    return { success: true, message: variants.length + ' variasi baru berhasil ditambahkan.' };
  } catch (err) {
    rollbackAppendsV2_(starts);
    changedVariantRows.forEach(function (x) { try { shVariants.getRange(x.row, 5).setValue(x.oldActive); shVariants.getRange(x.row, 7).setValue(x.oldUpdated); } catch (restoreError) {} });
    throw err;
  }
}

function updateVariantPricingActionV5_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var items = payload.items || [];
  var pelapor = textV2_(payload.nama_pelapor) || 'Admin';
  var trxId = textV2_(payload.transaction_id) || ('PRICE-' + Utilities.getUuid().substring(0, 8).toUpperCase());

  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'Tidak ada data harga/HPP yang dikirim.' };

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (payload.transaction_id && trxExistsV2_(shLog, trxId)) {
    return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  }

  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };

  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var varRows = readObjectsV2_(shVar);
  var rowMap = {};
  varRows.forEach(function (r) {
    if (upperV2_(r.Barcode_Induk) === upperV2_(barcode)) {
      rowMap[textV2_(r.ID_Variasi)] = r;
    }
  });

  var now = new Date();
  var updatedCount = 0;

  items.forEach(function (item) {
    var vid = textV2_(item.variant_id);
    var target = rowMap[vid];
    if (!target) return;

    var hpp = Math.max(0, Number(item.hpp) || 0);
    var hargaJual = Math.max(0, Number(item.harga_jual) || 0);
    var rNum = target._row;

    shVar.getRange(rNum, 7).setValue(now);
    shVar.getRange(rNum, 8, 1, 4).setValues([[
      hpp,
      hargaJual,
      '=IF(I' + rNum + '>0, I' + rNum + '-H' + rNum + ', 0)',
      '=IF(I' + rNum + '>0, (I' + rNum + '-H' + rNum + ')/I' + rNum + ', 0)'
    ]]);
    shVar.getRange(rNum, 8, 1, 3).setNumberFormat('#,##0');
    shVar.getRange(rNum, 11).setNumberFormat('0.0%');
    updatedCount++;
  });

  if (updatedCount > 0) {
    shLog.appendRow([now, trxId, 'UpdateHargaHPP', barcode, '', '', updatedCount, 'SUCCESS', '', product.name, '', updatedCount + ' variasi', product.name + ' (Update HPP & Harga)', pelapor, 0, 0, 0, 0]);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'HPP & Harga Jual berhasil diperbarui untuk ' + updatedCount + ' variasi.' };
}

function updateProductActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var name = textV2_(payload.nama);
  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  var sh = ss.getSheetByName(IBS.SHEET_MASTER);
  var product = findMasterV2_(sh, barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  if (name && name.length < 2) return { success: false, message: 'Nama produk minimal 2 karakter.' };
  var active = booleanV2_(payload.status_aktif, product.active);
  var now = new Date();
  var log = ss.getSheetByName(IBS.SHEET_LOG);
  var trxId = textV2_(payload.transaction_id) || ('EDIT-' + Utilities.getUuid());
  if (trxExistsV2_(log, trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  var oldName = sh.getRange(product.row, 2).getValue();
  var oldActive = sh.getRange(product.row, 3).getValue();
  var oldUpdated = sh.getRange(product.row, 5).getValue();
  try {
    if (name) sh.getRange(product.row, 2).setValue(name);
    sh.getRange(product.row, 3).setValue(active);
    sh.getRange(product.row, 5).setValue(now);
    log.appendRow([now, trxId, 'EditProduk', barcode, '', '', 0, 'SUCCESS', '', name || product.name, '', '', name || product.name, textV2_(payload.nama_pelapor), 0, 0, 0, 0]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Data produk berhasil diperbarui.' };
  } catch (err) {
    try { sh.getRange(product.row, 2).setValue(oldName); sh.getRange(product.row, 3).setValue(oldActive); sh.getRange(product.row, 5).setValue(oldUpdated); } catch (restoreError) { console.error('Rollback edit produk: ' + restoreError); }
    throw err;
  }
}

function setVariantStatusActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var variantId = textV2_(payload.variant_id);
  if (!barcode || !variantId) return { success: false, message: 'Barcode dan ID variasi wajib diisi.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  var sh = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var row = findVariantRowV2_(sh, variantId);
  if (!row || upperV2_(row.barcode) !== upperV2_(barcode)) return { success: false, message: 'Variasi tidak ditemukan pada produk ini.' };
  var active = booleanV2_(payload.active, false);
  if (active && !product.active) return { success: false, message: 'Aktifkan produk induk terlebih dahulu.' };
  var now = new Date();
  var log = ss.getSheetByName(IBS.SHEET_LOG);
  var trxId = textV2_(payload.transaction_id) || ('VARSTATUS-' + Utilities.getUuid());
  if (trxExistsV2_(log, trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  var oldActive = sh.getRange(row.row, 5).getValue();
  var oldUpdated = sh.getRange(row.row, 7).getValue();
  try {
    sh.getRange(row.row, 5).setValue(active);
    sh.getRange(row.row, 7).setValue(now);
    log.appendRow([now, trxId, 'StatusVariasi', barcode, '', '', 0, 'SUCCESS', variantId, product.name, JSON.stringify(row.attributes || {}), row.label, product.name + ' · ' + row.label, textV2_(payload.nama_pelapor), row.hpp, row.harga_jual, 0, 0]);
    SpreadsheetApp.flush();
    return { success: true, message: active ? 'Variasi diaktifkan.' : 'Variasi dinonaktifkan.' };
  } catch (err) {
    try { sh.getRange(row.row, 5).setValue(oldActive); sh.getRange(row.row, 7).setValue(oldUpdated); } catch (restoreError) { console.error('Rollback status variasi: ' + restoreError); }
    throw err;
  }
}

function saveStockTransactionV2_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var action = textV2_(payload.activity);
  var barcode = textV2_(payload.barcode);
  var variantId = textV2_(payload.variant_id);
  var from = textV2_(payload.location_from);
  var to = textV2_(payload.location_to);
  var qty = Number(payload.qty);
  var pelapor = textV2_(payload.nama_pelapor);

  if (!trxId || !barcode || !variantId) return { success: false, message: 'ID transaksi, barcode, dan variasi wajib diisi.' };
  if (!pelapor) return { success: false, message: 'Nama pelapor wajib diisi untuk transaksi ini.' };
  if (action === 'Opname') return { success: false, message: 'Stock opname hanya dapat dilakukan melalui sesi opname.' };
  if (['Masuk', 'Keluar', 'Mutasi'].indexOf(action) === -1) return { success: false, message: 'Aktivitas tidak valid.' };
  if (textV2_(payload.qty) === '' || !isFinite(qty) || qty < 1 || Math.floor(qty) !== qty) return { success: false, message: 'Kuantitas tidak valid.' };
  if ((action === 'Keluar' || action === 'Mutasi') && !isValidLocationV2_(ss, from)) return { success: false, message: 'Lokasi asal tidak valid atau nonaktif.' };
  if ((action === 'Masuk' || action === 'Mutasi') && !isValidLocationV2_(ss, to)) return { success: false, message: 'Lokasi tujuan tidak valid atau nonaktif.' };
  if (action === 'Mutasi' && from === to) return { success: false, message: 'Lokasi asal dan tujuan tidak boleh sama.' };

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (trxExistsV2_(shLog, trxId)) return { success: false, message: 'Transaksi ini sudah pernah diproses.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Barcode belum terdaftar.' };
  if (!product.active) return { success: false, message: 'Produk sedang nonaktif.' };
  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant || upperV2_(variant.barcode) !== upperV2_(barcode)) return { success: false, message: 'Variasi tidak cocok dengan barcode produk.' };
  if (!variant.active) return { success: false, message: 'Variasi sedang nonaktif.' };

  var stockSheet = ss.getSheetByName(IBS.SHEET_STOCK);
  var stockFrom = (action === 'Keluar' || action === 'Mutasi') ? findStockRowV2_(stockSheet, variantId, from) : null;
  var stockTo = (action === 'Masuk' || action === 'Mutasi') ? findStockRowV2_(stockSheet, variantId, to) : null;
  var available = stockFrom ? stockFrom.qty : 0;
  if ((action === 'Keluar' || action === 'Mutasi') && available < qty) {
    return { success: false, message: 'Stok ' + variant.label + ' di ' + from + ' tidak cukup. Tersedia: ' + available };
  }

  var hppSatuan = variant.hpp || 0;
  var hargaJualSatuan = 0;
  var totalOmzet = 0;
  var totalProfit = 0;

  if (action === 'Keluar') {
    hargaJualSatuan = (payload.harga_jual !== undefined && payload.harga_jual !== '' && Number(payload.harga_jual) >= 0)
      ? Number(payload.harga_jual)
      : (variant.harga_jual || 0);
    totalOmzet = qty * hargaJualSatuan;
    totalProfit = hargaJualSatuan > 0 ? (qty * (hargaJualSatuan - hppSatuan)) : 0;
  }

  var changes = [];
  var now = new Date();
  try {
    if (action === 'Keluar' || action === 'Mutasi') {
      changes.push(writeStockQuantityV2_(stockSheet, variantId, barcode, from, available - qty, now));
    }
    if (action === 'Masuk' || action === 'Mutasi') {
      var toQty = stockTo ? stockTo.qty : 0;
      changes.push(writeStockQuantityV2_(stockSheet, variantId, barcode, to, toQty + qty, now));
    }
    var display = product.name + (variant.label ? ' · ' + variant.label : '');
    shLog.appendRow([
      now, trxId, action, barcode, from, to, qty, 'SUCCESS', variantId,
      product.name, JSON.stringify(variant.attributes || {}), variant.label, display, pelapor,
      hppSatuan, hargaJualSatuan, totalOmzet, totalProfit
    ]);
    var lastLogRow = shLog.getLastRow();
    shLog.getRange(lastLogRow, 15, 1, 4).setNumberFormat('#,##0');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Transaksi ' + action.toLowerCase() + ' berhasil disimpan.',
      data: {
        barcode: barcode,
        nama: product.name,
        variant_id: variantId,
        variasi: variant.label,
        total_omzet: totalOmzet,
        total_profit: totalProfit
      }
    };
  } catch (err) {
    rollbackStockChangesV2_(stockSheet, changes);
    throw err;
  }
}

function normalizeAttributesV2_(input) {
  if (!Array.isArray(input)) throw new Error('Format atribut produk tidak valid.');
  if (input.length > 6) throw new Error('Maksimal 6 atribut per produk pada versi ini.');
  var seen = {};
  return input.map(function (a, index) {
    var name = textV2_(a.name);
    var type = textV2_(a.type || 'text').toLowerCase();
    if (!name) throw new Error('Nama atribut ke-' + (index + 1) + ' wajib diisi.');
    if (seen[name.toLowerCase()]) throw new Error('Nama atribut tidak boleh duplikat: ' + name);
    seen[name.toLowerCase()] = true;
    if (['text', 'color'].indexOf(type) === -1) throw new Error('Tipe atribut hanya boleh text atau color.');
    if (!Array.isArray(a.values) || !a.values.length) throw new Error('Atribut ' + name + ' harus memiliki minimal satu nilai.');
    var valueSeen = {};
    var values = a.values.map(function (v) {
      var value = textV2_(typeof v === 'string' ? v : v.value);
      var hex = typeof v === 'string' ? '' : textV2_(v.hex);
      if (!value) throw new Error('Nilai atribut ' + name + ' tidak boleh kosong.');
      if (valueSeen[value.toLowerCase()]) throw new Error('Nilai atribut duplikat: ' + value + ' pada ' + name);
      valueSeen[value.toLowerCase()] = true;
      if (type === 'color' && hex && !/^#[0-9a-fA-F]{6}$/.test(hex)) throw new Error('Kode warna ' + value + ' harus berformat HEX #RRGGBB.');
      if (type === 'color' && !hex) throw new Error('Kode HEX wajib diisi untuk nilai warna ' + value + '.');
      return { value: value, hex: hex };
    });
    return { name: name, type: type, values: values };
  });
}

function normalizeVariantsV2_(input, attributes) {
  if (!Array.isArray(input)) throw new Error('Format kombinasi variasi tidak valid.');
  if (input.length > 300) throw new Error('Maksimal 300 kombinasi variasi dalam satu permintaan.');
  var attrNames = attributes.map(function (a) { return a.name; });
  var seen = {};
  return input.map(function (v) {
    var attrs = v.attributes || {};
    var normalized = {};
    attrNames.forEach(function (name) {
      var value = textV2_(attrs[name]);
      var definition = attributes.filter(function (a) { return a.name === name; })[0];
      if (!value || !definition.values.some(function (x) { return x.value === value; })) throw new Error('Nilai kombinasi tidak valid untuk atribut ' + name + '.');
      normalized[name] = value;
    });
    Object.keys(attrs).forEach(function (key) {
      if (attrNames.indexOf(key) === -1) throw new Error('Atribut tidak dikenal: ' + key);
    });
    var key = variantKeyV2_(normalized, attributes);
    if (seen[key]) throw new Error('Kombinasi variasi duplikat: ' + makeVariantLabelV2_(normalized, attributes));
    seen[key] = true;
    return {
      attributes: normalized,
      active: booleanV2_(v.active, true),
      hpp: v.hpp,
      harga_jual: v.harga_jual
    };
  });
}

function variantKeyV2_(attrs, definitions) {
  return definitions.map(function (d) { return d.name + '=' + textV2_(attrs[d.name]).toLowerCase(); }).join('|');
}

function makeVariantLabelV2_(attrs, definitions) {
  if (!definitions.length) return 'Standar';
  return definitions.map(function (d) { return d.name + ': ' + textV2_(attrs[d.name]); }).join(' · ');
}

function appendRowsV2_(sheet, rows, starts) {
  if (!rows || !rows.length) return 0;
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  starts.push({ sheet: sheet, start: start, count: rows.length });
  return start;
}

function rollbackAppendsV2_(starts) {
  for (var i = starts.length - 1; i >= 0; i--) {
    var x = starts[i];
    try { x.sheet.deleteRows(x.start, x.count); } catch (e) { console.error('Rollback append error: ' + e); }
  }
}

function trxExistsV2_(logSheet, trxId) {
  if (!trxId || logSheet.getLastRow() < 2) return false;
  var ids = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getDisplayValues();
  return ids.some(function (r) { return textV2_(r[0]) === trxId; });
}

function newVariantIdV2_() {
  return 'VAR-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
}

function parseJsonObjectV2_(value) {
  if (!value) return {};
  try {
    var parsed = JSON.parse(String(value));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch (e) { return {}; }
}

function jsonResponseV2_(success, message, data) {
  var result = { success: !!success, message: message || '' };
  if (data !== undefined) result.data = data;
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// MODUL WORK ORDER JAHIT & LAPORAN QC
// ==========================================
function createWoJahit_(ss, payload) {
  var nomorWo = textV2_(payload.nomor_wo) || ('WO-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 900 + 100));
  var timLabel = textV2_(payload.tim_label);
  var lokasiAsal = textV2_(payload.lokasi_asal_bahan);
  var lokasiTujuan = textV2_(payload.lokasi_tujuan);
  var tanggalTarget = textV2_(payload.tanggal_target);
  var catatan = textV2_(payload.catatan);
  var pelapor = textV2_(payload.nama_pelapor);
  var items = payload.items || [];

  if (!timLabel || !lokasiAsal || !pelapor) return { success: false, message: 'Tim/Label, lokasi asal, dan nama pelapor wajib diisi.' };
  if (!isValidLocationV2_(ss, lokasiAsal)) return { success: false, message: 'Lokasi asal tidak valid.' };
  if (!items.length) return { success: false, message: 'Minimal harus ada 1 target item pada Work Order.' };

  var woId = 'WOJ-' + Utilities.getUuid().split('-')[0].toUpperCase();
  var shWo = ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT, IBS.HEADERS.WO_JAHIT);
  var shItem = ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT_ITEM, IBS.HEADERS.WO_JAHIT_ITEM);
  var now = new Date();

  shWo.appendRow([woId, nomorWo, 'DRAFT', timLabel, lokasiAsal, lokasiTujuan, tanggalTarget, catatan, pelapor, now]);

  var itemRows = [];
  items.forEach(function (item, idx) {
    var itemId = woId + '-ITM-' + (idx + 1);
    var deskripsi = textV2_(item.deskripsi);
    var qty = Number(item.target_qty) || 0;
    itemRows.push([itemId, woId, deskripsi, qty, textV2_(item.catatan)]);
  });

  if (itemRows.length) {
    shItem.getRange(shItem.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Work Order Jahit berhasil dibuat.', data: { wo_id: woId, nomor_wo: nomorWo } };
}

function getActiveWoJahit_(ss) {
  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  if (!shWo || shWo.getLastRow() < 2) return { wo_list: [] };

  var shItem = ss.getSheetByName(IBS.SHEET_WO_JAHIT_ITEM);
  var shLap = ensureSheetHeadersV2_(ss, IBS.SHEET_LAPORAN_JAHIT, IBS.HEADERS.LAPORAN_JAHIT);

  var woRows = readObjectsV2_(shWo);
  var itemRows = shItem && shItem.getLastRow() >= 2 ? readObjectsV2_(shItem) : [];
  var lapRows = shLap.getLastRow() >= 2 ? readObjectsV2_(shLap) : [];

  var list = woRows.filter(function (w) {
    var st = upperV2_(w.Status);
    return st !== 'SELESAI' && st !== 'BATAL';
  }).map(function (w) {
    var woId = textV2_(w.WO_ID);
    var noWo = textV2_(w.Nomor_WO);
    var woItems = itemRows.filter(function (it) { return textV2_(it.WO_ID) === woId; });
    var targetQty = woItems.reduce(function (sum, it) { return sum + (Number(it.Target_Qty) || 0); }, 0);
    var masukQty = lapRows.filter(function (lp) {
      return (woId && textV2_(lp.WO_ID) === woId) || upperV2_(lp.Nomor_WO) === upperV2_(noWo);
    }).reduce(function (sum, lp) {
      var st = upperV2_(lp.Status);
      if (st === 'APPROVED') return sum + (Number(lp.Qty_Lolos) || 0);
      if (st === 'PENDING') return sum + (Number(lp.Qty_Bagus) || 0);
      return sum;
    }, 0);
    var summary = woItems.map(function (it) {
      return textV2_(it.Model_Warna_Ukuran) + ' (' + (Number(it.Target_Qty) || 0) + ')';
    }).join(', ');

    return {
      wo_id: woId,
      nomor_wo: noWo,
      status: textV2_(w.Status),
      tim_label: textV2_(w.Tim_Label),
      lokasi_tujuan: textV2_(w.Lokasi_Tujuan),
      tanggal_target: textV2_(w.Tanggal_Target),
      target_qty: targetQty,
      masuk_qty: masukQty,
      sisa_qty: Math.max(0, targetQty - masukQty),
      target_summary: summary,
      items: woItems.map(function (it) {
        return {
          item_id: textV2_(it.Item_ID),
          deskripsi: textV2_(it.Model_Warna_Ukuran),
          target_qty: Number(it.Target_Qty) || 0,
          catatan: textV2_(it.Catatan)
        };
      })
    };
  });

  return { wo_list: list };
}

function submitLaporanJahit_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var woIdInput = textV2_(payload.wo_id);
  var nomorWoInput = textV2_(payload.nomor_wo);
  var namaPenjahit = textV2_(payload.nama_penjahit);
  var qtyBagus = Number(payload.qty_bagus) || 0;
  var qtyCacat = Number(payload.qty_cacat) || 0;

  if ((!woIdInput && !nomorWoInput) || !namaPenjahit) {
    return { success: false, message: 'Nomor WO dan Nama Penjahit wajib diisi.' };
  }
  if (qtyBagus < 0 || qtyCacat < 0 || (qtyBagus + qtyCacat) <= 0) {
    return { success: false, message: 'Jumlah jahitan minimal 1 pcs dan tidak boleh negatif.' };
  }

  var shLap = ensureSheetHeadersV2_(ss, IBS.SHEET_LAPORAN_JAHIT, IBS.HEADERS.LAPORAN_JAHIT);
  var existing = readObjectsV2_(shLap);
  if (trxId && existing.some(function (r) { return textV2_(r.Transaction_ID) === trxId; })) {
    return { success: false, message: 'Laporan ini sudah pernah dikirim.' };
  }

  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var woRows = shWo && shWo.getLastRow() >= 2 ? readObjectsV2_(shWo) : [];
  var wo = woRows.filter(function (w) {
    if (woIdInput && textV2_(w.WO_ID) === woIdInput) return true;
    if (nomorWoInput && upperV2_(w.Nomor_WO) === upperV2_(nomorWoInput)) return true;
    return false;
  })[0];

  if (wo && (upperV2_(wo.Status) === 'SELESAI' || upperV2_(wo.Status) === 'BATAL')) {
    return { success: false, message: 'Sesi WO ' + textV2_(wo.Nomor_WO) + ' sudah ditutup.' };
  }

  var finalWoId = wo ? textV2_(wo.WO_ID) : woIdInput;
  var finalNomorWo = wo ? textV2_(wo.Nomor_WO) : nomorWoInput;
  var defaultLokasi = wo ? textV2_(wo.Lokasi_Tujuan) : '';
  var reportId = 'QC-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000);

  shLap.appendRow([
    reportId, trxId, finalWoId, finalNomorWo, namaPenjahit,
    qtyBagus, qtyCacat, 'PENDING', new Date(),
    0, 0, 0, '', '', '', defaultLokasi, '', ''
  ]);
  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Laporan jahit berhasil dikirim untuk verifikasi QC.',
    data: { report_id: reportId, nomor_wo: finalNomorWo }
  };
}

function getPendingQcReports_(ss) {
  var shLap = ensureSheetHeadersV2_(ss, IBS.SHEET_LAPORAN_JAHIT, IBS.HEADERS.LAPORAN_JAHIT);
  if (shLap.getLastRow() < 2) return { reports: [] };

  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var shItem = ss.getSheetByName(IBS.SHEET_WO_JAHIT_ITEM);
  var woRows = shWo && shWo.getLastRow() >= 2 ? readObjectsV2_(shWo) : [];
  var itemRows = shItem && shItem.getLastRow() >= 2 ? readObjectsV2_(shItem) : [];

  var reports = readObjectsV2_(shLap).filter(function (r) {
    return upperV2_(r.Status) === 'PENDING';
  }).map(function (r) {
    var wo = woRows.filter(function (w) {
      return (textV2_(r.WO_ID) && textV2_(w.WO_ID) === textV2_(r.WO_ID)) ||
             upperV2_(w.Nomor_WO) === upperV2_(r.Nomor_WO);
    })[0];
    var woId = wo ? textV2_(wo.WO_ID) : textV2_(r.WO_ID);
    var items = itemRows.filter(function (it) { return textV2_(it.WO_ID) === woId; });
    var summary = items.map(function (it) {
      return textV2_(it.Model_Warna_Ukuran) + ' (' + (Number(it.Target_Qty) || 0) + ')';
    }).join(', ');

    return {
      report_id: textV2_(r.Report_ID),
      wo_id: woId,
      nomor_wo: textV2_(r.Nomor_WO),
      tim_label: wo ? textV2_(wo.Tim_Label) : '',
      lokasi_tujuan: textV2_(r.Lokasi_Tujuan) || (wo ? textV2_(wo.Lokasi_Tujuan) : 'Rak Utama'),
      nama_penjahit: textV2_(r.Nama_Penjahit),
      qty_bagus: Number(r.Qty_Bagus) || 0,
      qty_cacat: Number(r.Qty_Cacat) || 0,
      waktu_lapor: r.Waktu_Lapor,
      target_summary: summary
    };
  });

  return { reports: reports };
}

function submitQcReceipt_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var reportId = textV2_(payload.report_id);
  var namaPenerima = textV2_(payload.nama_penerima);
  var qtyTerima = Number(payload.qty_terima) || 0;
  var qtyLolos = Number(payload.qty_lolos) || 0;
  var qtyGagal = Number(payload.qty_gagal) || 0;
  var alasanGagal = textV2_(payload.alasan_gagal);
  var barcodeWip = textV2_(payload.barcode_wip);
  var variantId = textV2_(payload.id_variasi_wip);
  var lokasiTujuan = textV2_(payload.lokasi_tujuan);

  if (!reportId || !namaPenerima || !barcodeWip || !variantId || !lokasiTujuan) {
    return { success: false, message: 'Data laporan, petugas QC, variasi tujuan, dan lokasi wajib diisi.' };
  }
  if (qtyTerima <= 0 || qtyLolos < 0 || qtyGagal < 0 || qtyLolos > qtyTerima || qtyGagal > qtyTerima) {
    return { success: false, message: 'Kuantitas pemeriksaan QC tidak valid.' };
  }
  if (!isValidLocationV2_(ss, lokasiTujuan)) {
    return { success: false, message: 'Lokasi rak tujuan tidak valid atau nonaktif.' };
  }

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (trxId && trxExistsV2_(shLog, trxId)) {
    return { success: false, message: 'Transaksi QC ini sudah pernah diproses.' };
  }

  var shLap = ensureSheetHeadersV2_(ss, IBS.SHEET_LAPORAN_JAHIT, IBS.HEADERS.LAPORAN_JAHIT);
  var lapRows = readObjectsV2_(shLap);
  var rep = lapRows.filter(function (r) { return textV2_(r.Report_ID) === reportId; })[0];
  if (!rep) return { success: false, message: 'Laporan jahit tidak ditemukan.' };
  if (upperV2_(rep.Status) !== 'PENDING') {
    return { success: false, message: 'Laporan ini sudah diproses sebelumnya (Status: ' + textV2_(rep.Status) + ').' };
  }

  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcodeWip);
  if (!product) return { success: false, message: 'Barcode produk tujuan tidak ditemukan.' };
  if (!product.active) return { success: false, message: 'Produk tujuan sedang nonaktif.' };

  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant || upperV2_(variant.barcode) !== upperV2_(barcodeWip)) {
    return { success: false, message: 'Variasi tidak cocok dengan barcode produk tujuan.' };
  }
  if (!variant.active) return { success: false, message: 'Variasi produk tujuan sedang nonaktif.' };

  var now = new Date();
  var stockSheet = ss.getSheetByName(IBS.SHEET_STOCK);
  var changes = [];

  try {
    if (qtyLolos > 0) {
      var existingStock = findStockRowV2_(stockSheet, variantId, lokasiTujuan);
      var currentQty = existingStock ? existingStock.qty : 0;
      changes.push(writeStockQuantityV2_(stockSheet, variantId, barcodeWip, lokasiTujuan, currentQty + qtyLolos, now));

      var display = product.name + (variant.label ? ' · ' + variant.label : '');
      var nomorWoLabel = textV2_(rep.Nomor_WO);
      var asalLabel = nomorWoLabel ? ('Produksi Jahit (' + nomorWoLabel + ')') : 'Produksi Jahit';
      shLog.appendRow([
        now, trxId || ('QC-' + reportId), 'Masuk', barcodeWip, asalLabel, lokasiTujuan, qtyLolos, 'SUCCESS', variantId,
        product.name, JSON.stringify(variant.attributes || {}), variant.label, display,
        namaPenerima + ' (QC: ' + textV2_(rep.Nama_Penjahit) + ')',
        variant.hpp || 0, 0, 0, 0
      ]);
      shLog.getRange(shLog.getLastRow(), 15, 1, 4).setNumberFormat('#,##0');
    }

    shLap.getRange(rep._row, 8).setValue('APPROVED');
    shLap.getRange(rep._row, 10, 1, 9).setValues([[
      qtyTerima, qtyLolos, qtyGagal, alasanGagal, barcodeWip, variantId, lokasiTujuan, namaPenerima, now
    ]]);

    var nomorWo = textV2_(rep.Nomor_WO);
    var woId = textV2_(rep.WO_ID);
    var woClosed = false;
    if (nomorWo || woId) {
      var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
      var shItem = ss.getSheetByName(IBS.SHEET_WO_JAHIT_ITEM);
      var woRows = shWo && shWo.getLastRow() >= 2 ? readObjectsV2_(shWo) : [];
      var wo = woRows.filter(function (w) {
        return (woId && textV2_(w.WO_ID) === woId) || upperV2_(w.Nomor_WO) === upperV2_(nomorWo);
      })[0];

      if (wo) {
        var resolvedWoId = textV2_(wo.WO_ID);
        var resolvedNoWo = textV2_(wo.Nomor_WO);
        var targetTotal = (shItem && shItem.getLastRow() >= 2 ? readObjectsV2_(shItem) : [])
          .filter(function (it) { return textV2_(it.WO_ID) === resolvedWoId; })
          .reduce(function (sum, it) { return sum + (Number(it.Target_Qty) || 0); }, 0);

        var approvedTotal = readObjectsV2_(shLap)
          .filter(function (r) {
            return upperV2_(r.Status) === 'APPROVED' &&
              ((resolvedWoId && textV2_(r.WO_ID) === resolvedWoId) || upperV2_(r.Nomor_WO) === upperV2_(resolvedNoWo));
          })
          .reduce(function (sum, r) { return sum + (Number(r.Qty_Lolos) || 0); }, 0);

        if (targetTotal > 0 && approvedTotal >= targetTotal) {
          shWo.getRange(wo._row, 3).setValue('SELESAI');
          woClosed = true;
        } else if (upperV2_(wo.Status) === 'DRAFT') {
          shWo.getRange(wo._row, 3).setValue('PROSES');
        }
      }
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'QC tersimpan (+' + qtyLolos + ' stok)' + (woClosed ? ' · Target WO tercapai & sesi ditutup!' : '.')
    };
  } catch (err) {
    rollbackStockChangesV2_(stockSheet, changes);
    throw err;
  }
}

// ==========================================
// MODUL STOCK OPNAME SESI
// ==========================================
function getOpnameSessions_(ss) {
  var sh = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
  if (!sh || sh.getLastRow() < 2) return { success: true, data: { sessions: [] } };

  var rows = readObjectsV2_(sh);
  var activeSessions = rows.filter(function (r) { return textV2_(r.Status) === 'ACTIVE'; }).map(function (r) {
    return {
      session_id: textV2_(r.Session_ID),
      nama_sesi: textV2_(r.Nama_Sesi),
      lokasi: textV2_(r.Lokasi),
      dibuat_oleh: textV2_(r.Dibuat_Oleh),
      waktu_mulai: r.Waktu_Mulai
    };
  });
  return { success: true, data: { sessions: activeSessions } };
}

function createOpnameSession_(ss, payload) {
  var name = textV2_(payload.nama_sesi);
  var loc = textV2_(payload.lokasi);
  var pelapor = textV2_(payload.nama_pelapor);

  if (!name || !loc || !pelapor) return { success: false, message: 'Nama sesi, lokasi, dan pelapor wajib diisi.' };
  if (!isValidLocationV2_(ss, loc)) return { success: false, message: 'Lokasi tidak valid.' };

  var sessionId = 'OPN-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000);
  var sh = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  sh.appendRow([sessionId, name, loc, 'ACTIVE', pelapor, new Date(), '']);
  SpreadsheetApp.flush();
  return { success: true, message: 'Sesi opname berhasil dibuat.', data: { session_id: sessionId } };
}

function submitOpnameCount_(ss, payload) {
  var sessionId = textV2_(payload.session_id);
  var barcode = textV2_(payload.barcode);
  var variantId = textV2_(payload.variant_id);
  var qty = Number(payload.qty);
  var pelapor = textV2_(payload.nama_pelapor);

  if (!sessionId || !barcode || !variantId || !pelapor || isNaN(qty) || qty < 0) return { success: false, message: 'Data tidak lengkap atau tidak valid.' };

  var shSesi = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId && textV2_(r.Status) === 'ACTIVE'; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan atau sudah selesai.' };

  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant) return { success: false, message: 'Variasi tidak ditemukan.' };

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var stockRow = findStockRowV2_(shStock, variantId, textV2_(sesi.Lokasi));
  var snapshotQty = stockRow ? stockRow.qty : 0;

  var shDetail = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var existing = detailRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId && textV2_(r.ID_Variasi) === variantId; })[0];

  var now = new Date();
  if (existing) {
    shDetail.getRange(existing._row, 7).setValue(qty);
    shDetail.getRange(existing._row, 8).setValue(pelapor);
    shDetail.getRange(existing._row, 9).setValue(now);
    shDetail.getRange(existing._row, 10).setValue('RECOUNT');
  } else {
    var detailId = sessionId + '-' + variantId;
    shDetail.appendRow([detailId, sessionId, barcode, variantId, variant.label, snapshotQty, qty, pelapor, now, 'COUNTED']);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Hasil hitung disimpan (Sistem: ' + snapshotQty + ', Fisik: ' + qty + ').' };
}

function approveOpnameSession_(ss, payload) {
  var sessionId = textV2_(payload.session_id);
  var pelapor = textV2_(payload.nama_pelapor);
  if (!sessionId || !pelapor) return { success: false, message: 'Session ID dan pelapor wajib diisi.' };

  var shSesi = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan.' };
  if (textV2_(sesi.Status) !== 'ACTIVE') return { success: false, message: 'Sesi sudah selesai.' };

  var loc = textV2_(sesi.Lokasi);
  var shDetail = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var counts = detailRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; });

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var now = new Date();
  var changes = [];
  var logRows = [];

  try {
    counts.forEach(function (c) {
      var variantId = textV2_(c.ID_Variasi);
      var barcode = textV2_(c.Barcode);
      var qtyFisik = Number(c.Qty_Fisik);
      var qtySistem = Number(c.Qty_Sistem_Snapshot);
      var diff = qtyFisik - qtySistem;

      if (diff !== 0) {
        changes.push(writeStockQuantityV2_(shStock, variantId, barcode, loc, qtyFisik, now));
        var trxId = 'OPN-ADJ-' + Utilities.getUuid().substring(0, 8);
        var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
        var variant = findVariantRowV2_(shVar, variantId);
        var display = (product ? product.name : barcode) + ' · ' + textV2_(c.Label_Variasi);
        var hppSatuan = variant ? variant.hpp : 0;
        logRows.push([now, trxId, 'Opname', barcode, loc, loc, diff, 'SUCCESS', variantId, product ? product.name : barcode, '', textV2_(c.Label_Variasi), display, pelapor, hppSatuan, 0, 0, 0]);
      }
    });

    if (logRows.length) {
      var start = shLog.getLastRow() + 1;
      shLog.getRange(start, 1, logRows.length, logRows[0].length).setValues(logRows);
      shLog.getRange(start, 15, logRows.length, 4).setNumberFormat('#,##0');
    }

    shSesi.getRange(sesi._row, 4).setValue('COMPLETED');
    shSesi.getRange(sesi._row, 7).setValue(now);

    SpreadsheetApp.flush();
    return { success: true, message: 'Sesi opname selesai. ' + logRows.length + ' penyesuaian stok dicatat.' };
  } catch (err) {
    rollbackStockChangesV2_(shStock, changes);
    throw err;
  }
}

// ==========================================
// DATA HELPER FUNCTIONS
// ==========================================
function findMasterV2_(sheet, barcode) {
  var rows = readObjectsV2_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (upperV2_(rows[i].Barcode) === upperV2_(barcode)) {
      return { row: rows[i]._row, barcode: textV2_(rows[i].Barcode), name: textV2_(rows[i].Nama), active: booleanV2_(rows[i].Status_Aktif, false) };
    }
  }
  return null;
}

function findVariantRowV2_(sheet, variantId) {
  var rows = readObjectsV2_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (textV2_(rows[i].ID_Variasi) === textV2_(variantId)) {
      var hpp = Number(rows[i].HPP) || 0;
      var hargaJual = Number(rows[i].Harga_Jual) || 0;
      return {
        row: rows[i]._row,
        id: textV2_(rows[i].ID_Variasi),
        barcode: textV2_(rows[i].Barcode_Induk),
        label: textV2_(rows[i].Label_Variasi),
        attributes: parseJsonObjectV2_(rows[i].Atribut_JSON),
        active: booleanV2_(rows[i].Status_Aktif, false),
        hpp: hpp,
        harga_jual: hargaJual
      };
    }
  }
  return null;
}

function findStockRowV2_(sheet, variantId, location) {
  var rows = readObjectsV2_(sheet);
  var found = null;
  rows.forEach(function (r) {
    if (textV2_(r.ID_Variasi) === textV2_(variantId) && textV2_(r.Lokasi_Rak) === textV2_(location)) {
      if (found) throw new Error('Duplikasi saldo ditemukan untuk ' + variantId + ' / ' + location + '. Periksa STOK_VARIASI.');
      found = { row: r._row, qty: Number(r.Kuantitas) || 0 };
    }
  });
  return found;
}

function writeStockQuantityV2_(sheet, variantId, barcode, location, qty, now) {
  var existing = findStockRowV2_(sheet, variantId, location);
  if (existing) {
    var old = sheet.getRange(existing.row, 4).getValue();
    var oldUpdated = sheet.getRange(existing.row, 5).getValue();
    sheet.getRange(existing.row, 4).setValue(qty);
    sheet.getRange(existing.row, 5).setValue(now);
    return { type: 'update', row: existing.row, oldQty: old, oldUpdated: oldUpdated };
  }
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 5).setValues([[variantId, barcode, location, qty, now]]);
  return { type: 'append', row: row };
}

function rollbackStockChangesV2_(sheet, changes) {
  for (var i = changes.length - 1; i >= 0; i--) {
    var c = changes[i];
    try {
      if (c.type === 'append') sheet.deleteRow(c.row);
      else if (c.type === 'update') { sheet.getRange(c.row, 4).setValue(c.oldQty); sheet.getRange(c.row, 5).setValue(c.oldUpdated); }
    } catch (e) { console.error('Rollback stock error: ' + e); }
  }
}

function readObjectsV2_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var headers = values[0].map(function (h) { return textV2_(h); });
  return values.slice(1).map(function (row, i) {
    var obj = { _row: i + 2 };
    headers.forEach(function (h, j) { if (h) obj[h] = row[j]; });
    return obj;
  });
}

function textV2_(val) { return val == null ? '' : String(val).trim(); }
function upperV2_(val) { return textV2_(val).toUpperCase(); }
function booleanV2_(val, fallback) {
  if (val === true || val === false) return val;
  var v = upperV2_(val);
  if (v === 'TRUE' || v === '1' || v === 'YA' || v === 'AKTIF') return true;
  if (v === 'FALSE' || v === '0' || v === 'TIDAK' || v === 'NONAKTIF') return false;
  return fallback;
}