/**
 * INA BASIC STORE - Warehouse V4.2 (Full Integration & Bug-Fixed)
 * Menggabungkan stabilitas V3.1 (Sesi Opname, Migrasi Stok, Rollback, Format Teks)
 * dengan modul Produksi & Keuangan (Pengiriman & Penerimaan Bahan, QC, Tarif Upah,
 * Slip Gaji, serta Kalkulasi HPP yang menggabungkan Cost Bahan + Cost Jahit).
 */

var IBS = {
  SHEET_MASTER: 'MASTER_PRODUK',
  SHEET_ATTRIBUTES: 'ATRIBUT_PRODUK',
  SHEET_VARIANTS: 'VARIASI_PRODUK',
  SHEET_STOCK: 'STOK_VARIASI',
  SHEET_LEGACY_STOCK: 'STOK_LOKASI',
  SHEET_LOG: 'OPERATION_LOG',
  SHEET_MIGRATION: 'MIGRASI_ALOKASI',
  SHEET_LOCATIONS: 'MASTER_LOKASI',
  SHEET_OPNAME_SESI: 'OPNAME_SESI',
  SHEET_OPNAME_DETAIL: 'OPNAME_DETAIL',
  SHEET_BAHAN: 'MASTER_BAHAN',
  SHEET_WO_JAHIT: 'WO_JAHIT',
  SHEET_WO_JAHIT_ITEM: 'WO_JAHIT_ITEM',
  SHEET_DELIVERY_BAHAN: 'PENGIRIMAN_BAHAN',
  SHEET_TARIF: 'MASTER_TARIF',
  SHEET_LAPORAN_JAHIT: 'LAPORAN_JAHIT',
  SHEET_PENERIMAAN_QC: 'PENERIMAAN_QC',
  SHEET_REKAP_UPAH: 'REKAP_UPAH',
  SHEET_HPP: 'HPP_PRODUKSI',
  LOCK_TIMEOUT_MS: 15000,
  HEADERS: {
    MASTER_PRODUK: ['Barcode', 'Nama', 'Status_Aktif', 'Created_At', 'Updated_At'],
    ATRIBUT_PRODUK: ['Barcode_Induk', 'Nama_Atribut', 'Tipe_Atribut', 'Nilai', 'Kode_Warna', 'Urutan_Atribut', 'Urutan_Nilai', 'Status_Aktif'],
    VARIASI_PRODUK: ['ID_Variasi', 'Barcode_Induk', 'Label_Variasi', 'Atribut_JSON', 'Status_Aktif', 'Created_At', 'Updated_At'],
    STOK_VARIASI: ['ID_Variasi', 'Barcode_Induk', 'Lokasi_Rak', 'Kuantitas', 'Updated_At'],
    OPERATION_LOG: ['Timestamp', 'ID_Transaksi', 'Aktivitas', 'Barcode', 'Lokasi_Asal', 'Lokasi_Tujuan', 'Qty', 'Status', 'ID_Variasi', 'Nama_Produk', 'Variasi_JSON', 'Label_Variasi', 'Produk_Tampilan', 'Nama_Pelapor'],
    MIGRASI_ALOKASI: ['Barcode_Induk', 'Lokasi_Rak', 'ID_Variasi', 'Kuantitas'],
    MASTER_LOKASI: ['Kode_Lokasi', 'Nama_Lokasi', 'Jenis_Lokasi', 'Status_Aktif', 'Created_At'],
    OPNAME_SESI: ['Session_ID', 'Nama_Sesi', 'Lokasi', 'Status', 'Dibuat_Oleh', 'Waktu_Mulai', 'Waktu_Selesai'],
    OPNAME_DETAIL: ['Detail_ID', 'Session_ID', 'Barcode', 'ID_Variasi', 'Label_Variasi', 'Qty_Sistem_Snapshot', 'Qty_Fisik', 'Nama_Pelapor', 'Waktu', 'Status_Hitung'],
    MASTER_BAHAN: ['Kode_Bahan', 'Nama_Bahan', 'Satuan', 'Estimasi_Cost', 'Status_Aktif', 'Created_At'],
    WO_JAHIT: ['WO_ID', 'Nomor_WO', 'Status', 'Tim_Label', 'Lokasi_Asal_Bahan', 'Lokasi_Tujuan', 'Tanggal_Target', 'Catatan', 'Dibuat_Oleh', 'Created_At'],
    WO_JAHIT_ITEM: ['Item_ID', 'WO_ID', 'Model_Warna_Ukuran', 'Target_Qty', 'Catatan'],
    PENGIRIMAN_BAHAN: ['Delivery_ID', 'WO_ID', 'Kode_Bahan', 'Qty_Rencana', 'Qty_Aktual_Kirim', 'Kurir', 'Penerima', 'Qty_Terima', 'Status', 'Token_QR', 'Created_At'],
    MASTER_TARIF: ['ID_Tarif', 'Kode_Model_WIP', 'Deskripsi_Model', 'Upah_Per_Pcs', 'Status_Aktif', 'Created_At'],
    LAPORAN_JAHIT: ['Report_ID', 'WO_ID', 'Nama_Penjahit', 'Tim_Penjahit', 'Qty_Bagus', 'Qty_Cacat', 'Status', 'Created_At'],
    PENERIMAAN_QC: ['QC_ID', 'Report_ID', 'WO_ID', 'Tim_Penjahit', 'Qty_Lolos', 'Qty_Gagal', 'ID_Variasi', 'Nilai_Upah', 'Status_Pembayaran', 'Slip_ID', 'Nama_Admin_QC', 'Created_At'],
    REKAP_UPAH: ['Slip_ID', 'Nama_Tim_Penjahit', 'Periode_Mulai', 'Periode_Akhir', 'Total_Qty_Lolos', 'Total_Upah', 'Status_Pembayaran', 'Tanggal_Bayar', 'Nama_Admin', 'Created_At'],
    HPP_PRODUKSI: ['HPP_ID', 'WO_ID', 'Total_Cost_Bahan', 'Total_Cost_Jahit', 'Qty_Lolos_QC', 'HPP_Per_Pcs', 'Created_At']
  }
};

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER, IBS.HEADERS.MASTER_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_ATTRIBUTES, IBS.HEADERS.ATRIBUT_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_VARIANTS, IBS.HEADERS.VARIASI_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_STOCK, IBS.HEADERS.STOK_VARIASI);
  ensureSheetHeadersV2_(ss, IBS.SHEET_LOG, IBS.HEADERS.OPERATION_LOG);
  ensureSheetHeadersV2_(ss, IBS.SHEET_MIGRATION, IBS.HEADERS.MIGRASI_ALOKASI);

  var locSheet = ensureSheetHeadersV2_(ss, IBS.SHEET_LOCATIONS, IBS.HEADERS.MASTER_LOKASI);
  cleanEmptyTrailingRowsV2_(locSheet);
  if (locSheet.getLastRow() === 1) {
    locSheet.appendRow(['RAK-UTAMA', 'Rak Utama', 'warehouse', true, new Date()]);
  }

  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  ensureSheetHeadersV2_(ss, IBS.SHEET_BAHAN, IBS.HEADERS.MASTER_BAHAN);
  ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT, IBS.HEADERS.WO_JAHIT);
  ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT_ITEM, IBS.HEADERS.WO_JAHIT_ITEM);
  ensureSheetHeadersV2_(ss, IBS.SHEET_DELIVERY_BAHAN, IBS.HEADERS.PENGIRIMAN_BAHAN);
  ensureSheetHeadersV2_(ss, IBS.SHEET_TARIF, IBS.HEADERS.MASTER_TARIF);
  ensureSheetHeadersV2_(ss, IBS.SHEET_LAPORAN_JAHIT, IBS.HEADERS.LAPORAN_JAHIT);
  ensureSheetHeadersV2_(ss, IBS.SHEET_PENERIMAAN_QC, IBS.HEADERS.PENERIMAAN_QC);
  ensureSheetHeadersV2_(ss, IBS.SHEET_REKAP_UPAH, IBS.HEADERS.REKAP_UPAH);
  ensureSheetHeadersV2_(ss, IBS.SHEET_HPP, IBS.HEADERS.HPP_PRODUKSI);

  var sheets = [
    IBS.SHEET_MASTER, IBS.SHEET_ATTRIBUTES, IBS.SHEET_VARIANTS, IBS.SHEET_STOCK,
    IBS.SHEET_LOG, IBS.SHEET_MIGRATION, IBS.SHEET_LOCATIONS, IBS.SHEET_OPNAME_SESI,
    IBS.SHEET_OPNAME_DETAIL, IBS.SHEET_BAHAN, IBS.SHEET_WO_JAHIT, IBS.SHEET_WO_JAHIT_ITEM,
    IBS.SHEET_DELIVERY_BAHAN, IBS.SHEET_TARIF, IBS.SHEET_LAPORAN_JAHIT, IBS.SHEET_PENERIMAAN_QC,
    IBS.SHEET_REKAP_UPAH, IBS.SHEET_HPP
  ];
  sheets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      cleanEmptyTrailingRowsV2_(sh);
      sh.setFrozenRows(1);
      var width = Math.max(sh.getLastColumn(), 1);
      sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d9eaf7');
    }
  });

  // Pertahankan format teks (@) agar barcode/kode berawalan nol tidak berubah menjadi angka
  ss.getSheetByName(IBS.SHEET_MASTER).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_ATTRIBUTES).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_VARIANTS).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_STOCK).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('D:F').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('I:I').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_MIGRATION).getRange('A:C').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL).getRange('B:D').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_BAHAN).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_WO_JAHIT).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_DELIVERY_BAHAN).getRange('A:C').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_TARIF).getRange('A:B').setNumberFormat('@');

  // Pasang validasi checkbox HANYA pada baris data yang sudah terisi (bukan 999 baris kosong)
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_MASTER), 3);
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_ATTRIBUTES), 8);
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), 5);
  applyCheckboxValidationV2_(locSheet, 4);
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_BAHAN), 5);
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_TARIF), 5);

  return true;
}

function countRealDataRowsV2_(sh) {
  if (!sh || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var count = 0;
  for (var i = 0; i < vals.length; i++) {
    var hasData = false;
    for (var j = 0; j < vals[i].length; j++) {
      var v = vals[i][j];
      if (v !== '' && v !== null && v !== false) {
        hasData = true;
        break;
      }
    }
    if (hasData) count++;
  }
  return count;
}

function cleanEmptyTrailingRowsV2_(sh) {
  if (!sh) return;
  var maxRows = sh.getMaxRows();
  var maxCols = sh.getMaxColumns();
  if (maxRows < 2 || maxCols < 1) return;

  var vals = sh.getRange(1, 1, maxRows, maxCols).getValues();
  var lastRealRow = 1;
  for (var i = vals.length - 1; i >= 1; i--) {
    var hasData = false;
    for (var j = 0; j < vals[i].length; j++) {
      var v = vals[i][j];
      if (v !== '' && v !== null && v !== false) {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      lastRealRow = i + 1;
      break;
    }
  }

  if (maxRows > lastRealRow) {
    var emptyRange = sh.getRange(lastRealRow + 1, 1, maxRows - lastRealRow, maxCols);
    emptyRange.clearContent();
    emptyRange.clearDataValidations();
  }
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
  var isExactMatch = current.length === headers.length && headers.every(function (h, idx) { return current[idx] === h; });

  // Jika sheet belum memiliki data nyata sama sekali dan headernya masih versi lama, reset ke struktur V4.2
  if (!isExactMatch && countRealDataRowsV2_(sh) === 0) {
    sh.getRange(1, 1, 1, Math.max(sh.getMaxColumns(), headers.length)).clearContent();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  headers.forEach(function (header) {
    if (current.indexOf(header) === -1) {
      current.push(header);
      sh.getRange(1, current.length).setValue(header);
    }
  });
  return sh;
}

function applyCheckboxValidationV2_(sheet, col, startRow, numRows) {
  if (!sheet) return;
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  if (startRow !== undefined && numRows !== undefined) {
    if (numRows > 0) sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule);
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, col, lastRow - 1, 1).setDataValidation(rule);
  }
}

function assertV2SheetsReady_(ss) {
  var required = [IBS.SHEET_MASTER, IBS.SHEET_VARIANTS, IBS.SHEET_STOCK, IBS.SHEET_LOG];
  required.forEach(function (name) {
    if (!ss.getSheetByName(name)) throw new Error('Sheet ' + name + ' belum tersedia. Jalankan fungsi setup() terlebih dahulu.');
  });
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

function doGet() {
  return jsonResponseV2_(true, 'Ina Basic Store Warehouse API V4.2 aktif.', { version: '4.2' });
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
    else return jsonResponseV2_(false, 'Aplikasi lama tidak dapat menulis transaksi ke struktur V4. Perbarui aplikasi ke versi baru.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- READ ENDPOINTS (Tanpa Lock) ---
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
    return getActiveWoList_(ss);
  }

  if (action === 'GET_BAHAN_LIST') {
    return getBahanList_(ss);
  }

  if (action === 'GET_DELIVERIES') {
    return getDeliveriesList_(ss, payload);
  }

  if (action === 'GET_DATA_FOR_QC') {
    return getDataForQc_(ss);
  }

  if (action === 'GET_UNPAID_UPAH') {
    return getUnpaidUpah_(ss);
  }

  // --- WRITE ENDPOINTS (Dengan LockService) ---
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
      case 'CREATE_BAHAN':
        result = createBahan_(ss, payload);
        break;
      case 'CREATE_DELIVERY_BAHAN':
        result = createDeliveryBahan_(ss, payload);
        break;
      case 'CONFIRM_DELIVERY_BAHAN':
        result = confirmDeliveryBahan_(ss, payload);
        break;
      case 'CREATE_TARIF':
        result = createTarif_(ss, payload);
        break;
      case 'SUBMIT_LAPORAN_JAHIT':
        result = submitLaporanJahit_(ss, payload);
        break;
      case 'SUBMIT_QC_RECEIPT':
        result = submitQcReceipt_(ss, payload);
        break;
      case 'GENERATE_PAY_SLIP':
        result = generatePaySlip_(ss, payload);
        break;
      case 'CALCULATE_HPP':
        result = calculateHpp_(ss, payload);
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
// MODUL INTI V2: PRODUK, VARIASI & TRANSAKSI GUDANG
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
    variantList.push({
      id: textV2_(r.ID_Variasi),
      label: textV2_(r.Label_Variasi),
      attributes: attrs,
      active: active,
      stocks: stocks
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
    appendRowsV2_(shMaster, [[barcode, name, productActive, now, now]], starts, 3);
    var attrRows = [];
    attributes.forEach(function (a, ai) {
      a.values.forEach(function (v, vi) {
        attrRows.push([barcode, a.name, a.type, v.value, v.hex || '', ai + 1, vi + 1, true]);
      });
    });
    if (attrRows.length) appendRowsV2_(shAttrs, attrRows, starts, 8);

    var variantRows = [];
    var stockRows = [];
    variants.forEach(function (v) {
      var id = newVariantIdV2_();
      var label = makeVariantLabelV2_(v.attributes, attributes);
      variantRows.push([id, barcode, label, JSON.stringify(v.attributes), v.active, now, now]);
      activeLocs.forEach(function (loc) { stockRows.push([id, barcode, loc.value, 0, now]); });
    });
    appendRowsV2_(shVariants, variantRows, starts, 5);
    appendRowsV2_(shStock, stockRows, starts);
    appendRowsV2_(shLog, [[now, trxId, 'ProdukBaru', barcode, '', '', 0, 'SUCCESS', '', name, '', '', name, textV2_(payload.nama_pelapor)]], starts);
    SpreadsheetApp.flush();
    return { success: true, message: 'Produk dan ' + variants.length + ' variasi berhasil didaftarkan. Stok awal setiap lokasi = 0.', data: { barcode: barcode, nama: name, jumlah_variasi: variants.length } };
  } catch (err) {
    rollbackAppendsV2_(starts);
    throw err;
  }
}

function addVariantsActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var trxId = textV2_(payload.transaction_id);
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
    if (newAttrRows.length) appendRowsV2_(shAttrs, newAttrRows, starts, 8);
    var vr = [], sr = [];
    variants.forEach(function (v) {
      var id = newVariantIdV2_();
      vr.push([id, barcode, makeVariantLabelV2_(v.attributes, mergedDefs), JSON.stringify(v.attributes), v.active, now, now]);
      activeLocs.forEach(function (loc) { sr.push([id, barcode, loc.value, 0, now]); });
    });
    appendRowsV2_(shVariants, vr, starts, 5);
    appendRowsV2_(shStock, sr, starts);
    appendRowsV2_(shLog, [[now, trxId, 'TambahVariasi', barcode, '', '', 0, 'SUCCESS', '', product.name, '', '', product.name, textV2_(payload.nama_pelapor)]], starts);
    SpreadsheetApp.flush();
    return { success: true, message: variants.length + ' variasi baru berhasil ditambahkan.' };
  } catch (err) {
    rollbackAppendsV2_(starts);
    changedVariantRows.forEach(function (x) { try { shVariants.getRange(x.row, 5).setValue(x.oldActive); shVariants.getRange(x.row, 7).setValue(x.oldUpdated); } catch (restoreError) {} });
    throw err;
  }
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
    log.appendRow([now, trxId, 'EditProduk', barcode, '', '', 0, 'SUCCESS', '', name || product.name, '', '', name || product.name, textV2_(payload.nama_pelapor)]);
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
    log.appendRow([now, trxId, 'StatusVariasi', barcode, '', '', 0, 'SUCCESS', variantId, product.name, JSON.stringify(row.attributes || {}), row.label, product.name + ' · ' + row.label, textV2_(payload.nama_pelapor)]);
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
    shLog.appendRow([now, trxId, action, barcode, from, to, qty, 'SUCCESS', variantId, product.name, JSON.stringify(variant.attributes || {}), variant.label, display, pelapor]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Transaksi ' + action.toLowerCase() + ' berhasil disimpan.', data: { barcode: barcode, nama: product.name, variant_id: variantId, variasi: variant.label } };
  } catch (err) {
    rollbackStockChangesV2_(stockSheet, changes);
    throw err;
  }
}

// ==========================================
// MODUL FASE 2: STOCK OPNAME SESI (PERBAIKAN SELISIH BERTANDA)
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
  var sh = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
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

  var shSesi = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId && textV2_(r.Status) === 'ACTIVE'; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan atau sudah selesai.' };

  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant) return { success: false, message: 'Variasi tidak ditemukan.' };

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var stockRow = findStockRowV2_(shStock, variantId, textV2_(sesi.Lokasi));
  var snapshotQty = stockRow ? stockRow.qty : 0;

  var shDetail = ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL);
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

  var shSesi = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan.' };
  if (textV2_(sesi.Status) !== 'ACTIVE') return { success: false, message: 'Sesi sudah selesai.' };

  var loc = textV2_(sesi.Lokasi);
  var shDetail = ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var counts = detailRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; });

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var now = new Date();
  var changes = [];
  var logRows = [];

  try {
    counts.forEach(function (c) {
      var variantId = textV2_(c.ID_Variasi);
      var barcode = textV2_(c.Barcode);
      var qtyFisik = Number(c.Qty_Fisik);
      // Bandingkan dengan stok aktual saat approve agar mutasi di tengah sesi tidak hilang
      var liveStockRow = findStockRowV2_(shStock, variantId, loc);
      var qtyAktualSistem = liveStockRow ? liveStockRow.qty : Number(c.Qty_Sistem_Snapshot);
      var diff = qtyFisik - qtyAktualSistem;

      if (diff !== 0) {
        changes.push(writeStockQuantityV2_(shStock, variantId, barcode, loc, qtyFisik, now));
        var trxId = 'OPN-ADJ-' + Utilities.getUuid().substring(0, 8);
        var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
        var display = (product ? product.name : barcode) + ' · ' + textV2_(c.Label_Variasi);
        var statusNote = diff > 0 ? ('SURPLUS (+' + diff + ')') : ('MINUS (' + diff + ')');
        logRows.push([now, trxId, 'Opname', barcode, loc, loc, diff, statusNote, variantId, product ? product.name : barcode, '', textV2_(c.Label_Variasi), display, pelapor]);
      }
    });

    if (logRows.length) {
      var start = shLog.getLastRow() + 1;
      shLog.getRange(start, 1, logRows.length, logRows[0].length).setValues(logRows);
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
// MODUL FASE 3: WORK ORDER, BAHAN & PENGIRIMAN
// ==========================================
function getActiveWoList_(ss) {
  try {
    var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
    if (!shWo || shWo.getLastRow() < 2) return jsonResponseV2_(true, 'Belum ada WO', { wo_list: [] });
    var rows = readObjectsV2_(shWo).filter(function (w) {
      var st = upperV2_(w.Status);
      return st !== 'COMPLETED' && st !== 'CANCELLED';
    }).map(function (w) {
      return {
        wo_id: textV2_(w.WO_ID),
        nomor_wo: textV2_(w.Nomor_WO),
        status: textV2_(w.Status),
        tim_label: textV2_(w.Tim_Label),
        lokasi_asal: textV2_(w.Lokasi_Asal_Bahan),
        lokasi_tujuan: textV2_(w.Lokasi_Tujuan),
        tanggal_target: w.Tanggal_Target
      };
    });
    return jsonResponseV2_(true, 'Daftar WO aktif.', { wo_list: rows });
  } catch (e) {
    return jsonResponseV2_(false, e.message);
  }
}

function getBahanList_(ss) {
  try {
    var sh = ss.getSheetByName(IBS.SHEET_BAHAN);
    if (!sh || sh.getLastRow() < 2) return jsonResponseV2_(true, 'Belum ada master bahan', { bahan_list: [] });
    var rows = readObjectsV2_(sh).filter(function (b) {
      return booleanV2_(b.Status_Aktif, true);
    }).map(function (b) {
      return {
        kode_bahan: textV2_(b.Kode_Bahan),
        nama_bahan: textV2_(b.Nama_Bahan),
        satuan: textV2_(b.Satuan),
        estimasi_cost: Number(b.Estimasi_Cost) || 0
      };
    });
    return jsonResponseV2_(true, 'Daftar master bahan.', { bahan_list: rows });
  } catch (e) {
    return jsonResponseV2_(false, e.message);
  }
}

function getDeliveriesList_(ss, payload) {
  try {
    var sh = ss.getSheetByName(IBS.SHEET_DELIVERY_BAHAN);
    if (!sh || sh.getLastRow() < 2) return jsonResponseV2_(true, 'Belum ada pengiriman', { deliveries: [] });
    var onlyTransit = payload && payload.only_in_transit === true;
    var rows = readObjectsV2_(sh).filter(function (d) {
      if (onlyTransit) return upperV2_(d.Status) === 'IN_TRANSIT';
      return true;
    }).map(function (d) {
      return {
        delivery_id: textV2_(d.Delivery_ID),
        wo_id: textV2_(d.WO_ID),
        kode_bahan: textV2_(d.Kode_Bahan),
        qty_rencana: Number(d.Qty_Rencana) || 0,
        qty_aktual_kirim: Number(d.Qty_Aktual_Kirim) || 0,
        kurir: textV2_(d.Kurir),
        penerima: textV2_(d.Penerima),
        qty_terima: Number(d.Qty_Terima) || 0,
        status: textV2_(d.Status),
        token_qr: textV2_(d.Token_QR)
      };
    });
    return jsonResponseV2_(true, 'Daftar pengiriman bahan.', { deliveries: rows });
  } catch (e) {
    return jsonResponseV2_(false, e.message);
  }
}

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
  if (!isValidLocationV2_(ss, lokasiAsal)) return { success: false, message: 'Lokasi asal bahan tidak valid.' };
  if (!items.length) return { success: false, message: 'Minimal harus ada 1 target item WIP pada Work Order.' };

  var woId = 'WOJ-' + Utilities.getUuid().split('-')[0].toUpperCase();
  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var shItem = ss.getSheetByName(IBS.SHEET_WO_JAHIT_ITEM);
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

function createBahan_(ss, payload) {
  var kode = textV2_(payload.kode_bahan);
  var nama = textV2_(payload.nama_bahan);
  var satuan = textV2_(payload.satuan) || 'kg';
  var cost = Number(payload.estimasi_cost) || 0;
  if (!kode || !nama) return { success: false, message: 'Kode dan nama bahan wajib diisi.' };

  var sh = ss.getSheetByName(IBS.SHEET_BAHAN);
  var rows = readObjectsV2_(sh);
  if (rows.some(function (r) { return upperV2_(r.Kode_Bahan) === upperV2_(kode); })) {
    return { success: false, message: 'Kode bahan sudah terdaftar.' };
  }

  var nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, 1, 1, 6).setValues([[kode, nama, satuan, cost, true, new Date()]]);
  applyCheckboxValidationV2_(sh, 5, nextRow, 1);
  SpreadsheetApp.flush();
  return { success: true, message: 'Master bahan berhasil ditambahkan.' };
}

function createDeliveryBahan_(ss, payload) {
  var woId = textV2_(payload.wo_id);
  var kodeBahan = textV2_(payload.kode_bahan);
  var qtyRencana = Number(payload.qty_rencana);
  var qtyAktual = payload.qty_aktual_kirim !== undefined ? Number(payload.qty_aktual_kirim) : qtyRencana;
  var kurir = textV2_(payload.nama_pelapor);

  if (!woId || !kodeBahan || isNaN(qtyAktual) || qtyAktual <= 0 || !kurir) {
    return { success: false, message: 'Data pengiriman tidak valid.' };
  }

  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var wo = readObjectsV2_(shWo).filter(function (w) { return textV2_(w.WO_ID) === woId || textV2_(w.Nomor_WO) === woId; })[0];
  if (!wo) return { success: false, message: 'Work Order tidak ditemukan.' };

  var deliveryId = 'DEL-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 900 + 100);
  var tokenQr = Utilities.getUuid().substring(0, 12).toUpperCase();
  var now = new Date();

  // Jika kode_bahan juga terdaftar sebagai produk ber-stok di MASTER_PRODUK, potong stoknya dari Lokasi_Asal_Bahan
  var prodBahan = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), kodeBahan);
  if (prodBahan && prodBahan.active) {
    var variants = readObjectsV2_(ss.getSheetByName(IBS.SHEET_VARIANTS)).filter(function (v) {
      return upperV2_(v.Barcode_Induk) === upperV2_(kodeBahan) && booleanV2_(v.Status_Aktif, false);
    });
    if (variants.length > 0) {
      var varId = textV2_(variants[0].ID_Variasi);
      var locAsal = textV2_(wo.Lokasi_Asal_Bahan);
      var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
      var stockRow = findStockRowV2_(shStock, varId, locAsal);
      var avail = stockRow ? stockRow.qty : 0;
      if (avail < qtyAktual) {
        return { success: false, message: 'Stok bahan ' + kodeBahan + ' di ' + locAsal + ' tidak cukup (Tersedia: ' + avail + ').' };
      }
      writeStockQuantityV2_(shStock, varId, kodeBahan, locAsal, avail - qtyAktual, now);
      ss.getSheetByName(IBS.SHEET_LOG).appendRow([now, 'TRX-' + deliveryId, 'Keluar', kodeBahan, locAsal, 'Produksi (' + textV2_(wo.Nomor_WO) + ')', qtyAktual, 'SUCCESS', varId, prodBahan.name, '', textV2_(variants[0].Label_Variasi), prodBahan.name, kurir]);
    }
  }

  var sh = ss.getSheetByName(IBS.SHEET_DELIVERY_BAHAN);
  sh.appendRow([deliveryId, textV2_(wo.WO_ID), kodeBahan, qtyRencana || qtyAktual, qtyAktual, kurir, '', 0, 'IN_TRANSIT', tokenQr, now]);

  // Update status WO dari DRAFT ke IN_PROGRESS
  if (upperV2_(wo.Status) === 'DRAFT') {
    shWo.getRange(wo._row, 3).setValue('IN_PROGRESS');
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Pengiriman bahan dicatat (' + deliveryId + ').', data: { delivery_id: deliveryId, token: tokenQr } };
}

function confirmDeliveryBahan_(ss, payload) {
  var code = textV2_(payload.delivery_id || payload.token_qr);
  var penerima = textV2_(payload.penerima);
  var qtyTerima = Number(payload.qty_terima);

  if (!code || !penerima || isNaN(qtyTerima) || qtyTerima <= 0) {
    return { success: false, message: 'Kode pengiriman, nama penerima, dan qty terima wajib diisi.' };
  }

  var sh = ss.getSheetByName(IBS.SHEET_DELIVERY_BAHAN);
  var deliv = readObjectsV2_(sh).filter(function (d) {
    return upperV2_(d.Delivery_ID) === upperV2_(code) || upperV2_(d.Token_QR) === upperV2_(code);
  })[0];

  if (!deliv) return { success: false, message: 'Data pengiriman bahan tidak ditemukan.' };
  if (upperV2_(deliv.Status) === 'RECEIVED') return { success: false, message: 'Pengiriman ini sudah pernah dikonfirmasi diterima.' };

  sh.getRange(deliv._row, 7).setValue(penerima);
  sh.getRange(deliv._row, 8).setValue(qtyTerima);
  sh.getRange(deliv._row, 9).setValue('RECEIVED');

  SpreadsheetApp.flush();
  return { success: true, message: 'Penerimaan bahan (' + textV2_(deliv.Kode_Bahan) + ' · ' + qtyTerima + ') berhasil dikonfirmasi.' };
}

// ==========================================
// MODUL FASE 4 & 7: TARIF, UPAH, QC & HPP
// ==========================================
function createTarif_(ss, payload) {
  var kode = textV2_(payload.kode_model);
  var desc = textV2_(payload.deskripsi);
  var upah = Number(payload.upah) || 0;
  if (!kode || upah <= 0) return { success: false, message: 'Kode model dan upah wajib diisi (> 0).' };

  var id = 'TRF-' + Math.floor(Math.random() * 900 + 100);
  var sh = ss.getSheetByName(IBS.SHEET_TARIF);
  var nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, 1, 1, 6).setValues([[id, kode, desc, upah, true, new Date()]]);
  applyCheckboxValidationV2_(sh, 5, nextRow, 1);
  SpreadsheetApp.flush();
  return { success: true, message: 'Tarif berhasil disimpan.' };
}

function submitLaporanJahit_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var noWo = textV2_(payload.nomor_wo);
  var penjahit = textV2_(payload.nama_penjahit);
  var bagus = Number(payload.qty_bagus) || 0;
  var cacat = Number(payload.qty_cacat) || 0;

  if (!noWo || !penjahit) return { success: false, message: 'Nomor WO dan nama penjahit wajib diisi.' };
  if (bagus + cacat <= 0) return { success: false, message: 'Total kuantitas laporan minimal 1.' };

  var shRep = ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT);
  if (trxId && readObjectsV2_(shRep).some(function (r) { return textV2_(r.Report_ID) === trxId; })) {
    return { success: false, message: 'Laporan ini sudah pernah dikirim.' };
  }

  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var wo = readObjectsV2_(shWo).filter(function (w) {
    return upperV2_(w.Nomor_WO) === upperV2_(noWo) || upperV2_(w.WO_ID) === upperV2_(noWo);
  })[0];
  if (!wo) return { success: false, message: 'Nomor WO tidak ditemukan di sistem.' };
  if (upperV2_(wo.Status) === 'COMPLETED') return { success: false, message: 'WO ini sudah ditutup (COMPLETED).' };

  var reportId = trxId || ('REP-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 900 + 100));
  shRep.appendRow([
    reportId, textV2_(wo.WO_ID), penjahit, textV2_(wo.Tim_Label), bagus, cacat, 'PENDING', new Date()
  ]);
  SpreadsheetApp.flush();
  return { success: true, message: 'Laporan jahit berhasil dikirim (' + reportId + ').', data: { report_id: reportId } };
}

function getDataForQc_(ss) {
  try {
    var reports = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT)).filter(function (r) { return upperV2_(r.Status) === 'PENDING'; });
    var tarifs = readObjectsV2_(ss.getSheetByName(IBS.SHEET_TARIF)).filter(function (t) { return booleanV2_(t.Status_Aktif, true); });
    return jsonResponseV2_(true, 'Data loaded', { pending_reports: reports, tarifs: tarifs });
  } catch (e) {
    return jsonResponseV2_(false, e.message);
  }
}

function submitQcReceipt_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var reportId = textV2_(payload.report_id);
  var tarifId = textV2_(payload.tarif_id);
  var variantId = textV2_(payload.variant_id);
  var barcode = textV2_(payload.barcode);
  var lolos = Number(payload.qty_lolos) || 0;
  var gagal = Number(payload.qty_gagal) || 0;
  var namaQc = textV2_(payload.nama_qc);

  if (!reportId || !tarifId || !variantId || !barcode || !namaQc) {
    return { success: false, message: 'Data QC tidak lengkap.' };
  }

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (trxId && trxExistsV2_(shLog, trxId)) return { success: false, message: 'Transaksi QC ini sudah pernah diproses.' };

  var shRep = ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT);
  var report = readObjectsV2_(shRep).filter(function (r) { return textV2_(r.Report_ID) === reportId; })[0];
  if (!report) return { success: false, message: 'Laporan tidak ditemukan.' };
  if (upperV2_(report.Status) !== 'PENDING') return { success: false, message: 'Laporan ini sudah selesai di-QC.' };

  var tarif = readObjectsV2_(ss.getSheetByName(IBS.SHEET_TARIF)).filter(function (t) { return textV2_(t.ID_Tarif) === tarifId; })[0];
  if (!tarif) return { success: false, message: 'Tarif upah tidak ditemukan.' };
  var upahPerPcs = Number(tarif.Upah_Per_Pcs) || 0;
  var totalUpah = lolos * upahPerPcs;

  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var wo = readObjectsV2_(shWo).filter(function (w) { return textV2_(w.WO_ID) === textV2_(report.WO_ID); })[0];
  var locTo = (wo && textV2_(wo.Lokasi_Tujuan)) ? textV2_(wo.Lokasi_Tujuan) : 'Rak Utama';
  var now = new Date();
  var qcId = 'QC-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 900 + 100);

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!product || !variant) return { success: false, message: 'Produk atau variasi WIP tidak valid.' };

  if (lolos > 0) {
    var stockRow = findStockRowV2_(shStock, variantId, locTo);
    var qtyAwal = stockRow ? stockRow.qty : 0;
    writeStockQuantityV2_(shStock, variantId, barcode, locTo, qtyAwal + lolos, now);

    var display = product.name + (variant.label ? ' · ' + variant.label : '');
    shLog.appendRow([now, trxId || ('TRX-' + qcId), 'Masuk', barcode, 'Produksi (' + textV2_(report.WO_ID) + ')', locTo, lolos, 'SUCCESS', variantId, product.name, JSON.stringify(variant.attributes || {}), variant.label, display, namaQc]);
  }

  ss.getSheetByName(IBS.SHEET_PENERIMAAN_QC).appendRow([
    qcId, reportId, textV2_(report.WO_ID), textV2_(report.Tim_Penjahit), lolos, gagal, variantId, totalUpah, 'UNPAID', '', namaQc, now
  ]);

  shRep.getRange(report._row, 7).setValue('REVIEWED');
  if (wo && upperV2_(wo.Status) !== 'COMPLETED') {
    shWo.getRange(wo._row, 3).setValue('QC_PROCESS');
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'QC Selesai (' + lolos + ' masuk ' + locTo + '). Upah: Rp ' + totalUpah.toLocaleString('id-ID') };
}

function getUnpaidUpah_(ss) {
  try {
    var qcData = readObjectsV2_(ss.getSheetByName(IBS.SHEET_PENERIMAAN_QC)).filter(function (q) { return upperV2_(q.Status_Pembayaran) === 'UNPAID'; });
    var grouped = {};
    qcData.forEach(function (q) {
      var tim = textV2_(q.Tim_Penjahit) || 'Tanpa Tim';
      if (!grouped[tim]) grouped[tim] = { tim: tim, count: 0, total_qty: 0, total_upah: 0 };
      grouped[tim].count += 1;
      grouped[tim].total_qty += (Number(q.Qty_Lolos) || 0);
      grouped[tim].total_upah += (Number(q.Nilai_Upah) || 0);
    });
    var result = Object.keys(grouped).map(function (k) { return grouped[k]; });
    return jsonResponseV2_(true, 'Data upah', { unpaid: result });
  } catch (e) {
    return jsonResponseV2_(false, e.message);
  }
}

function generatePaySlip_(ss, payload) {
  var tim = textV2_(payload.tim_penjahit);
  var pelapor = textV2_(payload.nama_admin);
  if (!tim || !pelapor) return { success: false, message: 'Tim penjahit dan nama admin wajib diisi.' };

  var shQc = ss.getSheetByName(IBS.SHEET_PENERIMAAN_QC);
  var toPay = readObjectsV2_(shQc).filter(function (q) {
    return upperV2_(q.Status_Pembayaran) === 'UNPAID' && (textV2_(q.Tim_Penjahit) || 'Tanpa Tim') === tim;
  });

  if (!toPay.length) return { success: false, message: 'Tidak ada tagihan untuk tim ini.' };

  var totalLolos = 0;
  var totalUpah = 0;
  var minDate = null;
  var maxDate = null;

  toPay.forEach(function (q) {
    totalLolos += (Number(q.Qty_Lolos) || 0);
    totalUpah += (Number(q.Nilai_Upah) || 0);
    var dt = q.Created_At instanceof Date ? q.Created_At : new Date(q.Created_At);
    if (!isNaN(dt.getTime())) {
      if (!minDate || dt < minDate) minDate = dt;
      if (!maxDate || dt > maxDate) maxDate = dt;
    }
  });

  var now = new Date();
  var slipId = 'SLIP-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 900 + 100);

  toPay.forEach(function (q) {
    shQc.getRange(q._row, 9).setValue('PAID');
    shQc.getRange(q._row, 10).setValue(slipId);
  });

  ss.getSheetByName(IBS.SHEET_REKAP_UPAH).appendRow([
    slipId, tim, minDate || now, maxDate || now, totalLolos, totalUpah, 'PAID', now, pelapor, now
  ]);

  SpreadsheetApp.flush();
  return { success: true, message: 'Slip gaji ' + slipId + ' berhasil dibuat & dilunasi (Rp ' + totalUpah.toLocaleString('id-ID') + ').' };
}

function calculateHpp_(ss, payload) {
  var noWo = textV2_(payload.nomor_wo);
  var overhead = Number(payload.biaya_overhead) || 0;
  if (!noWo) return { success: false, message: 'Nomor WO wajib diisi.' };

  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var wo = readObjectsV2_(shWo).filter(function (w) {
    return upperV2_(w.Nomor_WO) === upperV2_(noWo) || upperV2_(w.WO_ID) === upperV2_(noWo);
  })[0];
  if (!wo) return { success: false, message: 'WO tidak ditemukan.' };
  if (upperV2_(wo.Status) === 'COMPLETED') return { success: false, message: 'WO ini sudah pernah ditutup dan dihitung HPP-nya.' };

  var woId = textV2_(wo.WO_ID);

  // 1. Hitung Total Cost Bahan dari PENGIRIMAN_BAHAN x MASTER_BAHAN
  var bahanMap = {};
  readObjectsV2_(ss.getSheetByName(IBS.SHEET_BAHAN)).forEach(function (b) {
    bahanMap[upperV2_(b.Kode_Bahan)] = Number(b.Estimasi_Cost) || 0;
  });

  var deliveries = readObjectsV2_(ss.getSheetByName(IBS.SHEET_DELIVERY_BAHAN)).filter(function (d) {
    return textV2_(d.WO_ID) === woId;
  });

  var totalCostBahan = 0;
  deliveries.forEach(function (d) {
    var qtyPakai = (Number(d.Qty_Terima) > 0) ? Number(d.Qty_Terima) : (Number(d.Qty_Aktual_Kirim) || 0);
    var unitCost = bahanMap[upperV2_(d.Kode_Bahan)] || 0;
    totalCostBahan += (qtyPakai * unitCost);
  });

  // 2. Hitung Total Cost Jahit & Qty Lolos dari PENERIMAAN_QC
  var qcRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_PENERIMAAN_QC)).filter(function (q) {
    return textV2_(q.WO_ID) === woId;
  });
  var totalCostJahit = 0;
  var totalQtyLolos = 0;
  qcRows.forEach(function (q) {
    totalCostJahit += (Number(q.Nilai_Upah) || 0);
    totalQtyLolos += (Number(q.Qty_Lolos) || 0);
  });

  if (totalQtyLolos <= 0) return { success: false, message: 'Qty Lolos QC masih 0. Selesaikan proses QC terlebih dahulu.' };

  var totalModal = totalCostBahan + totalCostJahit + overhead;
  var hppPerPcs = Math.round(totalModal / totalQtyLolos);
  var now = new Date();
  var hppId = 'HPP-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 900 + 100);

  ss.getSheetByName(IBS.SHEET_HPP).appendRow([
    hppId, woId, totalCostBahan, totalCostJahit + overhead, totalQtyLolos, hppPerPcs, now
  ]);
  shWo.getRange(wo._row, 3).setValue('COMPLETED');

  SpreadsheetApp.flush();
  return {
    success: true,
    message: 'HPP Berhasil dihitung: Rp ' + hppPerPcs.toLocaleString('id-ID') + ' /pcs (Bahan: Rp ' + totalCostBahan.toLocaleString('id-ID') + ', Jahit: Rp ' + totalCostJahit.toLocaleString('id-ID') + ').',
    data: { hpp_id: hppId, hpp_per_pcs: hppPerPcs, total_cost_bahan: totalCostBahan, total_cost_jahit: totalCostJahit, qty_lolos: totalQtyLolos }
  };
}

// ==========================================
// HELPER UTILITIES & VALIDATION
// ==========================================
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
    return { attributes: normalized, active: booleanV2_(v.active, true) };
  });
}

function variantKeyV2_(attrs, definitions) {
  return definitions.map(function (d) { return d.name + '=' + textV2_(attrs[d.name]).toLowerCase(); }).join('|');
}

function makeVariantLabelV2_(attrs, definitions) {
  if (!definitions.length) return 'Standar';
  return definitions.map(function (d) { return d.name + ': ' + textV2_(attrs[d.name]); }).join(' · ');
}

function appendRowsV2_(sheet, rows, starts, checkboxCol) {
  if (!rows || !rows.length) return;
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  if (checkboxCol) {
    applyCheckboxValidationV2_(sheet, checkboxCol, start, rows.length);
  }
  starts.push({ sheet: sheet, start: start, count: rows.length });
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
      return {
        row: rows[i]._row,
        id: textV2_(rows[i].ID_Variasi),
        barcode: textV2_(rows[i].Barcode_Induk),
        label: textV2_(rows[i].Label_Variasi),
        attributes: parseJsonObjectV2_(rows[i].Atribut_JSON),
        active: booleanV2_(rows[i].Status_Aktif, false)
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
  var result = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var hasData = false;
    for (var k = 0; k < row.length; k++) {
      if (row[k] !== '' && row[k] !== null && row[k] !== false) {
        hasData = true;
        break;
      }
    }
    if (!hasData) continue;
    var obj = { _row: i + 1 };
    headers.forEach(function (h, j) { if (h) obj[h] = row[j]; });
    result.push(obj);
  }
  return result;
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

function migrateLegacyStockV2() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(IBS.LOCK_TIMEOUT_MS)) throw new Error('Server sedang sibuk.');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    assertV2SheetsReady_(ss);
    var shLegacy = ss.getSheetByName(IBS.SHEET_LEGACY_STOCK);
    var shMap = ss.getSheetByName(IBS.SHEET_MIGRATION);
    var shVariants = ss.getSheetByName(IBS.SHEET_VARIANTS);
    var shNewStock = ss.getSheetByName(IBS.SHEET_STOCK);
    if (!shLegacy || shLegacy.getLastRow() < 2) throw new Error('Tidak ada data lama pada STOK_LOKASI.');
    var existingV2Stock = readObjectsV2_(shNewStock);
    if (existingV2Stock.some(function (r) { return (Number(r.Kuantitas) || 0) !== 0; })) {
      throw new Error('STOK_VARIASI sudah memiliki saldo non-nol. Migrasi dibatalkan agar stok baru tidak tertimpa.');
    }

    var legacy = readObjectsV2_(shLegacy);
    var variantRows = readObjectsV2_(shVariants);
    var allocations = readObjectsV2_(shMap);
    var legacyTotals = {};
    legacy.forEach(function (r) {
      var key = upperV2_(r.Barcode) + '|' + textV2_(r.Lokasi_Rak);
      legacyTotals[key] = (legacyTotals[key] || 0) + (Number(r.Kuantitas) || 0);
    });
    var allocatedTotals = {};
    var allocationByKey = {};
    allocations.forEach(function (r) {
      var barcode = textV2_(r.Barcode_Induk);
      var location = textV2_(r.Lokasi_Rak);
      var variantId = textV2_(r.ID_Variasi);
      var qty = Number(r.Kuantitas);
      if (!barcode || !variantId || !isValidLocationV2_(ss, location) || !isFinite(qty) || qty < 0 || Math.floor(qty) !== qty) throw new Error('Ada baris MIGRASI_ALOKASI yang tidak valid pada baris ' + r._row + '.');
      var variant = variantRows.filter(function (v) { return textV2_(v.ID_Variasi) === variantId; })[0];
      if (!variant || upperV2_(variant.Barcode_Induk) !== upperV2_(barcode)) throw new Error('ID variasi ' + variantId + ' tidak cocok dengan barcode ' + barcode + '.');
      var key = upperV2_(barcode) + '|' + location;
      allocatedTotals[key] = (allocatedTotals[key] || 0) + qty;
      var stockKey = variantId + '|' + location;
      if (allocationByKey[stockKey] !== undefined) throw new Error('Alokasi ganda untuk ' + variantId + ' / ' + location + '.');
      allocationByKey[stockKey] = qty;
    });

    Object.keys(legacyTotals).forEach(function (key) {
      if ((allocatedTotals[key] || 0) !== legacyTotals[key]) throw new Error('Total alokasi tidak sama dengan stok lama pada ' + key + '. Lama=' + legacyTotals[key] + ', alokasi=' + (allocatedTotals[key] || 0));
    });
    Object.keys(allocatedTotals).forEach(function (key) {
      if (legacyTotals[key] === undefined && allocatedTotals[key] !== 0) throw new Error('Ditemukan alokasi tanpa pasangan stok lama: ' + key);
    });

    var activeLocs = getActiveLocationsV2_(ss);
    var output = [];
    variantRows.forEach(function (v) {
      activeLocs.forEach(function (loc) {
        var key = textV2_(v.ID_Variasi) + '|' + loc.value;
        output.push([textV2_(v.ID_Variasi), textV2_(v.Barcode_Induk), loc.value, allocationByKey[key] || 0, new Date()]);
      });
    });
    if (!output.length) throw new Error('Belum ada variasi produk untuk dimigrasikan.');
    if (shNewStock.getLastRow() > 1) shNewStock.getRange(2, 1, shNewStock.getLastRow() - 1, shNewStock.getLastColumn()).clearContent();
    shNewStock.getRange(2, 1, output.length, output[0].length).setValues(output);
    SpreadsheetApp.flush();
    return 'Migrasi selesai. ' + output.length + ' saldo variasi/lokasi dibuat. STOK_LOKASI lama tetap dipertahankan.';
  } finally {
    lock.releaseLock();
  }
}