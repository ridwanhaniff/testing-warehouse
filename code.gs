/**
 * INA BASIC STORE - Warehouse V3 (kandidat M0)
 * Perubahan M0 dari code_1.gs: jalur SAVE_TRANSACTION 'Opname' dihapus.
 * Stock opname HANYA lewat sesi (CREATE_OPNAME_SESSION / SUBMIT_OPNAME_COUNT / APPROVE_OPNAME_SESSION).
 * Stok tidak berubah sebelum sesi opname disetujui.
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
    PENGIRIMAN_BAHAN: ['Delivery_ID', 'WO_ID', 'Kode_Bahan', 'Qty_Rencana', 'Qty_Aktual_Kirim', 'Kurir', 'Penerima', 'Qty_Terima', 'Status', 'Token_QR', 'Created_At']
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
  if (locSheet.getLastRow() === 1) {
    locSheet.appendRow(['RAK-UTAMA', 'Rak Utama', 'warehouse', true, new Date()]);
  }

  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  
  // Fase 3 Sheets setup
  ensureSheetHeadersV2_(ss, IBS.SHEET_BAHAN, IBS.HEADERS.MASTER_BAHAN);
  ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT, IBS.HEADERS.WO_JAHIT);
  ensureSheetHeadersV2_(ss, IBS.SHEET_WO_JAHIT_ITEM, IBS.HEADERS.WO_JAHIT_ITEM);
  ensureSheetHeadersV2_(ss, IBS.SHEET_DELIVERY_BAHAN, IBS.HEADERS.PENGIRIMAN_BAHAN);

  var sheets = [
    IBS.SHEET_MASTER, IBS.SHEET_ATTRIBUTES, IBS.SHEET_VARIANTS, IBS.SHEET_STOCK, 
    IBS.SHEET_LOG, IBS.SHEET_MIGRATION, IBS.SHEET_LOCATIONS, IBS.SHEET_OPNAME_SESI, 
    IBS.SHEET_OPNAME_DETAIL, IBS.SHEET_BAHAN, IBS.SHEET_WO_JAHIT, IBS.SHEET_WO_JAHIT_ITEM, IBS.SHEET_DELIVERY_BAHAN
  ];
  sheets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      sh.setFrozenRows(1);
      var width = Math.max(sh.getLastColumn(), 1);
      sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d9eaf7');
    }
  });

  // Pertahankan format teks (@) agar barcode berawalan nol tidak berubah menjadi angka
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

  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_MASTER), 3);
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_ATTRIBUTES), 8);
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), 5);
  applyCheckboxValidationV2_(locSheet, 4);
  applyCheckboxValidationV2_(ss.getSheetByName(IBS.SHEET_BAHAN), 5);

  return true;
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

function applyCheckboxValidationV2_(sheet, col) {
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  var range = sheet.getRange(2, col, rows, 1);
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
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
  rows.forEach(function(r) {
    if (booleanV2_(r.Status_Aktif, true)) {
       active.push({ value: textV2_(r.Nama_Lokasi), label: textV2_(r.Nama_Lokasi), kind: textV2_(r.Jenis_Lokasi) || 'warehouse' });
    }
  });
  return active.length ? active : [{ value: 'Rak Utama', label: 'Rak Utama', kind: 'warehouse' }];
}

function isValidLocationV2_(ss, value) { 
  var locs = getActiveLocationsV2_(ss);
  return locs.some(function(l) { return l.value === textV2_(value); });
}

function doGet() {
  return jsonResponseV2_(true, 'Ina Basic Store Warehouse API V3 aktif.', { version: '3.1' });
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
    else return jsonResponseV2_(false, 'Aplikasi lama tidak dapat menulis transaksi ke struktur V3. Perbarui aplikasi ke versi baru.');
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
    appendRowsV2_(shMaster, [[barcode, name, productActive, now, now]], starts);
    var attrRows = [];
    attributes.forEach(function (a, ai) {
      a.values.forEach(function (v, vi) {
        attrRows.push([barcode, a.name, a.type, v.value, v.hex || '', ai + 1, vi + 1, true]);
      });
    });
    if (attrRows.length) appendRowsV2_(shAttrs, attrRows, starts);

    var variantRows = [];
    var stockRows = [];
    variants.forEach(function (v) {
      var id = newVariantIdV2_();
      var label = makeVariantLabelV2_(v.attributes, attributes);
      variantRows.push([id, barcode, label, JSON.stringify(v.attributes), v.active, now, now]);
      activeLocs.forEach(function (loc) { stockRows.push([id, barcode, loc.value, 0, now]); });
    });
    appendRowsV2_(shVariants, variantRows, starts);
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
    if (newAttrRows.length) appendRowsV2_(shAttrs, newAttrRows, starts);
    var vr = [], sr = [];
    variants.forEach(function (v) {
      var id = newVariantIdV2_();
      vr.push([id, barcode, makeVariantLabelV2_(v.attributes, mergedDefs), JSON.stringify(v.attributes), v.active, now, now]);
      activeLocs.forEach(function (loc) { sr.push([id, barcode, loc.value, 0, now]); });
    });
    appendRowsV2_(shVariants, vr, starts);
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

function appendRowsV2_(sheet, rows, starts) {
  if (!rows || !rows.length) return;
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
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

// ==========================================
// MODUL FASE 3: WORK ORDER JAHIT & MASTER BAHAN
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
  if (!isValidLocationV2_(ss, lokasiAsal)) return { success: false, message: 'Lokasi asal bahan tidak valid.' };
  if (!items.length) return { success: false, message: 'Minimal harus ada 1 target item WIP pada Work Order.' };

  var woId = 'WOJ-' + Utilities.getUuid().split('-')[0].toUpperCase();
  var shWo = ss.getSheetByName(IBS.SHEET_WO_JAHIT);
  var shItem = ss.getSheetByName(IBS.SHEET_WO_JAHIT_ITEM);
  var now = new Date();

  shWo.appendRow([woId, nomorWo, 'DRAFT', timLabel, lokasiAsal, lokasiTujuan, tanggalTarget, catatan, pelapor, now]);

  var itemRows = [];
  items.forEach(function(item, idx) {
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

// ==========================================
// MODUL FASE 2: STOCK OPNAME SESI
// ==========================================
function getOpnameSessions_(ss) {
  var sh = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
  if (!sh || sh.getLastRow() < 2) return { success: true, data: { sessions: [] } };
  
  var rows = readObjectsV2_(sh);
  var activeSessions = rows.filter(function(r) { return textV2_(r.Status) === 'ACTIVE'; }).map(function(r) {
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
  var sesi = sesiRows.filter(function(r) { return textV2_(r.Session_ID) === sessionId && textV2_(r.Status) === 'ACTIVE'; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan atau sudah selesai.' };
  
  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant) return { success: false, message: 'Variasi tidak ditemukan.' };
  
  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var stockRow = findStockRowV2_(shStock, variantId, textV2_(sesi.Lokasi));
  var snapshotQty = stockRow ? stockRow.qty : 0;
  
  var shDetail = ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var existing = detailRows.filter(function(r) { return textV2_(r.Session_ID) === sessionId && textV2_(r.ID_Variasi) === variantId; })[0];
  
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
  return { success: true, message: 'Hasil hitung disimpan (Sistem: '+snapshotQty+', Fisik: '+qty+').' };
}

function approveOpnameSession_(ss, payload) {
  var sessionId = textV2_(payload.session_id);
  var pelapor = textV2_(payload.nama_pelapor);
  if (!sessionId || !pelapor) return { success: false, message: 'Session ID dan pelapor wajib diisi.' };
  
  var shSesi = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function(r) { return textV2_(r.Session_ID) === sessionId; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan.' };
  if (textV2_(sesi.Status) !== 'ACTIVE') return { success: false, message: 'Sesi sudah selesai.' };
  
  var loc = textV2_(sesi.Lokasi);
  var shDetail = ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var counts = detailRows.filter(function(r) { return textV2_(r.Session_ID) === sessionId; });
  
  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var now = new Date();
  var changes = [];
  var logRows = [];
  
  try {
    counts.forEach(function(c) {
      var variantId = textV2_(c.ID_Variasi);
      var barcode = textV2_(c.Barcode);
      var qtyFisik = Number(c.Qty_Fisik);
      var qtySistem = Number(c.Qty_Sistem_Snapshot);
      var diff = qtyFisik - qtySistem;
      
      if (diff !== 0) { 
        changes.push(writeStockQuantityV2_(shStock, variantId, barcode, loc, qtyFisik, now));
        var trxId = 'OPN-ADJ-' + Utilities.getUuid().substring(0,8);
        var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
        var display = (product ? product.name : barcode) + ' · ' + textV2_(c.Label_Variasi);
        logRows.push([now, trxId, 'Opname', barcode, loc, loc, Math.abs(diff), 'SUCCESS', variantId, product ? product.name : barcode, '', textV2_(c.Label_Variasi), display, pelapor]);
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
// MODUL FASE 3: MASTER BAHAN & PENGIRIMAN
// ==========================================
function createBahan_(ss, payload) {
  var kode = textV2_(payload.kode_bahan);
  var nama = textV2_(payload.nama_bahan);
  var satuan = textV2_(payload.satuan) || 'kg';
  var cost = Number(payload.estimasi_cost) || 0;
  if (!kode || !nama) return { success: false, message: 'Kode dan nama bahan wajib diisi.' };
  
  var sh = ss.getSheetByName(IBS.SHEET_BAHAN);
  var rows = readObjectsV2_(sh);
  if (rows.some(function(r) { return upperV2_(r.Kode_Bahan) === upperV2_(kode); })) {
    return { success: false, message: 'Kode bahan sudah terdaftar.' };
  }
  
  sh.appendRow([kode, nama, satuan, cost, true, new Date()]);
  SpreadsheetApp.flush();
  return { success: true, message: 'Master bahan berhasil ditambahkan.' };
}

function createDeliveryBahan_(ss, payload) {
  var woId = textV2_(payload.wo_id);
  var kodeBahan = textV2_(payload.kode_bahan);
  var qtyRencana = Number(payload.qty_rencana);
  var kurir = textV2_(payload.nama_pelapor); 
  
  if (!woId || !kodeBahan || isNaN(qtyRencana) || qtyRencana <= 0 || !kurir) {
    return { success: false, message: 'Data pengiriman tidak valid.' };
  }
  
  var deliveryId = 'DEL-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 900 + 100);
  var tokenQr = Utilities.getUuid();
  
  var sh = ss.getSheetByName(IBS.SHEET_DELIVERY_BAHAN);
  sh.appendRow([deliveryId, woId, kodeBahan, qtyRencana, qtyRencana, kurir, '', 0, 'IN_TRANSIT', tokenQr, new Date()]);
  SpreadsheetApp.flush();
  return { success: true, message: 'Pengiriman bahan dicatat.', data: { delivery_id: deliveryId, token: tokenQr } };
}

// Data Helper Functions (Diperlukan agar tidak terjadi error missing method)
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