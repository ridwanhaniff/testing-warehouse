/**
 * INA BASIC STORE - Dashboard & Reporting V5.2
 * Sinkron penuh dengan struktur sheet code.gs V5.2:
 * - Menu khusus di Google Sheets untuk refresh dashboard & setup cepat.
 * - Ringkasan nilai aset stok (HPP), omzet, dan estimasi profit penjualan.
 * - Rekap progres Work Order (WO) Jahit & status antrean QC.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ina Basic Store')
    .addItem('Refresh Dashboard Ringkasan', 'buildDashboardSummary')
    .addSeparator()
    .addItem('Jalankan Setup / Migrasi V5.2', 'setup')
    .addItem('Hapus Tab Lama Tidak Terpakai', 'hapusTabTidakTerpakai')
    .addToUi();
}

function buildDashboardSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  assertV2SheetsReady_(ss);

  var dashName = 'DASHBOARD';
  var dash = ss.getSheetByName(dashName);
  if (!dash) {
    dash = ss.insertSheet(dashName, 0);
  }
  dash.clear();
  dash.setHiddenGridlines(false);

  var masterRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_MASTER));
  var varRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_VARIANTS));
  var stockRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_STOCK));
  var logRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LOG));
  var woRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_WO_JAHIT));
  var woItemRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_WO_JAHIT_ITEM));
  var lapRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT));

  // Pemetaan master & variasi
  var masterMap = {};
  masterRows.forEach(function (m) {
    masterMap[upperV2_(m.Barcode)] = {
      nama: textV2_(m.Nama),
      aktif: booleanV2_(m.Status_Aktif, false)
    };
  });

  var variantMap = {};
  var activeVariantCount = 0;
  var missingPriceCount = 0;
  var totalMarginPctSum = 0;
  var pricedVariantCount = 0;

  varRows.forEach(function (v) {
    var vid = textV2_(v.ID_Variasi);
    var active = booleanV2_(v.Status_Aktif, false);
    var hpp = Number(v.HPP) || 0;
    var hj = Number(v.Harga_Jual) || 0;
    variantMap[vid] = {
      barcode: textV2_(v.Barcode_Induk),
      label: textV2_(v.Label_Variasi),
      active: active,
      hpp: hpp,
      harga_jual: hj
    };
    if (active) {
      activeVariantCount++;
      if (hpp <= 0 || hj <= 0) {
        missingPriceCount++;
      } else {
        totalMarginPctSum += ((hj - hpp) / hj);
        pricedVariantCount++;
      }
    }
  });

  // Hitung stok per variasi & nilai aset gudang
  var stockPerVariant = {};
  var totalStokUnit = 0;
  var totalNilaiAsetHpp = 0;
  var totalPotensiOmzet = 0;

  stockRows.forEach(function (s) {
    var vid = textV2_(s.ID_Variasi);
    var qty = Number(s.Kuantitas) || 0;
    var vInfo = variantMap[vid];
    if (!vInfo || !vInfo.active) return;

    stockPerVariant[vid] = (stockPerVariant[vid] || 0) + qty;
    if (qty > 0) {
      totalStokUnit += qty;
      totalNilaiAsetHpp += qty * vInfo.hpp;
      totalPotensiOmzet += qty * vInfo.harga_jual;
    }
  });

  // Hitung penjualan & transaksi bulan berjalan
  var now = new Date();
  var curMonth = now.getMonth();
  var curYear = now.getFullYear();
  var qtyKeluarBulanIni = 0;
  var omzetBulanIni = 0;
  var profitBulanIni = 0;

  logRows.forEach(function (lg) {
    if (upperV2_(lg.Aktivitas) !== 'KELUAR' || upperV2_(lg.Status) !== 'SUCCESS') return;
    var ts = lg.Timestamp instanceof Date ? lg.Timestamp : new Date(lg.Timestamp);
    if (!isNaN(ts.getTime()) && ts.getMonth() === curMonth && ts.getFullYear() === curYear) {
      qtyKeluarBulanIni += Number(lg.Qty) || 0;
      omzetBulanIni += Number(lg.Total_Omzet) || 0;
      profitBulanIni += Number(lg.Total_Profit) || 0;
    }
  });

  // Hitung antrean QC & WO aktif
  var pendingQcCount = lapRows.filter(function (r) { return upperV2_(r.Status) === 'PENDING'; }).length;
  var activeWoCount = woRows.filter(function (w) {
    var st = upperV2_(w.Status);
    return st !== 'SELESAI' && st !== 'BATAL' && st !== 'CLOSED';
  }).length;
  var avgMargin = pricedVariantCount > 0 ? (totalMarginPctSum / pricedVariantCount) : 0;

  // ==========================================
  // RENDER HEADER & KPI UTAMA
  // ==========================================
  dash.getRange('A1:G1').merge().setValue('INA BASIC STORE · DASHBOARD GUDANG & PRODUKSI (V5.2)')
    .setFontSize(14).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  dash.setRowHeight(1, 34);

  dash.getRange('A2:G2').merge().setValue('Terakhir diperbarui: ' + Utilities.formatDate(now, 'Asia/Jakarta', 'dd MMM yyyy HH:mm:ss') + ' WIB')
    .setFontSize(10).setFontColor('#4b5563').setBackground('#f3f4f6');

  var kpiHeaders = [
    ['Total Stok Aktif (Pcs)', 'Nilai Aset Gudang (HPP)', 'Potensi Omzet Stok', 'Omzet Keluar Bulan Ini', 'Estimasi Profit Bulan Ini', 'Rata-rata Margin Aktif', 'Antrean QC / WO Aktif'],
    [totalStokUnit, totalNilaiAsetHpp, totalPotensiOmzet, omzetBulanIni, profitBulanIni, avgMargin, pendingQcCount + ' QC / ' + activeWoCount + ' WO']
  ];
  dash.getRange(4, 1, 2, 7).setValues(kpiHeaders);
  dash.getRange(4, 1, 1, 7).setFontWeight('bold').setFontSize(10).setBackground('#dbeafe').setFontColor('#1e3a8a').setHorizontalAlignment('center');
  dash.getRange(5, 1, 1, 7).setFontWeight('bold').setFontSize(12).setBackground('#eff6ff').setHorizontalAlignment('center');
  dash.getRange(5, 1).setNumberFormat('#,##0');
  dash.getRange(5, 2, 1, 4).setNumberFormat('Rp #,##0');
  dash.getRange(5, 6).setNumberFormat('0.0%');
  dash.setRowHeight(5, 28);

  // ==========================================
  // TABEL 1: REKAP WORK ORDER JAHIT AKTIF
  // ==========================================
  var rowCursor = 7;
  dash.getRange(rowCursor, 1, 1, 7).merge().setValue('1. MONITORING WORK ORDER (WO) JAHIT AKTIF')
    .setFontWeight('bold').setBackground('#e5e7eb');
  rowCursor++;

  var woTableHeaders = ['Nomor WO', 'Tim Penjahit', 'Rak Tujuan', 'Target (Pcs)', 'Lolos QC (Pcs)', 'Menunggu QC (Pcs)', 'Sisa Target (Pcs)'];
  dash.getRange(rowCursor, 1, 1, 7).setValues([woTableHeaders]).setFontWeight('bold').setBackground('#f9fafb');
  rowCursor++;

  var woDataRows = [];
  woRows.forEach(function (w) {
    var st = upperV2_(w.Status);
    if (st === 'SELESAI' || st === 'BATAL' || st === 'CLOSED') return;

    var woId = textV2_(w.WO_ID);
    var noWo = textV2_(w.Nomor_WO);
    var target = woItemRows.filter(function (it) { return textV2_(it.WO_ID) === woId; })
      .reduce(function (s, it) { return s + (Number(it.Target_Qty) || 0); }, 0);
    var approved = lapRows.filter(function (lp) {
      return upperV2_(lp.Status) === 'APPROVED' && (textV2_(lp.WO_ID) === woId || upperV2_(lp.Nomor_WO) === upperV2_(noWo));
    }).reduce(function (s, lp) { return s + (Number(lp.Qty_Lolos) || 0); }, 0);
    var pending = lapRows.filter(function (lp) {
      return upperV2_(lp.Status) === 'PENDING' && (textV2_(lp.WO_ID) === woId || upperV2_(lp.Nomor_WO) === upperV2_(noWo));
    }).reduce(function (s, lp) { return s + (Number(lp.Qty_Bagus) || 0); }, 0);

    woDataRows.push([noWo, textV2_(w.Tim_Label), textV2_(w.Lokasi_Tujuan), target, approved, pending, Math.max(0, target - approved - pending)]);
  });

  if (woDataRows.length) {
    dash.getRange(rowCursor, 1, woDataRows.length, 7).setValues(woDataRows);
    dash.getRange(rowCursor, 4, woDataRows.length, 4).setNumberFormat('#,##0');
    rowCursor += woDataRows.length;
  } else {
    dash.getRange(rowCursor, 1, 1, 7).merge().setValue('Tidak ada Work Order aktif saat ini.').setFontStyle('italic').setFontColor('#6b7280');
    rowCursor++;
  }

  // ==========================================
  // TABEL 2: PERINGATAN STOK MENIPIS (<= 5 PCS)
  // ==========================================
  rowCursor += 1;
  dash.getRange(rowCursor, 1, 1, 7).merge().setValue('2. PERINGATAN STOK MENIPIS (<= 5 PCS) & KELENGKAPAN HARGA')
    .setFontWeight('bold').setBackground('#fef3c7').setFontColor('#92400e');
  rowCursor++;

  var lowHeaders = ['Barcode', 'Nama Produk', 'Variasi', 'Total Stok', 'HPP Satuan', 'Harga Jual', 'Status Data'];
  dash.getRange(rowCursor, 1, 1, 7).setValues([lowHeaders]).setFontWeight('bold').setBackground('#f9fafb');
  rowCursor++;

  var lowRows = [];
  Object.keys(variantMap).forEach(function (vid) {
    var v = variantMap[vid];
    if (!v.active) return;
    var m = masterMap[upperV2_(v.barcode)];
    if (m && !m.aktif) return;

    var qty = stockPerVariant[vid] || 0;
    var noPrice = (v.hpp <= 0 || v.harga_jual <= 0);
    if (qty <= 5 || noPrice) {
      var statusNote = [];
      if (qty <= 5) statusNote.push('Stok Menipis');
      if (noPrice) statusNote.push('HPP/Harga Belum Diisi');
      lowRows.push([v.barcode, m ? m.nama : v.barcode, v.label, qty, v.hpp, v.harga_jual, statusNote.join(' · ')]);
    }
  });

  lowRows.sort(function (a, b) { return a[3] - b[3]; });
  var displayLow = lowRows.slice(0, 30);

  if (displayLow.length) {
    dash.getRange(rowCursor, 1, displayLow.length, 7).setValues(displayLow);
    dash.getRange(rowCursor, 1, displayLow.length, 1).setNumberFormat('@');
    dash.getRange(rowCursor, 4, displayLow.length, 1).setNumberFormat('#,##0');
    dash.getRange(rowCursor, 5, displayLow.length, 2).setNumberFormat('Rp #,##0');
    rowCursor += displayLow.length;
  } else {
    dash.getRange(rowCursor, 1, 1, 7).merge().setValue('Semua variasi aktif memiliki stok > 5 pcs dan HPP/Harga Jual lengkap.').setFontColor('#059669');
    rowCursor++;
  }

  dash.autoResizeColumns(1, 7);
  SpreadsheetApp.flush();
  try {
    ss.toast('Dashboard ringkasan V5.2 berhasil diperbarui.', 'Ina Basic Store', 4);
  } catch (e) {}
}