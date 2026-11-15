/**
 * INA BASIC STORE - Executive Dashboard & Reporting V6.0
 * Terintegrasi penuh dengan arsitektur Code.gs V6.0:
 * - Menggantikan referensi WO_JAHIT lama dengan modul Gudang Penjahit (MASTER_TAILOR, STOK_TAILOR, LOG_TAILOR).
 * - Kartu KPI Eksekutif 2x4 (Gudang Utama, Omzet/Profit Bulan Ini, Saldo Penjahit, & Tagihan Upah UNPAID).
 * - Rekap Kinerja & Tagihan Upah per Penjahit.
 * - Rekap Nilai Aset Stok per Lokasi Rak.
 * - Top 10 Produk & Variasi Terlaris Bulan Ini.
 * - Peringatan Stok Menipis (<= 5 Pcs) & Kelengkapan Data Harga (HPP / Harga Jual).
 * - Log Aktivitas Gudang & QC Terbaru.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ina Basic Store')
    .addItem('Refresh Dashboard Ringkasan', 'buildDashboardSummary')
    .addSeparator()
    .addItem('Jalankan Setup / Migrasi V6.0', 'setup')
    .addItem('Jalankan Simulasi Pengujian V6.0', 'testSimulasiV6')
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
  dash.setHiddenGridlines(true);
  dash.setTabColor('#0f172a');

  // ==========================================
  // 1. PENGAMBILAN & AGREGASI DATA V6.0
  // ==========================================
  var masterRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_MASTER));
  var varRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_VARIANTS));
  var stockRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_STOCK));
  var logRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LOG));
  var locRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LOCATIONS));
  var opnameRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_OPNAME_SESI));
  var masterTailorRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_MASTER_TAILOR));
  var stokTailorRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_STOK_TAILOR));
  var logTailorRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LOG_TAILOR));

  var now = new Date();
  var curMonth = now.getMonth();
  var curYear = now.getFullYear();
  var periodeLabel = Utilities.formatDate(now, 'Asia/Jakarta', 'MMMM yyyy');

  // Pemetaan Master Produk
  var masterMap = {};
  var activeSkuCount = 0;
  masterRows.forEach(function (m) {
    var bCode = upperV2_(m.Barcode);
    var isAct = booleanV2_(m.Status_Aktif, false);
    masterMap[bCode] = {
      barcode: textV2_(m.Barcode),
      nama: textV2_(m.Nama),
      aktif: isAct
    };
    if (isAct) activeSkuCount++;
  });

  // Pemetaan Variasi Produk & Kelengkapan Harga
  var variantMap = {};
  var activeVariantCount = 0;
  var missingPriceCount = 0;
  var totalMarginPctSum = 0;
  var pricedVariantCount = 0;

  varRows.forEach(function (v) {
    var vid = textV2_(v.ID_Variasi);
    var bCode = textV2_(v.Barcode_Induk);
    var mInfo = masterMap[upperV2_(bCode)];
    var active = booleanV2_(v.Status_Aktif, false) && (!mInfo || mInfo.aktif);
    var hpp = Number(v.HPP) || 0;
    var hj = Number(v.Harga_Jual) || 0;

    variantMap[vid] = {
      id: vid,
      barcode: bCode,
      label: textV2_(v.Label_Variasi) || 'Standar',
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

  // Pemetaan Lokasi Rak
  var locMetaMap = {};
  locRows.forEach(function (l) {
    var namaLoc = textV2_(l.Nama_Lokasi);
    if (namaLoc) {
      locMetaMap[namaLoc] = {
        jenis: textV2_(l.Jenis_Lokasi) || 'warehouse',
        aktif: booleanV2_(l.Status_Aktif, true)
      };
    }
  });

  // Agregasi Stok Gudang Utama per Variasi & per Lokasi Rak
  var stockPerVariant = {};
  var stockPerLocation = {};
  var totalStokUnit = 0;
  var totalNilaiAsetHpp = 0;
  var totalPotensiOmzet = 0;

  stockRows.forEach(function (s) {
    var vid = textV2_(s.ID_Variasi);
    var rak = textV2_(s.Lokasi_Rak) || 'Rak Utama';
    var qty = Number(s.Kuantitas) || 0;
    var vInfo = variantMap[vid];
    if (!vInfo || !vInfo.active) return;

    stockPerVariant[vid] = (stockPerVariant[vid] || 0) + qty;

    if (!stockPerLocation[rak]) {
      stockPerLocation[rak] = {
        lokasi: rak,
        jenis: (locMetaMap[rak] && locMetaMap[rak].jenis) ? locMetaMap[rak].jenis : 'warehouse',
        variasiSet: {},
        totalQty: 0,
        nilaiHpp: 0,
        potensiOmzet: 0
      };
    }

    if (qty > 0) {
      totalStokUnit += qty;
      totalNilaiAsetHpp += qty * vInfo.hpp;
      totalPotensiOmzet += qty * vInfo.harga_jual;

      stockPerLocation[rak].variasiSet[vid] = true;
      stockPerLocation[rak].totalQty += qty;
      stockPerLocation[rak].nilaiHpp += qty * vInfo.hpp;
      stockPerLocation[rak].potensiOmzet += qty * vInfo.harga_jual;
    }
  });

  // Agregasi Transaksi Penjualan (Keluar) & Top Produk Bulan Ini
  var qtyKeluarBulanIni = 0;
  var trxKeluarBulanIni = 0;
  var omzetBulanIni = 0;
  var profitBulanIni = 0;
  var topSalesMap = {};

  logRows.forEach(function (lg) {
    if (upperV2_(lg.Aktivitas) !== 'KELUAR' || upperV2_(lg.Status) !== 'SUCCESS') return;
    var ts = lg.Timestamp instanceof Date ? lg.Timestamp : new Date(lg.Timestamp);
    if (isNaN(ts.getTime()) || ts.getMonth() !== curMonth || ts.getFullYear() !== curYear) return;

    var qty = Number(lg.Qty) || 0;
    var hppSat = Number(lg.HPP_Satuan) || 0;
    var omzet = Number(lg.Total_Omzet) || 0;
    var profit = Number(lg.Total_Profit) || 0;
    var vid = textV2_(lg.ID_Variasi) || (textV2_(lg.Barcode) + '|' + textV2_(lg.Label_Variasi));

    qtyKeluarBulanIni += qty;
    trxKeluarBulanIni += 1;
    omzetBulanIni += omzet;
    profitBulanIni += profit;

    if (!topSalesMap[vid]) {
      topSalesMap[vid] = {
        barcode: textV2_(lg.Barcode),
        display: textV2_(lg.Produk_Tampilan) || textV2_(lg.Nama_Produk) || textV2_(lg.Barcode),
        qty: 0,
        omzet: 0,
        totalHpp: 0,
        profit: 0
      };
    }
    topSalesMap[vid].qty += qty;
    topSalesMap[vid].omzet += omzet;
    topSalesMap[vid].totalHpp += (qty * hppSat);
    topSalesMap[vid].profit += profit;
  });

  // Agregasi Gudang Penjahit (MASTER_TAILOR, STOK_TAILOR, LOG_TAILOR)
  var masterTailorMap = {};
  var tailorSkuSet = {};
  masterTailorRows.forEach(function (mt) {
    var idT = textV2_(mt.ID_Tailor);
    var sku = upperV2_(mt.Barcode_Kategori);
    var act = booleanV2_(mt.Status_Aktif, true);
    masterTailorMap[idT] = {
      id_tailor: idT,
      sku: textV2_(mt.Barcode_Kategori),
      nama: textV2_(mt.Nama_Kategori),
      warna: textV2_(mt.Warna) || '-',
      ukuran: textV2_(mt.Ukuran) || '-',
      upah: Number(mt.Upah_Per_Pcs) || 0,
      aktif: act
    };
    if (act && sku) tailorSkuSet[sku] = true;
  });

  var tailorSummaryMap = {};
  var totalStokTailor = 0;
  var totalEstUpahBerjalan = 0;

  var ensureTailorSummary = function (namaRaw) {
    var nama = textV2_(namaRaw) || 'Tanpa Nama';
    var key = upperV2_(nama);
    if (!tailorSummaryMap[key]) {
      tailorSummaryMap[key] = {
        nama_penjahit: nama,
        varianSet: {},
        stokReady: 0,
        estUpahBerjalan: 0,
        lolosBulanIni: 0,
        rejectBulanIni: 0,
        unpaidUpah: 0,
        paidBulanIni: 0
      };
    }
    return tailorSummaryMap[key];
  };

  stokTailorRows.forEach(function (st) {
    var qty = Number(st.Stok_Ready) || 0;
    if (qty <= 0) return;
    var idT = textV2_(st.ID_Tailor);
    var mtInfo = masterTailorMap[idT];
    var upahSat = mtInfo ? mtInfo.upah : 0;
    var tObj = ensureTailorSummary(st.Nama_Penjahit);

    tObj.varianSet[idT || (textV2_(st.Barcode_Kategori) + '|' + textV2_(st.Warna) + '|' + textV2_(st.Ukuran))] = true;
    tObj.stokReady += qty;
    tObj.estUpahBerjalan += (qty * upahSat);

    totalStokTailor += qty;
    totalEstUpahBerjalan += (qty * upahSat);
  });

  var qcLolosBulanIni = 0;
  var qcRejectBulanIni = 0;
  var totalUnpaidUpah = 0;
  var paidUpahBulanIni = 0;

  logTailorRows.forEach(function (lt) {
    var jenis = upperV2_(lt.Jenis_Aktivitas);
    var statusUpah = upperV2_(lt.Status_Upah);
    var qtyBagus = Number(lt.Qty_Bagus) || 0;
    var qtyReject = Number(lt.Qty_Reject) || 0;
    var totUpah = Number(lt.Total_Upah) || 0;
    var ts = lt.Waktu instanceof Date ? lt.Waktu : new Date(lt.Waktu);
    var isThisMonth = !isNaN(ts.getTime()) && ts.getMonth() === curMonth && ts.getFullYear() === curYear;

    if (jenis === 'TRANSFER_GUDANG') {
      var tObj = ensureTailorSummary(lt.Nama_Penjahit);
      if (isThisMonth) {
        tObj.lolosBulanIni += qtyBagus;
        tObj.rejectBulanIni += qtyReject;
        qcLolosBulanIni += qtyBagus;
        qcRejectBulanIni += qtyReject;
      }
      if (statusUpah === 'UNPAID' && totUpah > 0) {
        tObj.unpaidUpah += totUpah;
        totalUnpaidUpah += totUpah;
      }
      if (statusUpah === 'PAID' && isThisMonth && totUpah > 0) {
        tObj.paidBulanIni += totUpah;
        paidUpahBulanIni += totUpah;
      }
    }
  });

  var tailorList = Object.keys(tailorSummaryMap).map(function (k) { return tailorSummaryMap[k]; });
  tailorList.sort(function (a, b) {
    return (b.unpaidUpah - a.unpaidUpah) || (b.stokReady - a.stokReady);
  });

  var activeTailorCount = tailorList.filter(function (t) { return t.stokReady > 0 || t.unpaidUpah > 0; }).length;
  var unpaidTailorCount = tailorList.filter(function (t) { return t.unpaidUpah > 0; }).length;

  // Hitung Stok Menipis & Sesi Opname Aktif
  var lowRows = [];
  var lowStockOnlyCount = 0;
  Object.keys(variantMap).forEach(function (vid) {
    var v = variantMap[vid];
    if (!v.active) return;
    var m = masterMap[upperV2_(v.barcode)];
    if (m && !m.aktif) return;

    var qty = stockPerVariant[vid] || 0;
    var noPrice = (v.hpp <= 0 || v.harga_jual <= 0);
    if (qty <= 5) lowStockOnlyCount++;

    if (qty <= 5 || noPrice) {
      var marginPct = v.harga_jual > 0 ? ((v.harga_jual - v.hpp) / v.harga_jual) : 0;
      var statusNote = [];
      if (qty === 0) statusNote.push('Stok Habis (0)');
      else if (qty <= 5) statusNote.push('Stok Menipis');
      if (noPrice) statusNote.push('HPP/Harga Kosong');
      lowRows.push([
        v.barcode,
        m ? m.nama : v.barcode,
        v.label,
        qty,
        v.hpp,
        v.harga_jual,
        marginPct,
        statusNote.join(' · ')
      ]);
    }
  });
  lowRows.sort(function (a, b) { return a[3] - b[3]; });

  var activeOpnameCount = opnameRows.filter(function (o) { return upperV2_(o.Status) === 'ACTIVE'; }).length;
  var avgMargin = pricedVariantCount > 0 ? (totalMarginPctSum / pricedVariantCount) : 0;

  // ==========================================
  // 2. HEADER UTAMA DASHBOARD (A1:H2)
  // ==========================================
  dash.getRange('A1:H1').merge()
    .setValue('INA BASIC STORE · EXECUTIVE WAREHOUSE & PRODUCTION DASHBOARD')
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#0f172a').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  dash.setRowHeight(1, 42);

  dash.getRange('A2:H2').merge()
    .setValue('Periode Laporan: ' + periodeLabel + '   |   Diperbarui: ' + Utilities.formatDate(now, 'Asia/Jakarta', 'dd MMM yyyy HH:mm:ss') + ' WIB   |   Arsitektur V6.0')
    .setFontSize(9).setFontColor('#475569')
    .setBackground('#f1f5f9')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  dash.setRowHeight(2, 24);
  dash.setRowHeight(3, 12);

  // ==========================================
  // 3. KARTU KPI EKSEKUTIF 2x4 (A4:H10)
  // ==========================================
  var renderKpiCard_ = function (topRow, startCol, title, mainVal, numFormat, subText, accentBg, accentText) {
    var titleRange = dash.getRange(topRow, startCol, 1, 2);
    var valRange = dash.getRange(topRow + 1, startCol, 1, 2);
    var subRange = dash.getRange(topRow + 2, startCol, 1, 2);
    var cardBox = dash.getRange(topRow, startCol, 3, 2);

    titleRange.merge().setValue(title)
      .setFontSize(8.5).setFontWeight('bold')
      .setBackground(accentBg).setFontColor(accentText)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');

    valRange.merge().setValue(mainVal)
      .setFontSize(13).setFontWeight('bold')
      .setBackground('#ffffff').setFontColor('#0f172a')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    if (numFormat) valRange.setNumberFormat(numFormat);

    subRange.merge().setValue(subText)
      .setFontSize(8.5).setFontColor('#64748b')
      .setBackground('#ffffff')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');

    cardBox.setBorder(true, true, true, true, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  };

  // Baris KPI 1: Gudang Utama & Penjualan Bulan Berjalan (Row 4-6)
  renderKpiCard_(
    4, 1,
    'TOTAL STOK GUDANG UTAMA',
    totalStokUnit, '#,##0" Pcs"',
    activeSkuCount + ' Produk Induk · ' + activeVariantCount + ' Variasi Aktif',
    '#eff6ff', '#1d4ed8'
  );
  renderKpiCard_(
    4, 3,
    'NILAI ASET GUDANG (HPP)',
    totalNilaiAsetHpp, 'Rp #,##0',
    'Potensi Omzet Jual: Rp ' + Math.round(totalPotensiOmzet).toLocaleString('id-ID'),
    '#eef2ff', '#4338ca'
  );
  renderKpiCard_(
    4, 5,
    'OMZET KELUAR (' + periodeLabel.toUpperCase() + ')',
    omzetBulanIni, 'Rp #,##0',
    qtyKeluarBulanIni.toLocaleString('id-ID') + ' Pcs Terjual · ' + trxKeluarBulanIni + ' Transaksi',
    '#ecfdf5', '#047857'
  );
  renderKpiCard_(
    4, 7,
    'ESTIMASI PROFIT (' + periodeLabel.toUpperCase() + ')',
    profitBulanIni, 'Rp #,##0',
    'Rata-rata Margin Aktif: ' + (avgMargin * 100).toFixed(1) + '%',
    '#f0fdf4', '#15803d'
  );
  dash.setRowHeight(4, 22);
  dash.setRowHeight(5, 30);
  dash.setRowHeight(6, 20);
  dash.setRowHeight(7, 10);

  // Baris KPI 2: Produksi Penjahit, Upah & Kontrol Audit (Row 8-10)
  renderKpiCard_(
    8, 1,
    'STOK SIAP DI PENJAHIT (TAILOR)',
    totalStokTailor, '#,##0" Pcs"',
    activeTailorCount + ' Penjahit Aktif · Est. Upah: Rp ' + Math.round(totalEstUpahBerjalan).toLocaleString('id-ID'),
    '#f0fdfa', '#0f766e'
  );
  renderKpiCard_(
    8, 3,
    'TAGIHAN UPAH BELUM BAYAR (UNPAID)',
    totalUnpaidUpah, 'Rp #,##0',
    unpaidTailorCount + ' Penjahit Menunggu Pencetakan Slip',
    '#f5f3ff', '#6d28d9'
  );
  renderKpiCard_(
    8, 5,
    'SETORAN LOLOS QC (' + periodeLabel.toUpperCase() + ')',
    qcLolosBulanIni, '#,##0" Pcs"',
    'Reject: ' + qcRejectBulanIni + ' Pcs · Upah Lunas: Rp ' + Math.round(paidUpahBulanIni).toLocaleString('id-ID'),
    '#f8fafc', '#334155'
  );
  renderKpiCard_(
    8, 7,
    'KONTROL STOK & AUDIT DATA',
    lowStockOnlyCount + ' Item Menipis', null,
    missingPriceCount + ' Harga Kosong · ' + activeOpnameCount + ' Sesi Opname Aktif',
    '#fffbeb', '#b45309'
  );
  dash.setRowHeight(8, 22);
  dash.setRowHeight(9, 30);
  dash.setRowHeight(10, 20);
  dash.setRowHeight(11, 14);

  var cursor = 12;

  // Helper Section & Table Styling
  var renderSectionTitle_ = function (titleText, bgHex, fgHex) {
    dash.getRange(cursor, 1, 1, 8).merge()
      .setValue(titleText)
      .setFontSize(10).setFontWeight('bold')
      .setBackground(bgHex || '#1e293b')
      .setFontColor(fgHex || '#ffffff')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    dash.setRowHeight(cursor, 30);
    cursor++;
  };

  var renderTableHeaders_ = function (headers) {
    dash.getRange(cursor, 1, 1, 8).setValues([headers])
      .setFontSize(9).setFontWeight('bold')
      .setBackground('#e2e8f0').setFontColor('#1e293b')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true);
    dash.setRowHeight(cursor, 32);
    var headerRow = cursor;
    cursor++;
    return headerRow;
  };

  var styleDataBlock_ = function (headerRow, dataStartRow, numDataRows, rightAlignCols, hasTotalRow) {
    if (numDataRows > 0) {
      var bgMatrix = [];
      for (var i = 0; i < numDataRows; i++) {
        var rowBg = (i % 2 === 0) ? '#ffffff' : '#f8fafc';
        bgMatrix.push([rowBg, rowBg, rowBg, rowBg, rowBg, rowBg, rowBg, rowBg]);
      }
      var dataRange = dash.getRange(dataStartRow, 1, numDataRows, 8);
      dataRange.setBackgrounds(bgMatrix)
        .setFontSize(9.5).setFontColor('#1e293b')
        .setVerticalAlignment('middle').setWrap(true);
      dash.setRowHeights(dataStartRow, numDataRows, 26);

      (rightAlignCols || []).forEach(function (colIdx) {
        dash.getRange(dataStartRow, colIdx, numDataRows, 1).setHorizontalAlignment('right');
      });
    }

    var totalRowsBlock = numDataRows + (hasTotalRow ? 1 : 0);
    dash.getRange(headerRow, 1, totalRowsBlock + 1, 8)
      .setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

    if (hasTotalRow) {
      var totRow = dataStartRow + numDataRows;
      dash.getRange(totRow, 1, 1, 8)
        .setBackground('#f1f5f9').setFontWeight('bold').setFontSize(9.5).setFontColor('#0f172a')
        .setVerticalAlignment('middle');
      (rightAlignCols || []).forEach(function (colIdx) {
        dash.getRange(totRow, colIdx, 1, 1).setHorizontalAlignment('right');
      });
      dash.setRowHeight(totRow, 28);
    }
  };

  // ==========================================
  // 4. TABEL 1: MONITORING GUDANG PENJAHIT & TAGIHAN UPAH
  // ==========================================
  renderSectionTitle_('1. MONITORING STOK PENJAHIT & REKAP UPAH (MASTER TAILOR V6.0)', '#0f172a', '#ffffff');
  var t1Header = renderTableHeaders_([
    'Nama Penjahit / Tim',
    'Varian Dikerjakan',
    'Stok Ready (Pcs)',
    'Est. Upah Berjalan',
    'Lolos QC (Bln Ini)',
    'Reject (Bln Ini)',
    'Tagihan UNPAID',
    'Status Upah'
  ]);
  var t1Start = cursor;

  if (tailorList.length > 0) {
    var t1Rows = tailorList.map(function (t) {
      var varCount = Object.keys(t.varianSet).length;
      var statusLabel = t.unpaidUpah > 0
        ? 'Menunggu Slip (UNPAID)'
        : (t.stokReady > 0 ? 'Proses Jahit Aktif' : 'Lunas (PAID)');
      return [
        t.nama_penjahit,
        varCount > 0 ? (varCount + ' Variasi Aktif') : '-',
        t.stokReady,
        t.estUpahBerjalan,
        t.lolosBulanIni,
        t.rejectBulanIni,
        t.unpaidUpah,
        statusLabel
      ];
    });

    dash.getRange(cursor, 1, t1Rows.length, 8).setValues(t1Rows);
    dash.getRange(cursor, 2, t1Rows.length, 1).setHorizontalAlignment('center');
    dash.getRange(cursor, 3, t1Rows.length, 1).setNumberFormat('#,##0');
    dash.getRange(cursor, 4, t1Rows.length, 1).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 5, t1Rows.length, 2).setNumberFormat('#,##0');
    dash.getRange(cursor, 7, t1Rows.length, 1).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 8, t1Rows.length, 1).setHorizontalAlignment('center');
    cursor += t1Rows.length;

    // Baris Total Tabel 1
    dash.getRange(cursor, 1, 1, 8).setValues([[
      'TOTAL KESELURUHAN PENJAHIT',
      tailorList.length + ' Penjahit',
      totalStokTailor,
      totalEstUpahBerjalan,
      qcLolosBulanIni,
      qcRejectBulanIni,
      totalUnpaidUpah,
      totalUnpaidUpah > 0 ? 'Ada Tagihan UNPAID' : 'Semua Lunas'
    ]]);
    dash.getRange(cursor, 2, 1, 1).setHorizontalAlignment('center');
    dash.getRange(cursor, 3, 1, 1).setNumberFormat('#,##0');
    dash.getRange(cursor, 4, 1, 1).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 5, 1, 2).setNumberFormat('#,##0');
    dash.getRange(cursor, 7, 1, 1).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 8, 1, 1).setHorizontalAlignment('center');

    styleDataBlock_(t1Header, t1Start, t1Rows.length, [3, 4, 5, 6, 7], true);

    // Highlight kolom Tagihan UNPAID yang > 0
    for (var i1 = 0; i1 < t1Rows.length; i1++) {
      if (t1Rows[i1][6] > 0) {
        dash.getRange(t1Start + i1, 7, 1, 2).setBackground('#f5f3ff').setFontColor('#5b21b6').setFontWeight('bold');
      }
    }
    cursor += 2;
  } else {
    dash.getRange(cursor, 1, 1, 8).merge()
      .setValue('Belum ada data stok atau transaksi penjahit yang tercatat.')
      .setFontStyle('italic').setFontColor('#64748b').setBackground('#ffffff');
    dash.setRowHeight(cursor, 28);
    dash.getRange(t1Header, 1, 2, 8).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    cursor += 2;
  }

  // ==========================================
  // 5. TABEL 2: REKAP ASET STOK PER LOKASI RAK
  // ==========================================
  renderSectionTitle_('2. DISTRIBUSI STOK & NILAI ASET PER LOKASI RAK', '#1e293b', '#ffffff');
  var t2Header = renderTableHeaders_([
    'Lokasi Rak',
    'Kategori Rak',
    'Variasi Tersimpan',
    'Total Stok (Pcs)',
    'Nilai Aset (HPP)',
    'Potensi Omzet',
    'Potensi Profit',
    'Porsi Stok'
  ]);
  var t2Start = cursor;

  var locList = Object.keys(stockPerLocation).map(function (k) { return stockPerLocation[k]; });
  locList.sort(function (a, b) { return b.totalQty - a.totalQty; });

  if (locList.length > 0) {
    var t2Rows = locList.map(function (loc) {
      var potProfit = loc.potensiOmzet - loc.nilaiHpp;
      var share = totalStokUnit > 0 ? (loc.totalQty / totalStokUnit) : 0;
      return [
        loc.lokasi,
        upperV2_(loc.jenis),
        Object.keys(loc.variasiSet).length + ' Variasi',
        loc.totalQty,
        loc.nilaiHpp,
        loc.potensiOmzet,
        potProfit,
        share
      ];
    });

    dash.getRange(cursor, 1, t2Rows.length, 8).setValues(t2Rows);
    dash.getRange(cursor, 2, t2Rows.length, 2).setHorizontalAlignment('center');
    dash.getRange(cursor, 4, t2Rows.length, 1).setNumberFormat('#,##0');
    dash.getRange(cursor, 5, t2Rows.length, 3).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 8, t2Rows.length, 1).setNumberFormat('0.0%');
    cursor += t2Rows.length;

    dash.getRange(cursor, 1, 1, 8).setValues([[
      'TOTAL GUDANG UTAMA',
      locList.length + ' Lokasi',
      activeVariantCount + ' Variasi Aktif',
      totalStokUnit,
      totalNilaiAsetHpp,
      totalPotensiOmzet,
      totalPotensiOmzet - totalNilaiAsetHpp,
      totalStokUnit > 0 ? 1 : 0
    ]]);
    dash.getRange(cursor, 2, 1, 2).setHorizontalAlignment('center');
    dash.getRange(cursor, 4, 1, 1).setNumberFormat('#,##0');
    dash.getRange(cursor, 5, 1, 3).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 8, 1, 1).setNumberFormat('0.0%');

    styleDataBlock_(t2Header, t2Start, t2Rows.length, [4, 5, 6, 7, 8], true);
    cursor += 2;
  } else {
    dash.getRange(cursor, 1, 1, 8).merge()
      .setValue('Belum ada stok aktif di lokasi rak manapun.')
      .setFontStyle('italic').setFontColor('#64748b').setBackground('#ffffff');
    dash.setRowHeight(cursor, 28);
    dash.getRange(t2Header, 1, 2, 8).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    cursor += 2;
  }

  // ==========================================
  // 6. TABEL 3: TOP 10 PRODUK TERLARIS BULAN INI
  // ==========================================
  renderSectionTitle_('3. TOP 10 PRODUK & VARIASI TERLARIS (' + periodeLabel.toUpperCase() + ')', '#065f46', '#ffffff');
  var t3Header = renderTableHeaders_([
    'Peringkat',
    'Barcode',
    'Nama Produk & Variasi',
    'Qty Terjual (Pcs)',
    'Total Omzet',
    'Total Modal (HPP)',
    'Estimasi Profit',
    'Margin Aktual'
  ]);
  var t3Start = cursor;

  var topSalesList = Object.keys(topSalesMap).map(function (k) { return topSalesMap[k]; });
  topSalesList.sort(function (a, b) { return (b.qty - a.qty) || (b.omzet - a.omzet); });
  var top10 = topSalesList.slice(0, 10);

  if (top10.length > 0) {
    var sumTopQty = 0;
    var sumTopOmzet = 0;
    var sumTopHpp = 0;
    var sumTopProfit = 0;

    var t3Rows = top10.map(function (item, idx) {
      sumTopQty += item.qty;
      sumTopOmzet += item.omzet;
      sumTopHpp += item.totalHpp;
      sumTopProfit += item.profit;
      var mPct = item.omzet > 0 ? (item.profit / item.omzet) : 0;
      return [
        '#' + (idx + 1),
        item.barcode,
        item.display,
        item.qty,
        item.omzet,
        item.totalHpp,
        item.profit,
        mPct
      ];
    });

    dash.getRange(cursor, 1, t3Rows.length, 8).setValues(t3Rows);
    dash.getRange(cursor, 1, t3Rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
    dash.getRange(cursor, 2, t3Rows.length, 1).setNumberFormat('@');
    dash.getRange(cursor, 4, t3Rows.length, 1).setNumberFormat('#,##0');
    dash.getRange(cursor, 5, t3Rows.length, 3).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 8, t3Rows.length, 1).setNumberFormat('0.0%');
    cursor += t3Rows.length;

    dash.getRange(cursor, 1, 1, 8).setValues([[
      'TOTAL',
      '-',
      'Akumulasi Top ' + top10.length + ' Produk Terlaris',
      sumTopQty,
      sumTopOmzet,
      sumTopHpp,
      sumTopProfit,
      sumTopOmzet > 0 ? (sumTopProfit / sumTopOmzet) : 0
    ]]);
    dash.getRange(cursor, 1, 1, 1).setHorizontalAlignment('center');
    dash.getRange(cursor, 4, 1, 1).setNumberFormat('#,##0');
    dash.getRange(cursor, 5, 1, 3).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 8, 1, 1).setNumberFormat('0.0%');

    styleDataBlock_(t3Header, t3Start, t3Rows.length, [4, 5, 6, 7, 8], true);
    cursor += 2;
  } else {
    dash.getRange(cursor, 1, 1, 8).merge()
      .setValue('Belum ada transaksi barang keluar (penjualan) pada bulan ' + periodeLabel + '.')
      .setFontStyle('italic').setFontColor('#64748b').setBackground('#ffffff');
    dash.setRowHeight(cursor, 28);
    dash.getRange(t3Header, 1, 2, 8).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    cursor += 2;
  }

  // ==========================================
  // 7. TABEL 4: PERINGATAN STOK MENIPIS & KELENGKAPAN HARGA
  // ==========================================
  renderSectionTitle_('4. PERINGATAN STOK MENIPIS (<= 5 PCS) & KELENGKAPAN HARGA JUAL', '#92400e', '#ffffff');
  var t4Header = renderTableHeaders_([
    'Barcode',
    'Nama Produk',
    'Variasi',
    'Sisa Stok (Pcs)',
    'HPP Satuan',
    'Harga Jual',
    'Margin (%)',
    'Status & Tindakan'
  ]);
  var t4Start = cursor;
  var displayLow = lowRows.slice(0, 30);

  if (displayLow.length > 0) {
    dash.getRange(cursor, 1, displayLow.length, 8).setValues(displayLow);
    dash.getRange(cursor, 1, displayLow.length, 1).setNumberFormat('@');
    dash.getRange(cursor, 4, displayLow.length, 1).setNumberFormat('#,##0');
    dash.getRange(cursor, 5, displayLow.length, 2).setNumberFormat('Rp #,##0');
    dash.getRange(cursor, 7, displayLow.length, 1).setNumberFormat('0.0%');
    dash.getRange(cursor, 8, displayLow.length, 1).setHorizontalAlignment('center');

    styleDataBlock_(t4Header, t4Start, displayLow.length, [4, 5, 6, 7], false);

    // Pewarnaan status berdasarkan tingkat urgensi
    for (var i4 = 0; i4 < displayLow.length; i4++) {
      var sisaQty = displayLow[i4][3];
      var statusCell = dash.getRange(t4Start + i4, 8);
      if (sisaQty === 0) {
        statusCell.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
      } else if (sisaQty <= 5) {
        statusCell.setBackground('#fef3c7').setFontColor('#92400e').setFontWeight('bold');
      } else {
        statusCell.setBackground('#e0e7ff').setFontColor('#3730a3').setFontWeight('bold');
      }
    }
    cursor += displayLow.length + 1;
  } else {
    dash.getRange(cursor, 1, 1, 8).merge()
      .setValue('Semua variasi produk aktif memiliki stok > 5 pcs dan data HPP/Harga Jual lengkap.')
      .setFontColor('#047857').setFontWeight('bold').setBackground('#ecfdf5');
    dash.setRowHeight(cursor, 28);
    dash.getRange(t4Header, 1, 2, 8).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    cursor += 2;
  }

  // ==========================================
  // 8. TABEL 5: 10 AKTIVITAS GUDANG & PENJAHIT TERBARU
  // ==========================================
  renderSectionTitle_('5. LOG AKTIVITAS TERBARU (GUDANG UTAMA & GUDANG PENJAHIT)', '#334155', '#ffffff');
  var t5Header = renderTableHeaders_([
    'Waktu (WIB)',
    'Modul',
    'Aktivitas',
    'Detail Produk / Item',
    'Asal / Penjahit',
    'Tujuan / Status',
    'Qty (Pcs)',
    'Pelapor / Admin'
  ]);
  var t5Start = cursor;

  var recentEvents = [];
  logRows.slice(-15).forEach(function (lg) {
    var ts = lg.Timestamp instanceof Date ? lg.Timestamp : new Date(lg.Timestamp);
    if (isNaN(ts.getTime())) return;
    recentEvents.push({
      ts: ts,
      modul: 'GUDANG',
      aktivitas: textV2_(lg.Aktivitas),
      detail: textV2_(lg.Produk_Tampilan) || textV2_(lg.Barcode),
      asal: textV2_(lg.Lokasi_Asal) || '-',
      tujuan: textV2_(lg.Lokasi_Tujuan) || textV2_(lg.Status),
      qty: Number(lg.Qty) || 0,
      pelapor: textV2_(lg.Nama_Pelapor) || '-'
    });
  });

  logTailorRows.slice(-15).forEach(function (lt) {
    var ts = lt.Waktu instanceof Date ? lt.Waktu : new Date(lt.Waktu);
    if (isNaN(ts.getTime())) return;
    recentEvents.push({
      ts: ts,
      modul: 'PENJAHIT',
      aktivitas: textV2_(lt.Jenis_Aktivitas),
      detail: textV2_(lt.Detail_Item),
      asal: textV2_(lt.Nama_Penjahit) || '-',
      tujuan: textV2_(lt.Status_Upah) || '-',
      qty: Number(lt.Qty_Bagus) || 0,
      pelapor: textV2_(lt.ID_Slip) || textV2_(lt.Nama_Penjahit) || '-'
    });
  });

  recentEvents.sort(function (a, b) { return b.ts.getTime() - a.ts.getTime(); });
  var topRecent = recentEvents.slice(0, 10);

  if (topRecent.length > 0) {
    var t5Rows = topRecent.map(function (ev) {
      return [
        Utilities.formatDate(ev.ts, 'Asia/Jakarta', 'dd/MM HH:mm'),
        ev.modul,
        ev.aktivitas,
        ev.detail,
        ev.asal,
        ev.tujuan,
        ev.qty,
        ev.pelapor
      ];
    });
    dash.getRange(cursor, 1, t5Rows.length, 8).setValues(t5Rows);
    dash.getRange(cursor, 1, t5Rows.length, 3).setHorizontalAlignment('center');
    dash.getRange(cursor, 7, t5Rows.length, 1).setNumberFormat('#,##0');
    styleDataBlock_(t5Header, t5Start, t5Rows.length, [7], false);
    cursor += t5Rows.length;
  } else {
    dash.getRange(cursor, 1, 1, 8).merge()
      .setValue('Belum ada riwayat aktivitas transaksi.')
      .setFontStyle('italic').setFontColor('#64748b').setBackground('#ffffff');
    dash.setRowHeight(cursor, 28);
    dash.getRange(t5Header, 1, 2, 8).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    cursor += 1;
  }

  // ==========================================
  // 9. PENGATURAN LEBAR KOLOM & FONT GLOBAL
  // ==========================================
  var finalRow = Math.max(cursor, 12);
  dash.setFrozenRows(2);
  dash.setFrozenColumns(0);

  dash.setColumnWidth(1, 145);
  dash.setColumnWidth(2, 165);
  dash.setColumnWidth(3, 215);
  dash.setColumnWidth(4, 120);
  dash.setColumnWidth(5, 145);
  dash.setColumnWidth(6, 145);
  dash.setColumnWidth(7, 145);
  dash.setColumnWidth(8, 175);

  dash.getRange(1, 1, finalRow, 8).setFontFamily('Arial');

  SpreadsheetApp.flush();
  try {
    ss.toast('Dashboard Eksekutif V6.0 berhasil diperbarui.', 'Ina Basic Store', 4);
  } catch (e) {}
}