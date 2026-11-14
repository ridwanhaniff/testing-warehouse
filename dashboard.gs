/**
 * INA BASIC STORE - Executive Dashboard & Periodic Report V4.2
 * Mencakup: Inventaris Stok Variasi, Penjualan/Keluar, Bahan Baku, WO Jahit,
 * Quality Control (QC), Rekap Upah & Slip Gaji, HPP Produksi, serta
 * Generator Tarik Laporan (Harian / Mingguan / Bulanan / Custom).
 */
var IBS_DASH = {
  DASH: 'DASHBOARD',
  DATA: '_DASH_DATA',
  REPORT: 'LAPORAN_REKAP',
  TZ: 'Asia/Jakarta',
  COLOR: {
    page: '#f1f5f9',
    dark: '#0f172a',
    slate: '#1e293b',
    text: '#1e293b',
    muted: '#64748b',
    line: '#cbd5e1',
    indigo: '#4f46e5',
    green: '#059669',
    red: '#e11d48',
    amber: '#d97706',
    sky: '#0284c7',
    violet: '#7c3aed'
  }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ina Warehouse')
    .addItem('Siapkan / Perbarui Dashboard V4.2', 'setupDashboard')
    .addItem('Tarik Laporan (Harian / Mingguan / Bulanan)', 'openReportDialog')
    .addSeparator()
    .addItem('Siapkan Sheet Sistem V4.2', 'setup')
    .addItem('Jalankan Migrasi Stok Lama', 'showMigrationReminderV2_')
    .addToUi();
}

function showMigrationReminderV2_() {
  SpreadsheetApp.getUi().alert(
    'Migrasi stok lama',
    'Isi sheet MIGRASI_ALOKASI dengan pembagian stok lama ke ID_Variasi. Pastikan total per Barcode_Induk + Lokasi_Rak sama persis dengan STOK_LOKASI. Setelah valid, jalankan fungsi migrateLegacyStockV2 dari editor Apps Script. Sheet lama tidak dihapus.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setupDashboard() {
  if (typeof setup === 'function') setup();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(IBS_DASH.TZ);
  if (!ss.getSheetByName(IBS_DASH.DASH)) ss.insertSheet(IBS_DASH.DASH);
  if (!ss.getSheetByName(IBS_DASH.DATA)) ss.insertSheet(IBS_DASH.DATA);

  buildDashDataV4_(ss);
  buildDashboardV4_(ss);

  var dash = ss.getSheetByName(IBS_DASH.DASH);
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
  SpreadsheetApp.flush();
}

// ============================================================================
// 1. PEMBANGUN SHEET HELPER TERSEMBUNYI (_DASH_DATA)
// ============================================================================
function buildDashDataV4_(ss) {
  var sh = ss.getSheetByName(IBS_DASH.DATA) || ss.insertSheet(IBS_DASH.DATA);
  sh.clear();

  // --- BLOK A: Grafik Tren 14 Hari (Cols A:C) ---
  sh.getRange('A1:C1').setValues([['Tanggal', 'Masuk', 'Keluar']]);
  var formulas = [];
  for (var r = 2; r <= 15; r++) {
    var dateRef = '$A' + r;
    formulas.push([
      '=TODAY()-' + (15 - r),
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&' + dateRef + ',OPERATION_LOG!$A$2:$A,"<"&(' + dateRef + '+1))',
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&' + dateRef + ',OPERATION_LOG!$A$2:$A,"<"&(' + dateRef + '+1))'
    ]);
  }
  sh.getRange(2, 1, formulas.length, 3).setFormulas(formulas);
  sh.getRange('A2:A15').setNumberFormat('dd MMM');
  sh.getRange('B2:C15').setNumberFormat('0');

  // --- BLOK B: Parameter Filter Dinamis (Cols E:H) ---
  sh.getRange('E1:H1').setValues([['TGL_AWAL', 'TGL_AKHIR', 'LOKASI_LIST', 'TIM_LIST']]);
  sh.getRange('E2').setFormula(
    '=IFS(' +
    'DASHBOARD!$D$4="Hari Ini",TODAY(),' +
    'DASHBOARD!$D$4="Kemarin",TODAY()-1,' +
    'DASHBOARD!$D$4="7 Hari Terakhir (Mingguan)",TODAY()-6,' +
    'DASHBOARD!$D$4="30 Hari Terakhir",TODAY()-29,' +
    'DASHBOARD!$D$4="Bulan Ini",EOMONTH(TODAY(),-1)+1,' +
    'DASHBOARD!$D$4="Bulan Lalu",EOMONTH(TODAY(),-2)+1,' +
    'DASHBOARD!$D$4="Custom Tanggal",IF(ISDATE(DASHBOARD!$G$4),DASHBOARD!$G$4,TODAY()-29),' +
    'DASHBOARD!$D$4="Semua Waktu",DATE(2000,1,1),' +
    'TRUE,TODAY()-29)'
  );
  sh.getRange('F2').setFormula(
    '=IFS(' +
    'DASHBOARD!$D$4="Kemarin",TODAY(),' +
    'DASHBOARD!$D$4="Bulan Lalu",EOMONTH(TODAY(),-1)+1,' +
    'DASHBOARD!$D$4="Custom Tanggal",IF(ISDATE(DASHBOARD!$J$4),DASHBOARD!$J$4+1,TODAY()+1),' +
    'TRUE,TODAY()+1)'
  );
  sh.getRange('E2:F2').setNumberFormat('yyyy-mm-dd');

  sh.getRange('G2').setValue('Semua Lokasi');
  sh.getRange('G3').setFormula('=IFERROR(SORT(UNIQUE(FILTER(MASTER_LOKASI!B2:B,MASTER_LOKASI!D2:D=TRUE))),IFERROR(SORT(UNIQUE(FILTER(STOK_VARIASI!C2:C,STOK_VARIASI!C2:C<>""))),""))');

  sh.getRange('H2').setValue('Semua Tim');
  sh.getRange('H3').setFormula('=IFERROR(SORT(UNIQUE(FILTER(WO_JAHIT!D2:D,WO_JAHIT!D2:D<>""))),"")');

  // --- BLOK C: Mirror OPERATION_LOG (Cols J:Q) ---
  sh.getRange('J1:Q1').setValues([['Timestamp', 'Aktivitas', 'Barcode', 'Produk_Tampilan', 'Qty', 'Lokasi_Asal', 'Lokasi_Tujuan', 'Pelapor']]);
  sh.getRange('J2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!A2:A))');
  sh.getRange('K2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!C2:C))');
  sh.getRange('L2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!D2:D))');
  sh.getRange('M2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",IF(OPERATION_LOG!M2:M<>"",OPERATION_LOG!M2:M,IFERROR(VLOOKUP(OPERATION_LOG!D2:D,MASTER_PRODUK!A:B,2,FALSE),OPERATION_LOG!D2:D))))');
  sh.getRange('N2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",N(OPERATION_LOG!G2:G)))');
  sh.getRange('O2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!E2:E))');
  sh.getRange('P2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!F2:F))');
  sh.getRange('Q2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!N2:N))');

  // --- BLOK D: Mirror PENGIRIMAN_BAHAN + MASTER_BAHAN (Cols S:AB) ---
  sh.getRange('S1:AB1').setValues([['Created_At', 'Delivery_ID', 'Nomor_WO', 'Tim_Penjahit', 'Kode_Bahan', 'Nama_Bahan', 'Qty_Kirim', 'Qty_Terima', 'Nilai_Modal_Rp', 'Status']]);
  sh.getRange('S2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",PENGIRIMAN_BAHAN!K2:K))');
  sh.getRange('T2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",PENGIRIMAN_BAHAN!A2:A))');
  sh.getRange('U2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",IFERROR(VLOOKUP(PENGIRIMAN_BAHAN!B2:B,WO_JAHIT!A:B,2,FALSE),PENGIRIMAN_BAHAN!B2:B)))');
  sh.getRange('V2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",IFERROR(VLOOKUP(PENGIRIMAN_BAHAN!B2:B,WO_JAHIT!A:D,4,FALSE),"-")))');
  sh.getRange('W2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",PENGIRIMAN_BAHAN!C2:C))');
  sh.getRange('X2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",IFERROR(VLOOKUP(PENGIRIMAN_BAHAN!C2:C,MASTER_BAHAN!A:B,2,FALSE),PENGIRIMAN_BAHAN!C2:C)))');
  sh.getRange('Y2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",N(PENGIRIMAN_BAHAN!E2:E)))');
  sh.getRange('Z2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",N(PENGIRIMAN_BAHAN!H2:H)))');
  sh.getRange('AA2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",IF(N(PENGIRIMAN_BAHAN!H2:H)>0,N(PENGIRIMAN_BAHAN!H2:H),N(PENGIRIMAN_BAHAN!E2:E))*IFERROR(VLOOKUP(PENGIRIMAN_BAHAN!C2:C,MASTER_BAHAN!A:D,4,FALSE),0)))');
  sh.getRange('AB2').setFormula('=ARRAYFORMULA(IF(PENGIRIMAN_BAHAN!A2:A="","",PENGIRIMAN_BAHAN!I2:I))');

  // --- BLOK E: Mirror PENERIMAAN_QC + WO_JAHIT (Cols AD:AK) ---
  sh.getRange('AD1:AK1').setValues([['Created_At', 'QC_ID', 'Nomor_WO', 'Tim_Penjahit', 'Qty_Lolos', 'Qty_Gagal', 'Nilai_Upah', 'Status_Bayar']]);
  sh.getRange('AD2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",PENERIMAAN_QC!L2:L))');
  sh.getRange('AE2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",PENERIMAAN_QC!A2:A))');
  sh.getRange('AF2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",IFERROR(VLOOKUP(PENERIMAAN_QC!C2:C,WO_JAHIT!A:B,2,FALSE),PENERIMAAN_QC!C2:C)))');
  sh.getRange('AG2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",PENERIMAAN_QC!D2:D))');
  sh.getRange('AH2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",N(PENERIMAAN_QC!E2:E)))');
  sh.getRange('AI2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",N(PENERIMAAN_QC!F2:F)))');
  sh.getRange('AJ2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",N(PENERIMAAN_QC!H2:H)))');
  sh.getRange('AK2').setFormula('=ARRAYFORMULA(IF(PENERIMAAN_QC!A2:A="","",PENERIMAAN_QC!I2:I))');

  // --- BLOK F: Mirror HPP_PRODUKSI + WO_JAHIT (Cols AM:AT) ---
  sh.getRange('AM1:AT1').setValues([['Created_At', 'HPP_ID', 'Nomor_WO', 'Tim_Penjahit', 'Cost_Bahan', 'Cost_Jahit', 'Qty_Lolos', 'HPP_Per_Pcs']]);
  sh.getRange('AM2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",HPP_PRODUKSI!G2:G))');
  sh.getRange('AN2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",HPP_PRODUKSI!A2:A))');
  sh.getRange('AO2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",IFERROR(VLOOKUP(HPP_PRODUKSI!B2:B,WO_JAHIT!A:B,2,FALSE),HPP_PRODUKSI!B2:B)))');
  sh.getRange('AP2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",IFERROR(VLOOKUP(HPP_PRODUKSI!B2:B,WO_JAHIT!A:D,4,FALSE),"-")))');
  sh.getRange('AQ2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",N(HPP_PRODUKSI!C2:C)))');
  sh.getRange('AR2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",N(HPP_PRODUKSI!D2:D)))');
  sh.getRange('AS2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",N(HPP_PRODUKSI!E2:E)))');
  sh.getRange('AT2').setFormula('=ARRAYFORMULA(IF(HPP_PRODUKSI!A2:A="","",N(HPP_PRODUKSI!F2:F)))');

  // --- BLOK G: Mirror WO_JAHIT Agregat Target vs QC (Cols AV:BC) ---
  sh.getRange('AV1:BC1').setValues([['Created_At', 'WO_ID', 'Nomor_WO', 'Tim_Label', 'Status', 'Target_Qty', 'Lolos_QC', 'Gagal_QC']]);
  sh.getRange('AV2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",WO_JAHIT!J2:J))');
  sh.getRange('AW2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",WO_JAHIT!A2:A))');
  sh.getRange('AX2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",WO_JAHIT!B2:B))');
  sh.getRange('AY2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",WO_JAHIT!D2:D))');
  sh.getRange('AZ2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",WO_JAHIT!C2:C))');
  sh.getRange('BA2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",SUMIF(WO_JAHIT_ITEM!B2:B,WO_JAHIT!A2:A,WO_JAHIT_ITEM!D2:D)))');
  sh.getRange('BB2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",SUMIF(PENERIMAAN_QC!C2:C,WO_JAHIT!A2:A,PENERIMAAN_QC!E2:E)))');
  sh.getRange('BC2').setFormula('=ARRAYFORMULA(IF(WO_JAHIT!A2:A="","",SUMIF(PENERIMAAN_QC!C2:C,WO_JAHIT!A2:A,PENERIMAAN_QC!F2:F)))');

  sh.hideSheet();
}

function sparkFormulaV4_(col, row, first, last, color) {
  return '=IF(' + col + row + '="","",IFERROR(SPARKLINE(' + col + row + ',{"charttype","bar";"max",MAX($' + col + '$' + first + ':$' + col + '$' + last + ');"color1","' + color + '"}),""))';
}

// ============================================================================
// 2. PEMBANGUN HALAMAN UTAMA DASHBOARD V4.2
// ============================================================================
function buildDashboardV4_(ss) {
  var sh = ss.getSheetByName(IBS_DASH.DASH) || ss.insertSheet(IBS_DASH.DASH);
  sh.getCharts().forEach(function (chart) { sh.removeChart(chart); });
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();

  var C = IBS_DASH.COLOR;
  var R = function (a1) { return sh.getRange(a1); };
  var dataSh = ss.getSheetByName(IBS_DASH.DATA);

  sh.setHiddenGridlines(true);

  // Atur lebar kolom agar setiap kolom tabel memiliki ruang lega tanpa merge di area data
  sh.setColumnWidth(1, 16);   // A: Margin kiri
  sh.setColumnWidth(2, 165);  // B
  sh.setColumnWidth(3, 135);  // C
  sh.setColumnWidth(4, 95);   // D
  sh.setColumnWidth(5, 95);   // E
  sh.setColumnWidth(6, 105);  // F
  sh.setColumnWidth(7, 115);  // G
  sh.setColumnWidth(8, 22);   // H: Gutter tengah
  sh.setColumnWidth(9, 155);  // I
  sh.setColumnWidth(10, 135); // J
  sh.setColumnWidth(11, 110); // K
  sh.setColumnWidth(12, 95);  // L
  sh.setColumnWidth(13, 110); // M
  sh.setColumnWidth(14, 105); // N
  sh.setColumnWidth(15, 16);  // O: Margin kanan

  sh.setRowHeights(1, 95, 22);
  [1, 6, 9, 12, 25, 42, 55, 68, 81, 94].forEach(function (r) { sh.setRowHeight(r, 12); });
  sh.setRowHeight(2, 46);
  sh.setRowHeight(3, 24);
  sh.setRowHeight(4, 28);
  sh.setRowHeight(5, 28);
  sh.setRowHeight(8, 48);
  sh.setRowHeight(11, 48);

  R('A1:O95').setBackground(C.page).setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');

  // Banner Header
  R('B2:N2').merge().setValue('  INA BASIC STORE  ·  EXECUTIVE WAREHOUSE, PRODUKSI & FINANCE DASHBOARD')
    .setBackground(C.dark).setFontColor('#ffffff').setFontSize(16).setFontWeight('bold');
  R('B3:N3').merge().setFormula('="  Diperbarui: "&TEXT(NOW(),"dd MMM yyyy HH:mm")&" WIB  |  Gunakan menu Ina Warehouse → Tarik Laporan untuk ekspor rekap Harian / Mingguan / Bulanan"')
    .setBackground(C.slate).setFontColor('#cbd5e1').setFontSize(9);

  // --- BARIS FILTER 1: PERIODE WAKTU ---
  R('B4').setValue('Periode Waktu').setFontWeight('bold').setFontColor(C.text);
  R('C4:D4').merge().setValue('30 Hari Terakhir').setBackground('#fef9c3').setFontWeight('bold').setHorizontalAlignment('center');
  R('C4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([
    'Hari Ini', 'Kemarin', '7 Hari Terakhir (Mingguan)', '30 Hari Terakhir',
    'Bulan Ini', 'Bulan Lalu', 'Semua Waktu', 'Custom Tanggal'
  ], true).build());

  R('E4').setValue('Tgl Mulai').setHorizontalAlignment('right').setFontColor(C.muted);
  R('F4:G4').merge().setFormula('=TODAY()-29').setBackground('#fef9c3').setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');

  R('I4').setValue('Tgl Akhir').setHorizontalAlignment('right').setFontColor(C.muted);
  R('J4').setFormula('=TODAY()').setBackground('#fef9c3').setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');

  R('K4:N4').merge().setFormula('="Rentang Aktif: "&TEXT(_DASH_DATA!E2,"dd MMM yyyy")&" s/d "&TEXT(_DASH_DATA!F2-1,"dd MMM yyyy")')
    .setFontColor(C.indigo).setFontWeight('bold').setHorizontalAlignment('right').setFontSize(9);

  // --- BARIS FILTER 2: LOKASI, TIM PENJAHIT & BATAS KRITIS ---
  R('B5').setValue('Filter Lokasi').setFontWeight('bold').setFontColor(C.text);
  R('C5:D5').merge().setValue('Semua Lokasi').setBackground('#fef9c3').setFontWeight('bold').setHorizontalAlignment('center');
  R('C5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(dataSh.getRange('G2:G100'), true).build());

  R('E5').setValue('Tim Penjahit').setHorizontalAlignment('right').setFontColor(C.muted);
  R('F5:G5').merge().setValue('Semua Tim').setBackground('#fef9c3').setFontWeight('bold').setHorizontalAlignment('center');
  R('F5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(dataSh.getRange('H2:H100'), true).build());

  R('I5:J5').merge().setValue('Batas Stok Kritis (Unit)').setHorizontalAlignment('right').setFontColor(C.muted);
  R('K5').setValue(10).setBackground('#fef9c3').setFontWeight('bold').setHorizontalAlignment('center');

  R('B4:N5').setBorder(true, true, true, true, false, false, C.line, SpreadsheetApp.BorderStyle.SOLID);

  // Helper Kartu KPI & Header Section
  function kpi(labelRange, valueRange, label, formula, color, numFmt) {
    R(labelRange).merge().setValue(label).setFontSize(8).setFontWeight('bold').setFontColor(C.muted).setHorizontalAlignment('center').setBackground('#ffffff');
    var vCell = R(valueRange).merge().setFormula(formula).setFontSize(17).setFontColor(color).setHorizontalAlignment('center').setFontWeight('bold').setBackground('#ffffff');
    if (numFmt) vCell.setNumberFormat(numFmt);
    R(labelRange.split(':')[0] + ':' + valueRange.split(':')[1]).setBorder(true, true, true, true, false, false, C.line, SpreadsheetApp.BorderStyle.SOLID);
  }

  function section(range, title) {
    R(range).merge().setValue('  ' + title).setBackground(C.dark).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
  }

  var dateCondLog = 'OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2';
  var dateCondBahan = '_DASH_DATA!$S$2:$S,">="&_DASH_DATA!$E$2,_DASH_DATA!$S$2:$S,"<"&_DASH_DATA!$F$2';
  var dateCondQc = '_DASH_DATA!$AD$2:$AD,">="&_DASH_DATA!$E$2,_DASH_DATA!$AD$2:$AD,"<"&_DASH_DATA!$F$2';
  var dateCondHpp = '_DASH_DATA!$AM$2:$AM,">="&_DASH_DATA!$E$2,_DASH_DATA!$AM$2:$AM,"<"&_DASH_DATA!$F$2';

  // --- KPI BARIS 1: GUDANG & PENJUALAN ---
  kpi('B7:C7', 'B8:C8', 'PRODUK AKTIF', '=COUNTIF(MASTER_PRODUK!$C$2:$C,TRUE)', C.indigo, '#,##0');
  kpi('D7:E7', 'D8:E8', 'VARIASI AKTIF', '=COUNTIF(VARIASI_PRODUK!$E$2:$E,TRUE)', C.sky, '#,##0');
  kpi('F7:G7', 'F8:G8', 'TOTAL STOK GUDANG', '=IF($C$5="Semua Lokasi",SUM(STOK_VARIASI!$D$2:$D),SUMIFS(STOK_VARIASI!$D$2:$D,STOK_VARIASI!$C$2:$C,$C$5))', C.text, '#,##0');
  kpi('I7:J7', 'I8:J8', 'BARANG MASUK (PCS)', '=IF($C$5="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",' + dateCondLog + '),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",' + dateCondLog + ',OPERATION_LOG!$F$2:$F,$C$5))', C.green, '#,##0');
  kpi('K7:L7', 'K8:L8', 'PENJUALAN / KELUAR (PCS)', '=IF($C$5="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",' + dateCondLog + '),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",' + dateCondLog + ',OPERATION_LOG!$E$2:$E,$C$5))', C.red, '#,##0');
  kpi('M7:N7', 'M8:N8', 'STOK KRITIS (<= BATAS)', '=IFERROR(ROWS(QUERY(QUERY(STOK_VARIASI!$A$1:$D,"select A, sum(D) where A is not null "&IF($C$5="Semua Lokasi","","and C=\'"&$C$5&"\' ")&"group by A",1),"select Col1 where Col2 <= "&$K$5,1))-1,0)', C.amber, '#,##0');

  // --- KPI BARIS 2: PRODUKSI, BAHAN, UPAH & HPP ---
  kpi('B10:C10', 'B11:C11', 'WO JAHIT AKTIF', '=COUNTIFS(WO_JAHIT!$C$2:$C,"<>COMPLETED",WO_JAHIT!$C$2:$C,"<>CANCELLED",WO_JAHIT!$A$2:$A,"<>")', C.sky, '#,##0');
  kpi('D10:E10', 'D11:E11', 'NILAI BAHAN TERKIRIM', '=IF($F$5="Semua Tim",SUMIFS(_DASH_DATA!$AA$2:$AA,' + dateCondBahan + '),SUMIFS(_DASH_DATA!$AA$2:$AA,' + dateCondBahan + ',_DASH_DATA!$V$2:$V,$F$5))', C.indigo, '"Rp "#,##0');
  kpi('F10:G10', 'F11:G11', 'LOLOS QC (PCS)', '=IF($F$5="Semua Tim",SUMIFS(_DASH_DATA!$AH$2:$AH,' + dateCondQc + '),SUMIFS(_DASH_DATA!$AH$2:$AH,' + dateCondQc + ',_DASH_DATA!$AG$2:$AG,$F$5))', C.green, '#,##0');
  kpi('I10:J10', 'I11:J11', 'UPAH BELUM DIBAYAR', '=IF($F$5="Semua Tim",SUMIFS(_DASH_DATA!$AJ$2:$AJ,_DASH_DATA!$AK$2:$AK,"UNPAID"),SUMIFS(_DASH_DATA!$AJ$2:$AJ,_DASH_DATA!$AK$2:$AK,"UNPAID",_DASH_DATA!$AG$2:$AG,$F$5))', C.red, '"Rp "#,##0');
  kpi('K10:L10', 'K11:L11', 'UPAH SUDAH DIBAYAR', '=IF($F$5="Semua Tim",SUMIFS(REKAP_UPAH!$F$2:$F,REKAP_UPAH!$H$2:$H,">="&_DASH_DATA!$E$2,REKAP_UPAH!$H$2:$H,"<"&_DASH_DATA!$F$2),SUMIFS(REKAP_UPAH!$F$2:$F,REKAP_UPAH!$H$2:$H,">="&_DASH_DATA!$E$2,REKAP_UPAH!$H$2:$H,"<"&_DASH_DATA!$F$2,REKAP_UPAH!$B$2:$B,$F$5))', C.violet, '"Rp "#,##0');
  kpi('M10:N10', 'M11:N11', 'RATA-RATA HPP / PCS', '=IFERROR(IF($F$5="Semua Tim",AVERAGEIFS(_DASH_DATA!$AT$2:$AT,' + dateCondHpp + '),AVERAGEIFS(_DASH_DATA!$AT$2:$AT,' + dateCondHpp + ',_DASH_DATA!$AP$2:$AP,$F$5)),0)', C.text, '"Rp "#,##0');

  // ============================================================================
  // BLOK 1: TOP PENJUALAN / KELUAR & STOK KRITIS (Rows 13..24)
  // ============================================================================
  section('B13:G13', 'TOP 10 PENJUALAN / VARIASI KELUAR (PERIODE TERPILIH)');
  R('B14:G14').setValues([['Produk / Variasi', 'Barcode', 'Trx', 'Total Qty', 'Porsi', 'Visual']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('B15').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$J$1:$Q,"select M, L, count(J), sum(N) where K=\'Keluar\' and J >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and J < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($C$5="Semua Lokasi","","and O=\'"&$C$5&"\' ")&"group by M, L order by sum(N) desc limit 10 label M \'\', L \'\', count(J) \'\', sum(N) \'\'",0),{"","","",""})'
  );
  for (var r = 15; r <= 24; r++) {
    R('B' + r + ':G' + r).setBackground('#ffffff');
    R('F' + r).setFormula('=IF(E' + r + '="","",IFERROR(E' + r + '/SUM($E$15:$E$24),0))').setNumberFormat('0%');
    R('G' + r).setFormula(sparkFormulaV4_('E', r, 15, 24, C.red));
  }
  R('D15:E24').setNumberFormat('#,##0').setHorizontalAlignment('center');
  R('C15:C24').setHorizontalAlignment('center');
  R('F15:F24').setHorizontalAlignment('center');

  section('I13:N13', 'PERINGATAN STOK KRITIS & HABIS PER VARIASI');
  R('I14:N14').setValues([['ID Variasi', 'Stok', 'Nama Produk', 'Variasi', 'Status', 'Visual']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('I15').setFormula(
    '=IFERROR(QUERY(QUERY(STOK_VARIASI!$A$1:$D,"select A, sum(D) where A is not null "&IF($C$5="Semua Lokasi","","and C=\'"&$C$5&"\' ")&"group by A label A \'\', sum(D) \'\'",1),"select Col1, Col2 where Col2 <= "&$K$5&" order by Col2 asc limit 10",0),{"",""})'
  );
  for (r = 15; r <= 24; r++) {
    R('I' + r + ':N' + r).setBackground('#ffffff');
    R('K' + r).setFormula('=IF(I' + r + '="","",IFERROR(VLOOKUP(VLOOKUP(I' + r + ',VARIASI_PRODUK!$A:$C,2,FALSE),MASTER_PRODUK!$A:$B,2,FALSE),""))');
    R('L' + r).setFormula('=IF(I' + r + '="","",IFERROR(VLOOKUP(I' + r + ',VARIASI_PRODUK!$A:$C,3,FALSE),""))');
    R('M' + r).setFormula('=IF(J' + r + '="","",IF(J' + r + '<=0,"HABIS","MENIPIS"))').setHorizontalAlignment('center').setFontWeight('bold');
    R('N' + r).setFormula(sparkFormulaV4_('J', r, 15, 24, C.amber));
  }
  R('J15:J24').setNumberFormat('#,##0').setHorizontalAlignment('center');

  // ============================================================================
  // BLOK 2: GRAFIK TREN 14 HARI & PERSEBARAN LOKASI / AKTIVITAS (Rows 26..41)
  // ============================================================================
  section('B26:G26', 'TREN BARANG MASUK VS KELUAR · 14 HARI TERAKHIR');
  R('B27:G40').setBackground('#ffffff');
  R('B41:G41').merge().setFormula('=IF(COUNTA(OPERATION_LOG!A2:A)=0,"Belum ada transaksi. Grafik akan bergerak setelah transaksi pertama.","Grafik menampilkan jumlah unit Masuk vs Keluar per hari selama 14 hari terakhir.")')
    .setFontColor(C.muted).setFontSize(8).setHorizontalAlignment('center').setBackground('#ffffff');

  section('I26:N26', 'PERSEBARAN STOK PER LOKASI GUDANG');
  R('I27:N27').setValues([['Lokasi', 'Total Unit', 'Variasi', 'Porsi', 'Visual', 'Keterangan']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('I28').setFormula(
    '=IFERROR(QUERY(STOK_VARIASI!$A$1:$D,"select C, sum(D), count(A) where C is not null "&IF($C$5="Semua Lokasi","","and C=\'"&$C$5&"\' ")&"group by C order by sum(D) desc limit 6 label C \'\', sum(D) \'\', count(A) \'\'",1),{"","",""})'
  );
  for (r = 28; r <= 33; r++) {
    R('I' + r + ':N' + r).setBackground('#ffffff');
    R('L' + r).setFormula('=IF(J' + r + '="","",IFERROR(J' + r + '/SUM($J$28:$J$33),0))').setNumberFormat('0%').setHorizontalAlignment('center');
    R('M' + r).setFormula(sparkFormulaV4_('J', r, 28, 33, C.indigo));
    R('N' + r).setFormula('=IF(I' + r + '="","","Aktif")').setHorizontalAlignment('center').setFontColor(C.green);
  }
  R('J28:K33').setNumberFormat('#,##0').setHorizontalAlignment('center');

  section('I35:N35', 'RINGKASAN AKTIVITAS TRANSAKSI GUDANG');
  R('I36:N36').setValues([['Aktivitas', 'Transaksi', 'Total Unit', 'Porsi Trx', 'Visual', 'Status']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('I37').setFormula(
    '=IFERROR(QUERY(OPERATION_LOG!$A$1:$N,"select C, count(B), sum(G) where C is not null and A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($C$5="Semua Lokasi","","and (E=\'"&$C$5&"\' or F=\'"&$C$5&"\') ")&"group by C order by count(B) desc limit 5 label C \'\', count(B) \'\', sum(G) \'\'",1),{"","",""})'
  );
  for (r = 37; r <= 41; r++) {
    R('I' + r + ':N' + r).setBackground('#ffffff');
    R('L' + r).setFormula('=IF(J' + r + '="","",IFERROR(J' + r + '/SUM($J$37:$J$41),0))').setNumberFormat('0%').setHorizontalAlignment('center');
    R('M' + r).setFormula(sparkFormulaV4_('J', r, 37, 41, C.sky));
    R('N' + r).setFormula('=IF(I' + r + '="","","Tercatat")').setHorizontalAlignment('center').setFontColor(C.muted);
  }
  R('J37:K41').setNumberFormat('#,##0').setHorizontalAlignment('center');

  // ============================================================================
  // BLOK 3: REKAP BAHAN BAKU & PENGIRIMAN BAHAN TERBARU (Rows 43..54)
  // ============================================================================
  section('B43:G43', 'REKAP PENGGUNAAN & MODAL BAHAN BAKU (PERIODE TERPILIH)');
  R('B44:G44').setValues([['Kode Bahan', 'Nama Bahan', 'Qty Kirim', 'Qty Terima', 'Modal Bahan (Rp)', 'Visual']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('B45').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$S$1:$AB,"select W, X, sum(Y), sum(Z), sum(AA) where S is not null and S >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and S < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($F$5="Semua Tim","","and V=\'"&$F$5&"\' ")&"group by W, X order by sum(AA) desc limit 10 label W \'\', X \'\', sum(Y) \'\', sum(Z) \'\', sum(AA) \'\'",0),{"","","","",""})'
  );
  for (r = 45; r <= 54; r++) {
    R('B' + r + ':G' + r).setBackground('#ffffff');
    R('G' + r).setFormula(sparkFormulaV4_('F', r, 45, 54, C.indigo));
  }
  R('D45:E54').setNumberFormat('#,##0.##').setHorizontalAlignment('center');
  R('F45:F54').setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');

  section('I43:N43', 'STATUS SURAT JALAN / PENGIRIMAN BAHAN TERBARU');
  R('I44:N44').setValues([['Delivery ID', 'No WO', 'Kode Bahan', 'Qty Kirim', 'Qty Terima', 'Status']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('I45').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$S$1:$AB,"select T, U, W, Y, Z, AB where S is not null and S >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and S < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($F$5="Semua Tim","","and V=\'"&$F$5&"\' ")&"order by S desc limit 10 label T \'\', U \'\', W \'\', Y \'\', Z \'\', AB \'\'",0),{"","","","","",""})'
  );
  R('I45:N54').setBackground('#ffffff');
  R('L45:M54').setNumberFormat('#,##0.##').setHorizontalAlignment('center');
  R('N45:N54').setHorizontalAlignment('center').setFontWeight('bold');

  // ============================================================================
  // BLOK 4: MONITORING WORK ORDER (WO) JAHIT & HASIL QC (Rows 56..67)
  // ============================================================================
  section('B56:G56', 'MONITORING WORK ORDER (WO) JAHIT');
  R('B57:G57').setValues([['Nomor WO', 'Tim Penjahit', 'Status WO', 'Target Qty', 'Lolos QC', 'Reject QC']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('B58').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$AV$1:$BC,"select AX, AY, AZ, BA, BB, BC where AV is not null and AV >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and AV < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($F$5="Semua Tim","","and AY=\'"&$F$5&"\' ")&"order by AV desc limit 10 label AX \'\', AY \'\', AZ \'\', BA \'\', BB \'\', BC \'\'",0),{"","","","","",""})'
  );
  R('B58:G67').setBackground('#ffffff');
  R('D58:D67').setHorizontalAlignment('center').setFontWeight('bold');
  R('E58:G67').setNumberFormat('#,##0').setHorizontalAlignment('center');

  section('I56:N56', 'HASIL QUALITY CONTROL (QC) & PERHITUNGAN UPAH TERBARU');
  R('I57:N57').setValues([['QC ID', 'No WO', 'Tim Penjahit', 'Lolos', 'Nilai Upah (Rp)', 'Status Bayar']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('I58').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$AD$1:$AK,"select AE, AF, AG, AH, AJ, AK where AD is not null and AD >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and AD < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($F$5="Semua Tim","","and AG=\'"&$F$5&"\' ")&"order by AD desc limit 10 label AE \'\', AF \'\', AG \'\', AH \'\', AJ \'\', AK \'\'",0),{"","","","","",""})'
  );
  R('I58:N67').setBackground('#ffffff');
  R('L58:L67').setNumberFormat('#,##0').setHorizontalAlignment('center');
  R('M58:M67').setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
  R('N58:N67').setHorizontalAlignment('center').setFontWeight('bold');

  // ============================================================================
  // BLOK 5: REKAP UPAH PER TIM PENJAHIT & SLIP GAJI (Rows 69..80)
  // ============================================================================
  section('B69:G69', 'REKAPITULASI UPAH PER TIM PENJAHIT (PERIODE TERPILIH)');
  R('B70:G70').setValues([['Tim Penjahit', 'Pcs Lolos', 'Total Upah (Rp)', 'Belum Bayar (Rp)', 'Sudah Bayar (Rp)', 'Visual']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('B71').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$AD$1:$AK,"select AG, sum(AH), sum(AJ) where AD is not null and AD >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and AD < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($F$5="Semua Tim","","and AG=\'"&$F$5&"\' ")&"group by AG order by sum(AJ) desc limit 10 label AG \'\', sum(AH) \'\', sum(AJ) \'\'",0),{"","",""})'
  );
  for (r = 71; r <= 80; r++) {
    R('B' + r + ':G' + r).setBackground('#ffffff');
    R('E' + r).setFormula('=IF(B' + r + '="","",SUMIFS(_DASH_DATA!$AJ$2:$AJ,_DASH_DATA!$AG$2:$AG,B' + r + ',_DASH_DATA!$AK$2:$AK,"UNPAID",' + dateCondQc + '))');
    R('F' + r).setFormula('=IF(B' + r + '="","",SUMIFS(_DASH_DATA!$AJ$2:$AJ,_DASH_DATA!$AG$2:$AG,B' + r + ',_DASH_DATA!$AK$2:$AK,"PAID",' + dateCondQc + '))');
    R('G' + r).setFormula(sparkFormulaV4_('D', r, 71, 80, C.violet));
  }
  R('C71:C80').setNumberFormat('#,##0').setHorizontalAlignment('center');
  R('D71:F80').setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');

  section('I69:N69', 'RIWAYAT SLIP GAJI PENJAHIT TERBARU (REKAP_UPAH)');
  R('I70:N70').setValues([['Slip ID', 'Tim Penjahit', 'Periode Mulai', 'Qty Lolos', 'Total Upah (Rp)', 'Tgl Bayar']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('I71').setFormula(
    '=IFERROR(QUERY(REKAP_UPAH!$A$1:$J,"select A, B, C, E, F, H where A is not null and H >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and H < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($F$5="Semua Tim","","and B=\'"&$F$5&"\' ")&"order by H desc limit 10 label A \'\', B \'\', C \'\', E \'\', F \'\', H \'\'",1),{"","","","","",""})'
  );
  R('I71:N80').setBackground('#ffffff');
  R('K71:K80').setNumberFormat('dd MMM yyyy').setHorizontalAlignment('center');
  R('L71:L80').setNumberFormat('#,##0').setHorizontalAlignment('center');
  R('M71:M80').setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
  R('N71:N80').setNumberFormat('dd MMM yyyy').setHorizontalAlignment('center');

  // ============================================================================
  // BLOK 6: LAPORAN HPP PRODUKSI PER BATCH & LOG TRANSAKSI (Rows 82..93)
  // ============================================================================
  section('B82:G82', 'LAPORAN HPP PRODUKSI PER BATCH (HPP_PRODUKSI)');
  R('B83:G83').setValues([['Nomor WO', 'Tim Penjahit', 'Cost Bahan', 'Cost Jahit', 'Qty Lolos', 'HPP / Pcs']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('B84').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$AM$1:$AT,"select AO, AP, AQ, AR, AS, AT where AM is not null and AM >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and AM < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($F$5="Semua Tim","","and AP=\'"&$F$5&"\' ")&"order by AM desc limit 10 label AO \'\', AP \'\', AQ \'\', AR \'\', AS \'\', AT \'\'",0),{"","","","","",""})'
  );
  R('B84:G93').setBackground('#ffffff');
  R('D84:E93').setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
  R('F84:F93').setNumberFormat('#,##0').setHorizontalAlignment('center');
  R('G84:G93').setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right').setFontWeight('bold');

  section('I82:N82', '10 LOG TRANSAKSI GUDANG TERBARU');
  R('I83:N83').setValues([['Waktu', 'Aktivitas', 'Produk / Variasi', 'Qty', 'Tujuan / Asal', 'Pelapor']]).setBackground('#e2e8f0').setFontWeight('bold');
  R('I84').setFormula(
    '=IFERROR(QUERY(_DASH_DATA!$J$1:$Q,"select J, K, M, N, P, Q where J is not null "&IF($C$5="Semua Lokasi","","and (O=\'"&$C$5&"\' or P=\'"&$C$5&"\') ")&"order by J desc limit 10 label J \'\', K \'\', M \'\', N \'\', P \'\', Q \'\'",0),{"","","","","",""})'
  );
  R('I84:N93').setBackground('#ffffff');
  R('I84:I93').setNumberFormat('dd MMM HH:mm').setHorizontalAlignment('center');
  R('J84:J93').setHorizontalAlignment('center');
  R('L84:L93').setNumberFormat('#,##0').setHorizontalAlignment('center');

  // --- CHART TREN 14 HARI ---
  var chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataSh.getRange('A1:C15'))
    .setNumHeaders(1)
    .setPosition(27, 2, 4, 4)
    .setOption('title', 'Unit Barang Masuk vs Keluar (14 Hari Terakhir)')
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#10b981', '#f43f5e'])
    .setOption('backgroundColor', '#ffffff')
    .setOption('vAxis', { minValue: 0, viewWindow: { min: 0 }, format: '0' })
    .setOption('hAxis', { slantedText: false })
    .setOption('width', 698)
    .setOption('height', 298)
    .build();
  sh.insertChart(chart);

  // --- CONDITIONAL FORMATTING ---
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('HABIS').setBackground('#fee2e2').setFontColor('#b91c1c').setRanges([R('M15:M24')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('MENIPIS').setBackground('#fef3c7').setFontColor('#b45309').setRanges([R('M15:M24')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('IN_TRANSIT').setBackground('#fef3c7').setFontColor('#b45309').setRanges([R('N45:N54')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('RECEIVED').setBackground('#dcfce7').setFontColor('#15803d').setRanges([R('N45:N54')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('UNPAID').setBackground('#fee2e2').setFontColor('#b91c1c').setRanges([R('N58:N67')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PAID').setBackground('#dcfce7').setFontColor('#15803d').setRanges([R('N58:N67')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('COMPLETED').setBackground('#dcfce7').setFontColor('#15803d').setRanges([R('D58:D67')]).build()
  ];
  sh.setConditionalFormatRules(rules);
}

// ============================================================================
// 3. GENERATOR TARIK LAPORAN PERIODIK (HARIAN / MINGGUAN / BULANAN / CUSTOM)
// ============================================================================
function openReportDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locations = ['Semua Lokasi'];
  var shLoc = ss.getSheetByName('MASTER_LOKASI');
  if (shLoc && shLoc.getLastRow() >= 2) {
    shLoc.getRange(2, 2, shLoc.getLastRow() - 1, 3).getValues().forEach(function (r) {
      if (r[0] && r[2] !== false && locations.indexOf(String(r[0])) === -1) locations.push(String(r[0]));
    });
  }

  var teams = ['Semua Tim'];
  var shWo = ss.getSheetByName('WO_JAHIT');
  if (shWo && shWo.getLastRow() >= 2) {
    shWo.getRange(2, 4, shWo.getLastRow() - 1, 1).getValues().forEach(function (r) {
      if (r[0] && teams.indexOf(String(r[0])) === -1) teams.push(String(r[0]));
    });
  }

  var locOpts = locations.map(function (l) { return '<option value="' + l + '">' + l + '</option>'; }).join('');
  var timOpts = teams.map(function (t) { return '<option value="' + t + '">' + t + '</option>'; }).join('');

  var todayStr = Utilities.formatDate(new Date(), IBS_DASH.TZ, 'yyyy-MM-dd');
  var d30 = new Date();
  d30.setDate(d30.getDate() - 29);
  var startStr = Utilities.formatDate(d30, IBS_DASH.TZ, 'yyyy-MM-dd');

  var html = [
    '<!DOCTYPE html><html><head><base target="_top"><style>',
    'body{font-family:Roboto,Arial,sans-serif;padding:16px;color:#1e293b;background:#f8fafc;margin:0;}',
    'label{display:block;font-size:11px;font-weight:700;text-transform:uppercase;color:#64748b;margin-bottom:4px;margin-top:12px;}',
    'select,input{width:100%;padding:9px 10px;border:1px solid #cbd5e1;border-radius:8px;font-size:13px;box-sizing:border-box;background:#fff;}',
    '.row{display:flex;gap:10px;}.col{flex:1;}',
    'button{width:100%;margin-top:18px;padding:12px;background:#0f172a;color:#fff;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;}',
    'button:disabled{background:#94a3b8;}',
    '.hint{font-size:11px;color:#64748b;margin-top:8px;line-height:1.4;}',
    '</style></head><body>',
    '<label>Mode Pengelompokan Rekap</label>',
    '<select id="groupBy">',
    '  <option value="DAILY">Harian (Per Tanggal)</option>',
    '  <option value="WEEKLY">Mingguan (Per Minggu Sen–Min)</option>',
    '  <option value="MONTHLY">Bulanan (Per Bulan)</option>',
    '</select>',
    '<label>Preset Rentang Waktu</label>',
    '<select id="preset" onchange="applyPreset()">',
    '  <option value="30D">30 Hari Terakhir</option>',
    '  <option value="TODAY">Hari Ini</option>',
    '  <option value="7D">7 Hari Terakhir (1 Minggu)</option>',
    '  <option value="THIS_MONTH">Bulan Ini</option>',
    '  <option value="LAST_MONTH">Bulan Lalu</option>',
    '  <option value="90D">3 Bulan Terakhir</option>',
    '  <option value="ALL">Semua Waktu</option>',
    '</select>',
    '<div class="row">',
    '  <div class="col"><label>Tanggal Mulai</label><input type="date" id="startDate" value="' + startStr + '"></div>',
    '  <div class="col"><label>Tanggal Akhir</label><input type="date" id="endDate" value="' + todayStr + '"></div>',
    '</div>',
    '<div class="row">',
    '  <div class="col"><label>Filter Lokasi</label><select id="lokasi">' + locOpts + '</select></div>',
    '  <div class="col"><label>Filter Tim Penjahit</label><select id="tim">' + timOpts + '</select></div>',
    '</div>',
    '<p class="hint">Hasil akan dicetak ke sheet <b>LAPORAN_REKAP</b> lengkap dengan ringkasan KPI, tabel periodik, penjualan, bahan, upah, dan HPP.</p>',
    '<button id="btnRun" onclick="runReport()">Tarik & Tampilkan Laporan</button>',
    '<script>',
    'function fmt(d){return d.toISOString().slice(0,10);}',
    'function applyPreset(){',
    '  var p=document.getElementById("preset").value, now=new Date(), s=new Date(), e=new Date();',
    '  if(p==="TODAY"){}',
    '  else if(p==="7D"){s.setDate(now.getDate()-6);}',
    '  else if(p==="30D"){s.setDate(now.getDate()-29);}',
    '  else if(p==="90D"){s.setDate(now.getDate()-89);}',
    '  else if(p==="THIS_MONTH"){s=new Date(now.getFullYear(),now.getMonth(),1);}',
    '  else if(p==="LAST_MONTH"){s=new Date(now.getFullYear(),now.getMonth()-1,1);e=new Date(now.getFullYear(),now.getMonth(),0);}',
    '  else if(p==="ALL"){s=new Date(2020,0,1);}',
    '  document.getElementById("startDate").value=fmt(s);',
    '  document.getElementById("endDate").value=fmt(e);',
    '}',
    'function runReport(){',
    '  var btn=document.getElementById("btnRun"); btn.disabled=true; btn.textContent="Menyusun Laporan...";',
    '  var params={',
    '    groupBy:document.getElementById("groupBy").value,',
    '    startDate:document.getElementById("startDate").value,',
    '    endDate:document.getElementById("endDate").value,',
    '    lokasi:document.getElementById("lokasi").value,',
    '    tim:document.getElementById("tim").value',
    '  };',
    '  google.script.run.withSuccessHandler(function(){google.script.host.close();}).withFailureHandler(function(err){alert(err.message);btn.disabled=false;btn.textContent="Tarik & Tampilkan Laporan";}).generatePeriodicReport(params);',
    '}',
    '</script></body></html>'
  ].join('');

  var ui = HtmlService.createHtmlOutput(html).setWidth(430).setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Tarik Laporan Lengkap (Harian / Mingguan / Bulanan)');
}

function generatePeriodicReport(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(IBS_DASH.TZ);

  var groupBy = (params && params.groupBy) || 'DAILY';
  var startStr = (params && params.startDate) || '2020-01-01';
  var endStr = (params && params.endDate) || Utilities.formatDate(new Date(), IBS_DASH.TZ, 'yyyy-MM-dd');
  var filterLokasi = (params && params.lokasi) || 'Semua Lokasi';
  var filterTim = (params && params.tim) || 'Semua Tim';

  var startDt = new Date(startStr + 'T00:00:00+07:00');
  var endDt = new Date(endStr + 'T23:59:59+07:00');

  function inRange(dtVal) {
    if (!dtVal) return false;
    var d = dtVal instanceof Date ? dtVal : new Date(dtVal);
    if (isNaN(d.getTime())) return false;
    return d >= startDt && d <= endDt;
  }

  function periodKey(dtVal) {
    var d = dtVal instanceof Date ? dtVal : new Date(dtVal);
    if (groupBy === 'MONTHLY') {
      return Utilities.formatDate(d, IBS_DASH.TZ, 'yyyy-MM');
    }
    if (groupBy === 'WEEKLY') {
      var day = d.getDay(); // 0=Sun..6=Sat
      var diffToMon = (day === 0 ? -6 : 1 - day);
      var mon = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diffToMon);
      var sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6);
      return Utilities.formatDate(mon, IBS_DASH.TZ, 'yyyy-MM-dd') + ' s/d ' + Utilities.formatDate(sun, IBS_DASH.TZ, 'dd MMM yyyy');
    }
    return Utilities.formatDate(d, IBS_DASH.TZ, 'yyyy-MM-dd');
  }

  function readRows(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return [];
    var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var headers = vals[0].map(function (h) { return String(h).trim(); });
    return vals.slice(1).map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { if (h) obj[h] = row[i]; });
      return obj;
    });
  }

  // Maps referensi
  var bahanMap = {};
  readRows('MASTER_BAHAN').forEach(function (b) {
    bahanMap[String(b.Kode_Bahan).toUpperCase()] = {
      nama: String(b.Nama_Bahan || b.Kode_Bahan),
      satuan: String(b.Satuan || 'unit'),
      cost: Number(b.Estimasi_Cost) || 0
    };
  });

  var woMap = {};
  readRows('WO_JAHIT').forEach(function (w) {
    woMap[String(w.WO_ID)] = {
      nomor_wo: String(w.Nomor_WO || w.WO_ID),
      tim: String(w.Tim_Label || '-')
    };
  });

  var buckets = {};
  function ensureBucket(k) {
    if (!buckets[k]) {
      buckets[k] = {
        periode: k,
        masuk: 0,
        keluar: 0,
        bahan_qty: 0,
        bahan_cost: 0,
        qc_lolos: 0,
        qc_gagal: 0,
        upah_terbit: 0,
        upah_dibayar: 0,
        hpp_batch: 0
      };
    }
    return buckets[k];
  }

  // 1. OPERATION_LOG (Masuk & Penjualan/Keluar)
  var salesByProduct = {};
  var totMasuk = 0, totKeluar = 0;
  readRows('OPERATION_LOG').forEach(function (log) {
    if (!inRange(log.Timestamp)) return;
    var act = String(log.Aktivitas || '');
    var qty = Number(log.Qty) || 0;
    var asal = String(log.Lokasi_Asal || '');
    var tujuan = String(log.Lokasi_Tujuan || '');
    if (filterLokasi !== 'Semua Lokasi' && asal !== filterLokasi && tujuan !== filterLokasi) return;

    var pk = periodKey(log.Timestamp);
    var b = ensureBucket(pk);
    if (act === 'Masuk') {
      b.masuk += qty;
      totMasuk += qty;
    } else if (act === 'Keluar') {
      b.keluar += qty;
      totKeluar += qty;
      var prodLabel = String(log.Produk_Tampilan || log.Barcode || '-');
      var bc = String(log.Barcode || '-');
      var sKey = bc + '|' + prodLabel;
      if (!salesByProduct[sKey]) salesByProduct[sKey] = { barcode: bc, produk: prodLabel, trx: 0, qty: 0 };
      salesByProduct[sKey].trx += 1;
      salesByProduct[sKey].qty += qty;
    }
  });

  // 2. PENGIRIMAN_BAHAN (Bahan Baku)
  var bahanUsage = {};
  var totBahanQty = 0, totBahanCost = 0;
  readRows('PENGIRIMAN_BAHAN').forEach(function (d) {
    if (!inRange(d.Created_At)) return;
    var wo = woMap[String(d.WO_ID)] || { nomor_wo: String(d.WO_ID), tim: '-' };
    if (filterTim !== 'Semua Tim' && wo.tim !== filterTim) return;

    var kode = String(d.Kode_Bahan || '-').toUpperCase();
    var info = bahanMap[kode] || { nama: kode, satuan: 'unit', cost: 0 };
    var qKirim = Number(d.Qty_Aktual_Kirim) || 0;
    var qTerima = Number(d.Qty_Terima) || 0;
    var qPakai = qTerima > 0 ? qTerima : qKirim;
    var modal = qPakai * info.cost;

    var pk = periodKey(d.Created_At);
    var b = ensureBucket(pk);
    b.bahan_qty += qPakai;
    b.bahan_cost += modal;
    totBahanQty += qPakai;
    totBahanCost += modal;

    if (!bahanUsage[kode]) bahanUsage[kode] = { kode: kode, nama: info.nama, satuan: info.satuan, harga: info.cost, kirim: 0, terima: 0, modal: 0 };
    bahanUsage[kode].kirim += qKirim;
    bahanUsage[kode].terima += qTerima;
    bahanUsage[kode].modal += modal;
  });

  // 3. PENERIMAAN_QC (Hasil QC & Upah Terbit)
  var upahByTim = {};
  var totLolos = 0, totGagal = 0, totUpahTerbit = 0;
  readRows('PENERIMAAN_QC').forEach(function (q) {
    if (!inRange(q.Created_At)) return;
    var tim = String(q.Tim_Penjahit || 'Tanpa Tim');
    if (filterTim !== 'Semua Tim' && tim !== filterTim) return;

    var lolos = Number(q.Qty_Lolos) || 0;
    var gagal = Number(q.Qty_Gagal) || 0;
    var upah = Number(q.Nilai_Upah) || 0;
    var st = String(q.Status_Pembayaran || 'UNPAID').toUpperCase();

    var pk = periodKey(q.Created_At);
    var b = ensureBucket(pk);
    b.qc_lolos += lolos;
    b.qc_gagal += gagal;
    b.upah_terbit += upah;
    totLolos += lolos;
    totGagal += gagal;
    totUpahTerbit += upah;

    if (!upahByTim[tim]) upahByTim[tim] = { tim: tim, lolos: 0, gagal: 0, unpaid: 0, paid: 0, total: 0 };
    upahByTim[tim].lolos += lolos;
    upahByTim[tim].gagal += gagal;
    if (st === 'PAID') upahByTim[tim].paid += upah;
    else upahByTim[tim].unpaid += upah;
    upahByTim[tim].total += upah;
  });

  // 4. REKAP_UPAH (Slip Gaji Dibayar)
  var totUpahDibayar = 0;
  readRows('REKAP_UPAH').forEach(function (u) {
    var tgl = u.Tanggal_Bayar || u.Created_At;
    if (!inRange(tgl)) return;
    var tim = String(u.Nama_Tim_Penjahit || 'Tanpa Tim');
    if (filterTim !== 'Semua Tim' && tim !== filterTim) return;

    var nom = Number(u.Total_Upah) || 0;
    var pk = periodKey(tgl);
    var b = ensureBucket(pk);
    b.upah_dibayar += nom;
    totUpahDibayar += nom;
  });

  // 5. HPP_PRODUKSI
  var hppList = [];
  var sumHppUnit = 0;
  readRows('HPP_PRODUKSI').forEach(function (h) {
    if (!inRange(h.Created_At)) return;
    var wo = woMap[String(h.WO_ID)] || { nomor_wo: String(h.WO_ID), tim: '-' };
    if (filterTim !== 'Semua Tim' && wo.tim !== filterTim) return;

    var pk = periodKey(h.Created_At);
    var b = ensureBucket(pk);
    b.hpp_batch += 1;

    var hppPcs = Number(h.HPP_Per_Pcs) || 0;
    sumHppUnit += hppPcs;
    hppList.push({
      tanggal: h.Created_At,
      hpp_id: String(h.HPP_ID),
      nomor_wo: wo.nomor_wo,
      tim: wo.tim,
      bahan: Number(h.Total_Cost_Bahan) || 0,
      jahit: Number(h.Total_Cost_Jahit) || 0,
      qty: Number(h.Qty_Lolos_QC) || 0,
      hpp_pcs: hppPcs
    });
  });

  var avgHpp = hppList.length ? Math.round(sumHppUnit / hppList.length) : 0;

  // ============================================================================
  // TULIS HASIL KE SHEET LAPORAN_REKAP
  // ============================================================================
  var shRep = ss.getSheetByName(IBS_DASH.REPORT) || ss.insertSheet(IBS_DASH.REPORT);
  shRep.clear();
  shRep.setHiddenGridlines(true);

  var groupLabel = groupBy === 'MONTHLY' ? 'BULANAN' : (groupBy === 'WEEKLY' ? 'MINGGUAN' : 'HARIAN');
  shRep.getRange('A1:J1').merge().setValue('LAPORAN REKAPITULASI OPERASIONAL, BAHAN, UPAH & HPP (' + groupLabel + ')')
    .setBackground(IBS_DASH.COLOR.dark).setFontColor('#ffffff').setFontSize(14).setFontWeight('bold');
  shRep.getRange('A2:J2').merge().setValue(
    'Periode: ' + startStr + ' s/d ' + endStr + '  |  Lokasi: ' + filterLokasi + '  |  Tim Penjahit: ' + filterTim +
    '  |  Dicetak: ' + Utilities.formatDate(new Date(), IBS_DASH.TZ, 'dd MMM yyyy HH:mm') + ' WIB'
  ).setBackground(IBS_DASH.COLOR.slate).setFontColor('#cbd5e1').setFontSize(9);

  // Ringkasan KPI
  shRep.getRange('A4:J4').setValues([[
    'Masuk (Pcs)', 'Penjualan/Keluar', 'Bahan Terpakai', 'Modal Bahan (Rp)',
    'Lolos QC (Pcs)', 'Reject QC (Pcs)', 'Upah Terbit (Rp)', 'Upah Dibayar (Rp)', 'Batch HPP', 'Rata-rata HPP/Pcs'
  ]]).setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');

  shRep.getRange('A5:J5').setValues([[
    totMasuk, totKeluar, totBahanQty, totBahanCost,
    totLolos, totGagal, totUpahTerbit, totUpahDibayar, hppList.length, avgHpp
  ]]).setBackground('#ffffff').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  shRep.getRange('A5:C5').setNumberFormat('#,##0.##');
  shRep.getRange('D5').setNumberFormat('"Rp "#,##0');
  shRep.getRange('E5:F5').setNumberFormat('#,##0');
  shRep.getRange('G5:H5').setNumberFormat('"Rp "#,##0');
  shRep.getRange('I5').setNumberFormat('#,##0');
  shRep.getRange('J5').setNumberFormat('"Rp "#,##0');
  shRep.getRange('A4:J5').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  var curRow = 7;

  // TABEL 1: REKAPITULASI PERIODIK (HARIAN / MINGGUAN / BULANAN)
  shRep.getRange(curRow, 1, 1, 10).merge().setValue('1. REKAPITULASI PERIODIK (' + groupLabel + ')')
    .setBackground(IBS_DASH.COLOR.dark).setFontColor('#ffffff').setFontWeight('bold');
  curRow++;
  shRep.getRange(curRow, 1, 1, 10).setValues([[
    'Periode (' + groupLabel + ')', 'Masuk (Pcs)', 'Keluar / Jual (Pcs)', 'Bahan (Qty)',
    'Modal Bahan (Rp)', 'Lolos QC (Pcs)', 'Reject QC (Pcs)', 'Upah Terbit (Rp)', 'Upah Dibayar (Rp)', 'Batch HPP'
  ]]).setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');
  curRow++;

  var sortedKeys = Object.keys(buckets).sort();
  if (sortedKeys.length === 0) {
    shRep.getRange(curRow, 1, 1, 10).merge().setValue('Tidak ada transaksi pada rentang waktu & filter yang dipilih.').setHorizontalAlignment('center');
    curRow += 2;
  } else {
    var periodRows = sortedKeys.map(function (k) {
      var b = buckets[k];
      return [b.periode, b.masuk, b.keluar, b.bahan_qty, b.bahan_cost, b.qc_lolos, b.qc_gagal, b.upah_terbit, b.upah_dibayar, b.hpp_batch];
    });
    shRep.getRange(curRow, 1, periodRows.length, 10).setValues(periodRows).setBackground('#ffffff');
    shRep.getRange(curRow, 2, periodRows.length, 3).setNumberFormat('#,##0.##').setHorizontalAlignment('center');
    shRep.getRange(curRow, 5, periodRows.length, 1).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    shRep.getRange(curRow, 6, periodRows.length, 2).setNumberFormat('#,##0').setHorizontalAlignment('center');
    shRep.getRange(curRow, 8, periodRows.length, 2).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    shRep.getRange(curRow, 10, periodRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    curRow += periodRows.length;

    // Baris Total
    shRep.getRange(curRow, 1, 1, 10).setValues([[
      'TOTAL PERIODE', totMasuk, totKeluar, totBahanQty, totBahanCost, totLolos, totGagal, totUpahTerbit, totUpahDibayar, hppList.length
    ]]).setBackground('#fef9c3').setFontWeight('bold');
    shRep.getRange(curRow, 2, 1, 3).setNumberFormat('#,##0.##').setHorizontalAlignment('center');
    shRep.getRange(curRow, 5, 1, 1).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    shRep.getRange(curRow, 6, 1, 2).setNumberFormat('#,##0').setHorizontalAlignment('center');
    shRep.getRange(curRow, 8, 1, 2).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    shRep.getRange(curRow, 10, 1, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    curRow += 2;
  }

  // TABEL 2: RINCIAN PENJUALAN / BARANG KELUAR
  shRep.getRange(curRow, 1, 1, 5).merge().setValue('2. RINCIAN PENJUALAN / BARANG KELUAR PER PRODUK')
    .setBackground(IBS_DASH.COLOR.dark).setFontColor('#ffffff').setFontWeight('bold');
  curRow++;
  shRep.getRange(curRow, 1, 1, 5).setValues([['Barcode', 'Produk / Variasi', 'Jumlah Transaksi', 'Total Unit Keluar', 'Porsi (%)']])
    .setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');
  curRow++;
  var salesArr = Object.keys(salesByProduct).map(function (k) { return salesByProduct[k]; }).sort(function (a, b) { return b.qty - a.qty; });
  if (salesArr.length) {
    var sRows = salesArr.map(function (s) { return [s.barcode, s.produk, s.trx, s.qty, totKeluar > 0 ? s.qty / totKeluar : 0]; });
    shRep.getRange(curRow, 1, sRows.length, 5).setValues(sRows).setBackground('#ffffff');
    shRep.getRange(curRow, 3, sRows.length, 2).setNumberFormat('#,##0').setHorizontalAlignment('center');
    shRep.getRange(curRow, 5, sRows.length, 1).setNumberFormat('0.0%').setHorizontalAlignment('center');
    curRow += sRows.length + 1;
  } else {
    shRep.getRange(curRow, 1, 1, 5).merge().setValue('Belum ada data barang keluar pada periode ini.');
    curRow += 2;
  }

  // TABEL 3: RINCIAN PEMAKAIAN & MODAL BAHAN BAKU
  shRep.getRange(curRow, 1, 1, 7).merge().setValue('3. RINCIAN PEMAKAIAN & MODAL BAHAN BAKU')
    .setBackground(IBS_DASH.COLOR.dark).setFontColor('#ffffff').setFontWeight('bold');
  curRow++;
  shRep.getRange(curRow, 1, 1, 7).setValues([['Kode Bahan', 'Nama Bahan', 'Satuan', 'Harga/Satuan (Rp)', 'Qty Dikirim', 'Qty Diterima', 'Total Modal (Rp)']])
    .setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');
  curRow++;
  var bahanArr = Object.keys(bahanUsage).map(function (k) { return bahanUsage[k]; }).sort(function (a, b) { return b.modal - a.modal; });
  if (bahanArr.length) {
    var bRows = bahanArr.map(function (b) { return [b.kode, b.nama, b.satuan, b.harga, b.kirim, b.terima, b.modal]; });
    shRep.getRange(curRow, 1, bRows.length, 7).setValues(bRows).setBackground('#ffffff');
    shRep.getRange(curRow, 4, bRows.length, 1).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    shRep.getRange(curRow, 5, bRows.length, 2).setNumberFormat('#,##0.##').setHorizontalAlignment('center');
    shRep.getRange(curRow, 7, bRows.length, 1).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    curRow += bRows.length + 1;
  } else {
    shRep.getRange(curRow, 1, 1, 7).merge().setValue('Belum ada data distribusi bahan pada periode ini.');
    curRow += 2;
  }

  // TABEL 4: REKAPITULASI UPAH PENJAHIT
  shRep.getRange(curRow, 1, 1, 6).merge().setValue('4. REKAPITULASI UPAH PENJAHIT')
    .setBackground(IBS_DASH.COLOR.dark).setFontColor('#ffffff').setFontWeight('bold');
  curRow++;
  shRep.getRange(curRow, 1, 1, 6).setValues([['Tim Penjahit', 'Pcs Lolos QC', 'Pcs Reject', 'Belum Dibayar (Rp)', 'Sudah Dibayar (Rp)', 'Total Upah (Rp)']])
    .setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');
  curRow++;
  var upahArr = Object.keys(upahByTim).map(function (k) { return upahByTim[k]; }).sort(function (a, b) { return b.total - a.total; });
  if (upahArr.length) {
    var uRows = upahArr.map(function (u) { return [u.tim, u.lolos, u.gagal, u.unpaid, u.paid, u.total]; });
    shRep.getRange(curRow, 1, uRows.length, 6).setValues(uRows).setBackground('#ffffff');
    shRep.getRange(curRow, 2, uRows.length, 2).setNumberFormat('#,##0').setHorizontalAlignment('center');
    shRep.getRange(curRow, 4, uRows.length, 3).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    curRow += uRows.length + 1;
  } else {
    shRep.getRange(curRow, 1, 1, 6).merge().setValue('Belum ada data upah QC pada periode ini.');
    curRow += 2;
  }

  // TABEL 5: LAPORAN HPP PRODUKSI PER BATCH
  shRep.getRange(curRow, 1, 1, 8).merge().setValue('5. RINCIAN HPP PRODUKSI PER WORK ORDER')
    .setBackground(IBS_DASH.COLOR.dark).setFontColor('#ffffff').setFontWeight('bold');
  curRow++;
  shRep.getRange(curRow, 1, 1, 8).setValues([['Tanggal Tutup', 'HPP ID', 'Nomor WO', 'Tim Penjahit', 'Cost Bahan (Rp)', 'Cost Jahit (Rp)', 'Qty Lolos', 'HPP / Pcs (Rp)']])
    .setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');
  curRow++;
  if (hppList.length) {
    var hRows = hppList.map(function (h) { return [h.tanggal, h.hpp_id, h.nomor_wo, h.tim, h.bahan, h.jahit, h.qty, h.hpp_pcs]; });
    shRep.getRange(curRow, 1, hRows.length, 8).setValues(hRows).setBackground('#ffffff');
    shRep.getRange(curRow, 1, hRows.length, 1).setNumberFormat('dd MMM yyyy').setHorizontalAlignment('center');
    shRep.getRange(curRow, 5, hRows.length, 2).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right');
    shRep.getRange(curRow, 7, hRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    shRep.getRange(curRow, 8, hRows.length, 1).setNumberFormat('"Rp "#,##0').setHorizontalAlignment('right').setFontWeight('bold');
  } else {
    shRep.getRange(curRow, 1, 1, 8).merge().setValue('Belum ada batch HPP ditutup pada periode ini.');
  }

  shRep.setRowHeight(1, 36);
  shRep.setRowHeight(2, 22);
  shRep.setRowHeight(5, 34);
  for (var c = 1; c <= 10; c++) shRep.autoResizeColumn(c);

  ss.setActiveSheet(shRep);
  SpreadsheetApp.flush();
}