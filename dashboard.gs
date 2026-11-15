/**
 * INA BASIC STORE - Dashboard V5.1 (HPP, Profit & Rekap Upah Jahit Manual)
 */
var IBS_DASH = {
  DASH: 'DASHBOARD',
  DATA: '_DASH_DATA',
  REKAP: 'LAPORAN_REKAP',
  UPAH: 'REKAP_UPAH_JAHIT',
  TZ: 'Asia/Jakarta',
  COLOR: {
    page: '#f1f5f9', dark: '#0f172a', text: '#1e293b', muted: '#64748b',
    line: '#cbd5e1', indigo: '#4f46e5', green: '#059669', red: '#e11d48',
    amber: '#d97706', sky: '#0284c7', teal: '#0d9488', purple: '#7c3aed'
  }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ina Warehouse')
    .addItem('Siapkan / Perbarui Dashboard V5.1', 'setupDashboard')
    .addItem('Siapkan Sheet Sistem V5.1', 'setup')
    .addSeparator()
    .addItem('Tarik Rekap Upah Jahit (Mingguan)', 'tarikRekapUpahMingguan')
    .addItem('Tarik Rekap Upah Jahit (Semua Periode)', 'tarikRekapUpahSemua')
    .addSeparator()
    .addItem('Tarik Laporan Harian (30 Hari Terakhir)', 'tarikLaporanHarian')
    .addItem('Tarik Laporan Mingguan (12 Minggu Terakhir)', 'tarikLaporanMingguan')
    .addItem('Tarik Laporan Bulanan (12 Bulan Terakhir)', 'tarikLaporanBulanan')
    .addToUi();
}

function setupDashboard() {
  if (typeof setup === 'function') setup();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(IBS_DASH.TZ);
  if (!ss.getSheetByName(IBS_DASH.DASH)) ss.insertSheet(IBS_DASH.DASH);
  if (!ss.getSheetByName(IBS_DASH.DATA)) ss.insertSheet(IBS_DASH.DATA);

  buildDashDataV5_(ss);
  buildDashboardV5_(ss);

  var dash = ss.getSheetByName(IBS_DASH.DASH);
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
  SpreadsheetApp.flush();
}

function buildDashDataV5_(ss) {
  var sh = ss.getSheetByName(IBS_DASH.DATA) || ss.insertSheet(IBS_DASH.DATA);
  sh.clear();

  sh.getRange('A1:I1').setValues([['Tanggal', 'Masuk', 'Keluar', '', 'TGL_AWAL', 'TGL_AKHIR', '', 'LOKASI', '']]);
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

  sh.getRange('E2').setFormula('=IFS(DASHBOARD!$D$4="Hari Ini",TODAY(),DASHBOARD!$D$4="Kemarin",TODAY()-1,DASHBOARD!$D$4="7 Hari Terakhir",TODAY()-7,DASHBOARD!$D$4="30 Hari Terakhir",TODAY()-30,DASHBOARD!$D$4="Bulan Ini",EOMONTH(TODAY(),-1)+1,DASHBOARD!$D$4="Bulan Lalu",EOMONTH(TODAY(),-2)+1,DASHBOARD!$D$4="Semua Waktu",DATE(2000,1,1),TRUE,TODAY()-30)');
  sh.getRange('F2').setFormula('=IFS(DASHBOARD!$D$4="Kemarin",TODAY(),DASHBOARD!$D$4="Bulan Lalu",EOMONTH(TODAY(),-1)+1,TRUE,TODAY()+1)');
  sh.getRange('H2').setValue('Semua Lokasi');
  sh.getRange('H3').setFormula('=IFERROR(SORT(UNIQUE(FILTER(MASTER_LOKASI!B2:B,MASTER_LOKASI!D2:D=TRUE))), IFERROR(SORT(UNIQUE(FILTER(STOK_VARIASI!C2:C,STOK_VARIASI!C2:C<>""))),""))');

  sh.getRange('K1:V1').setValues([[
    'Timestamp', 'Aktivitas', 'Item_Key', 'Produk_Tampilan', 'Qty',
    'Lokasi_Asal', 'Lokasi_Tujuan', 'HPP_Satuan', 'Harga_Jual_Satuan',
    'Total_Omzet', 'Total_Profit', 'Total_HPP_Terjual'
  ]]);
  sh.getRange('K2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!A2:A))');
  sh.getRange('L2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!C2:C))');
  sh.getRange('M2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",IF(OPERATION_LOG!I2:I<>"",OPERATION_LOG!I2:I,OPERATION_LOG!D2:D)))');
  sh.getRange('N2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",IF(OPERATION_LOG!M2:M<>"",OPERATION_LOG!M2:M,IFERROR(VLOOKUP(OPERATION_LOG!D2:D,MASTER_PRODUK!A:B,2,FALSE),OPERATION_LOG!D2:D))))');
  sh.getRange('O2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!G2:G))');
  sh.getRange('P2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!E2:E))');
  sh.getRange('Q2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!F2:F))');
  sh.getRange('R2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",N(OPERATION_LOG!O2:O)))');
  sh.getRange('S2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",N(OPERATION_LOG!P2:P)))');
  sh.getRange('T2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",N(OPERATION_LOG!Q2:Q)))');
  sh.getRange('U2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",N(OPERATION_LOG!R2:R)))');
  sh.getRange('V2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",IF(OPERATION_LOG!C2:C="Keluar", N(OPERATION_LOG!G2:G)*N(OPERATION_LOG!O2:O), 0)))');

  sh.getRange('X1:AE1').setValues([[
    'ID_Variasi', 'Produk_Variasi', 'Lokasi', 'Qty',
    'HPP_Satuan', 'Harga_Jual', 'Nilai_Aset_HPP', 'Potensi_Omzet'
  ]]);
  sh.getRange('X2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",STOK_VARIASI!A2:A))');
  sh.getRange('Y2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",IFERROR(VLOOKUP(STOK_VARIASI!B2:B,MASTER_PRODUK!A:B,2,FALSE),"")&" · "&IFERROR(VLOOKUP(STOK_VARIASI!A2:A,VARIASI_PRODUK!A:C,3,FALSE),STOK_VARIASI!A2:A)))');
  sh.getRange('Z2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",STOK_VARIASI!C2:C))');
  sh.getRange('AA2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",N(STOK_VARIASI!D2:D)))');
  sh.getRange('AB2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",IFERROR(N(VLOOKUP(STOK_VARIASI!A2:A,VARIASI_PRODUK!A:H,8,FALSE)),0)))');
  sh.getRange('AC2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",IFERROR(N(VLOOKUP(STOK_VARIASI!A2:A,VARIASI_PRODUK!A:I,9,FALSE)),0)))');
  sh.getRange('AD2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",N(STOK_VARIASI!D2:D)*IFERROR(N(VLOOKUP(STOK_VARIASI!A2:A,VARIASI_PRODUK!A:H,8,FALSE)),0)))');
  sh.getRange('AE2').setFormula('=ARRAYFORMULA(IF(STOK_VARIASI!A2:A="","",N(STOK_VARIASI!D2:D)*IFERROR(N(VLOOKUP(STOK_VARIASI!A2:A,VARIASI_PRODUK!A:I,9,FALSE)),0)))');

  sh.hideSheet();
}

function topItemQueryV5_(activity) {
  var locClause = activity === 'Keluar' ? 'P' : 'Q';
  if (activity === 'Keluar') {
    return '=IFERROR(QUERY(_DASH_DATA!$K$1:$V,"select N, sum(O), sum(T), sum(U) where L=\'Keluar\' and K >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and K < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi",""," and ' + locClause + '=\'"&$G$4&"\' ")&" group by N order by sum(O) desc limit 10 label N \'Produk / Variasi\', sum(O) \'Qty\', sum(T) \'Omzet\', sum(U) \'Profit\'",1),{"Produk / Variasi","Qty","Omzet","Profit"})';
  }
  return '=IFERROR(QUERY(_DASH_DATA!$K$1:$V,"select N, count(K), sum(O) where L=\'' + activity + '\' and K >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and K < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi",""," and ' + locClause + '=\'"&$G$4&"\' ")&" group by N order by sum(O) desc limit 10 label N \'Produk / Variasi\', count(K) \'Transaksi\', sum(O) \'Total Qty\'",1),{"Produk / Variasi","Transaksi","Total Qty"})';
}

function sparkFormulaV5_(col, row, first, last, color) {
  return '=IF(' + col + row + '="","",IFERROR(SPARKLINE(' + col + row + ',{"charttype","bar";"max",MAX($' + col + '$' + first + ':$' + col + '$' + last + ');"color1","' + color + '"}),""))';
}

function buildDashboardV5_(ss) {
  var sh = ss.getSheetByName(IBS_DASH.DASH) || ss.insertSheet(IBS_DASH.DASH);
  sh.getCharts().forEach(function (chart) { sh.removeChart(chart); });
  sh.clear(); sh.clearConditionalFormatRules(); sh.getDataRange().clearDataValidations();
  var C = IBS_DASH.COLOR;
  var R = function (a1) { return sh.getRange(a1); };
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 14);
  for (var c = 2; c <= 14; c++) sh.setColumnWidth(c, 102);
  sh.setColumnWidth(15, 14);

  sh.setRowHeights(1, 72, 22);
  [1, 5, 8, 11, 24, 33, 41, 54].forEach(function (r) { sh.setRowHeight(r, 12); });
  sh.setRowHeight(2, 46);
  sh.setRowHeight(3, 24);
  sh.setRowHeight(7, 48);
  sh.setRowHeight(10, 48);

  R('A1:O72').setBackground(C.page).setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');
  R('B2:N2').merge().setValue('  INA BASIC STORE  ·  WAREHOUSE & PROFIT DASHBOARD V5.1').setBackground(C.dark).setFontColor('#fff').setFontSize(16).setFontWeight('bold');
  R('B3:N3').merge().setFormula('="  Diperbarui: "&TEXT(NOW(),"dd MMM yyyy HH:mm")&"  |  Gunakan menu Ina Warehouse > Tarik Rekap Upah Jahit atau Tarik Laporan"').setBackground('#1e293b').setFontColor('#cbd5e1');

  R('B4:C4').merge().setValue('Filter Waktu').setFontWeight('bold');
  R('D4').setValue('30 Hari Terakhir').setBackground('#fef9c3').setFontWeight('bold');
  R('E4:F4').merge().setValue('Filter Lokasi').setHorizontalAlignment('right').setFontWeight('bold');
  R('G4:H4').merge().setValue('Semua Lokasi').setBackground('#fef9c3').setFontWeight('bold');
  R('I4:J4').merge().setValue('Batas Kritis').setHorizontalAlignment('right').setFontWeight('bold');
  R('K4').setValue(10).setBackground('#fef9c3').setHorizontalAlignment('center').setFontWeight('bold');

  R('D4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Hari Ini', 'Kemarin', '7 Hari Terakhir', '30 Hari Terakhir', 'Bulan Ini', 'Bulan Lalu', 'Semua Waktu'], true).build());
  R('G4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(IBS_DASH.DATA).getRange('H2:H100'), true).build());

  function kpi(labelRange, valueRange, label, formula, color, numFormat) {
    R(labelRange).merge().setValue(label).setFontSize(8).setFontColor(C.muted).setHorizontalAlignment('center').setBackground('#ffffff');
    var vCell = R(valueRange).merge().setFormula(formula).setFontSize(17).setFontColor(color).setHorizontalAlignment('center').setFontWeight('bold').setBackground('#ffffff');
    if (numFormat) vCell.setNumberFormat(numFormat);
    R(labelRange.split(':')[0] + ':' + valueRange.split(':')[1]).setBorder(true, true, true, true, false, false, C.line, SpreadsheetApp.BorderStyle.SOLID);
  }
  function section(range, title) {
    R(range).merge().setValue('  ' + title).setBackground(C.dark).setFontColor('#fff').setFontWeight('bold');
  }

  kpi('B6:C6', 'B7:C7', 'PRODUK AKTIF', '=COUNTIF(MASTER_PRODUK!$C$2:$C,TRUE)', C.indigo, '#,##0');
  kpi('D6:E6', 'D7:E7', 'VARIASI AKTIF', '=COUNTIF(VARIASI_PRODUK!$E$2:$E,TRUE)', C.sky, '#,##0');
  kpi('F6:G6', 'F7:G7', 'TOTAL STOK (PCS)', '=IF($G$4="Semua Lokasi",SUM(STOK_VARIASI!$D$2:$D),SUMIFS(STOK_VARIASI!$D$2:$D,STOK_VARIASI!$C$2:$C,$G$4))', C.text, '#,##0');
  kpi('I6:J6', 'I7:J7', 'BARANG MASUK', '=IF($G$4="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2,OPERATION_LOG!$F$2:$F,$G$4))', C.green, '#,##0');
  kpi('K6:L6', 'K7:L7', 'PENJUALAN / KELUAR', '=IF($G$4="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2,OPERATION_LOG!$E$2:$E,$G$4))', C.red, '#,##0');
  kpi('M6:N6', 'M7:N7', 'STOK KRITIS', '=IFERROR(ROWS(QUERY(QUERY(STOK_VARIASI!$A$1:$D,"select A, sum(D) where A is not null "&IF($G$4="Semua Lokasi","","and C=\'"&$G$4&"\' ")&"group by A",1),"select Col1 where Col2 <= "&$K$4,1))-1,0)', C.amber, '#,##0');

  kpi('B9:C9', 'B10:C10', 'NILAI ASET STOK (HPP)', '=IF($G$4="Semua Lokasi",SUM(_DASH_DATA!$AD$2:$AD),SUMIFS(_DASH_DATA!$AD$2:$AD,_DASH_DATA!$Z$2:$Z,$G$4))', C.purple, '"Rp "#,##0');
  kpi('D9:E9', 'D10:E10', 'POTENSI OMZET STOK', '=IF($G$4="Semua Lokasi",SUM(_DASH_DATA!$AE$2:$AE),SUMIFS(_DASH_DATA!$AE$2:$AE,_DASH_DATA!$Z$2:$Z,$G$4))', C.teal, '"Rp "#,##0');
  kpi('F9:G9', 'F10:G10', 'OMZET PENJUALAN', '=IF($G$4="Semua Lokasi",SUMIFS(_DASH_DATA!$T$2:$T,_DASH_DATA!$L$2:$L,"Keluar",_DASH_DATA!$K$2:$K,">="&_DASH_DATA!$E$2,_DASH_DATA!$K$2:$K,"<"&_DASH_DATA!$F$2),SUMIFS(_DASH_DATA!$T$2:$T,_DASH_DATA!$L$2:$L,"Keluar",_DASH_DATA!$K$2:$K,">="&_DASH_DATA!$E$2,_DASH_DATA!$K$2:$K,"<"&_DASH_DATA!$F$2,_DASH_DATA!$P$2:$P,$G$4))', C.sky, '"Rp "#,##0');
  kpi('I9:J9', 'I10:J10', 'TOTAL HPP TERJUAL', '=IF($G$4="Semua Lokasi",SUMIFS(_DASH_DATA!$V$2:$V,_DASH_DATA!$L$2:$L,"Keluar",_DASH_DATA!$K$2:$K,">="&_DASH_DATA!$E$2,_DASH_DATA!$K$2:$K,"<"&_DASH_DATA!$F$2),SUMIFS(_DASH_DATA!$V$2:$V,_DASH_DATA!$L$2:$L,"Keluar",_DASH_DATA!$K$2:$K,">="&_DASH_DATA!$E$2,_DASH_DATA!$K$2:$K,"<"&_DASH_DATA!$F$2,_DASH_DATA!$P$2:$P,$G$4))', C.amber, '"Rp "#,##0');
  kpi('K9:L9', 'K10:L10', 'PROFIT BERSIH (RP)', '=IF($G$4="Semua Lokasi",SUMIFS(_DASH_DATA!$U$2:$U,_DASH_DATA!$L$2:$L,"Keluar",_DASH_DATA!$K$2:$K,">="&_DASH_DATA!$E$2,_DASH_DATA!$K$2:$K,"<"&_DASH_DATA!$F$2),SUMIFS(_DASH_DATA!$U$2:$U,_DASH_DATA!$L$2:$L,"Keluar",_DASH_DATA!$K$2:$K,">="&_DASH_DATA!$E$2,_DASH_DATA!$K$2:$K,"<"&_DASH_DATA!$F$2,_DASH_DATA!$P$2:$P,$G$4))', C.green, '"Rp "#,##0');
  kpi('M9:N9', 'M10:N10', 'RATA-RATA MARGIN', '=IFERROR(K10/F10,0)', C.indigo, '0.0%');

  section('B12:G12', 'TOP 10 PENJUALAN / VARIASI KELUAR');
  R('B13:C13').merge().setValue('Produk / Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('D13').setFormula(topItemQueryV5_('Keluar')).setBackground('#e2e8f0').setFontWeight('bold');
  for (var r = 14; r <= 23; r++) {
    R('B' + r + ':C' + r).merge().setFormula('=IF(D' + r + '="","",D' + r + ')').setBackground('#fff');
  }
  R('F14:G23').setNumberFormat('#,##0');

  section('I12:N12', 'STOK KRITIS PER VARIASI');
  R('I13:J13').merge().setValue('Produk / Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('K13').setValue('ID Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('L13').setValue('Stok').setBackground('#e2e8f0').setFontWeight('bold');
  R('M13:N13').merge().setValue('Status').setBackground('#e2e8f0').setFontWeight('bold');
  R('K14').setFormula('=IFERROR(QUERY(QUERY(STOK_VARIASI!$A$1:$D,"select A, sum(D) where A is not null "&IF($G$4="Semua Lokasi","","and C=\'"&$G$4&"\' ")&"group by A label A \'\', sum(D) \'\'",1),"select Col1, Col2 where Col2 <= "&$K$4&" order by Col2 asc limit 10",0),{"",""})');
  for (r = 14; r <= 23; r++) {
    R('I' + r + ':J' + r).merge().setFormula('=IF(K' + r + '="","",IFERROR(VLOOKUP(VLOOKUP(K' + r + ',VARIASI_PRODUK!$A:$C,2,FALSE),MASTER_PRODUK!$A:$B,2,FALSE),"")&" · "&IFERROR(VLOOKUP(K' + r + ',VARIASI_PRODUK!$A:$C,3,FALSE),K' + r + '))').setBackground('#fff');
    R('M' + r + ':N' + r).merge().setFormula('=IF(L' + r + '="","",IF(L' + r + '<=0,"HABIS","MENIPIS"))').setBackground('#fff').setHorizontalAlignment('center');
  }

  section('B25:G25', 'TREN BARANG MASUK VS KELUAR · 14 HARI');
  R('B26:G40').setBackground('#fff');

  section('I25:N25', 'PERSEBARAN STOK & NILAI ASET HPP PER LOKASI');
  R('I26').setFormula('=IFERROR(QUERY(_DASH_DATA!$X$1:$AE,"select Z, sum(AA), sum(AD) where Z is not null "&IF($G$4="Semua Lokasi","","and Z=\'"&$G$4&"\' ")&"group by Z order by sum(AA) desc limit 6 label Z \'Lokasi\', sum(AA) \'Total Unit\', sum(AD) \'Aset HPP (Rp)\'",1),{"Lokasi","Total Unit","Aset HPP (Rp)"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('L26').setValue('Porsi').setBackground('#e2e8f0').setFontWeight('bold');
  R('M26:N26').merge().setValue('Visual').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 27; r <= 32; r++) {
    R('L' + r).setFormula('=IF(J' + r + '="","",IFERROR(J' + r + '/SUM($J$27:$J$32),0))').setNumberFormat('0%').setBackground('#fff');
    R('M' + r + ':N' + r).merge().setFormula(sparkFormulaV5_('J', r, 27, 32, C.indigo)).setBackground('#fff');
  }
  R('K27:K32').setNumberFormat('#,##0');

  section('I34:N34', 'AKTIVITAS TRANSAKSI');
  R('I35').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$R,"select C, count(B) where C is not null and A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi","","and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') ")&"group by C order by count(B) desc limit 5 label C \'Aktivitas\', count(B) \'Transaksi\'",1),{"Aktivitas","Transaksi"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('K35:N35').merge().setValue('Visual').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 36; r <= 40; r++) {
    R('K' + r + ':N' + r).merge().setFormula(sparkFormulaV5_('J', r, 36, 40, C.sky)).setBackground('#fff');
  }

  section('B42:G42', '10 TRANSAKSI TERBARU');
  R('B43').setFormula('=IFERROR(QUERY(_DASH_DATA!$K$1:$V,"select K, L, N, O, T, U where K is not null "&IF($G$4="Semua Lokasi","","and (P=\'"&$G$4&"\' or Q=\'"&$G$4&"\') ")&"order by K desc limit 10 label K \'Waktu\', L \'Aktivitas\', N \'Produk / Variasi\', O \'Qty\', T \'Omzet\', U \'Profit\'",1),{"Waktu","Aktivitas","Produk / Variasi","Qty","Omzet","Profit"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('B44:B53').setNumberFormat('dd MMM HH:mm').setBackground('#fff');
  R('F44:G53').setNumberFormat('#,##0');

  section('I42:N42', 'TOP 10 MARGIN KEUNTUNGAN VARIASI');
  R('I43:J43').merge().setValue('Produk / Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('K43').setValue('HPP').setBackground('#e2e8f0').setFontWeight('bold');
  R('L43').setValue('Harga Jual').setBackground('#e2e8f0').setFontWeight('bold');
  R('M43').setValue('Profit/Pcs').setBackground('#e2e8f0').setFontWeight('bold');
  R('N43').setValue('Margin').setBackground('#e2e8f0').setFontWeight('bold');
  R('I44').setFormula('=IFERROR(QUERY(VARIASI_PRODUK!$A$1:$K,"select C, H, I, J, K where E = true and I > 0 order by J desc limit 10 label C \'\', H \'\', I \'\', J \'\', K \'\'",1),{"","","","",""})');
  for (r = 44; r <= 53; r++) {
    R('I' + r + ':J' + r).merge().setBackground('#fff');
  }
  R('K44:M53').setNumberFormat('#,##0');
  R('N44:N53').setNumberFormat('0.0%');

  sh.getRangeList(['E14:G23', 'K14:L23', 'I27:K32', 'I36:J40', 'B44:G53', 'K44:N53']).getRanges().forEach(function (rg) {
    rg.setBackground('#fff').setHorizontalAlignment('center');
  });

  var data = ss.getSheetByName(IBS_DASH.DATA);
  var chart = sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(data.getRange('A1:C15')).setNumHeaders(1)
    .setPosition(26, 2, 2, 2).setOption('title', 'Jumlah unit Masuk vs Keluar (14 Hari)').setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#10b981', '#f43f5e']).setOption('backgroundColor', '#ffffff')
    .setOption('vAxis', { minValue: 0, viewWindow: { min: 0 }, format: '0' }).setOption('hAxis', { slantedText: false }).setOption('width', 596).setOption('height', 316).build();
  sh.insertChart(chart);

  var habis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('HABIS').setBackground('#fee2e2').setFontColor('#b91c1c').setRanges([R('M14:M23')]).build();
  var menipis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('MENIPIS').setBackground('#fef3c7').setFontColor('#b45309').setRanges([R('M14:M23')]).build();
  sh.setConditionalFormatRules([habis, menipis]);
}

function tarikRekapUpahMingguan() { generateRekapUpah_('MINGGUAN'); }
function tarikRekapUpahSemua() { generateRekapUpah_('SEMUA'); }

function generateRekapUpah_(mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lapSh = ss.getSheetByName('LAPORAN_JAHIT');
  if (!lapSh || lapSh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Belum ada data setoran di sheet LAPORAN_JAHIT.');
    return;
  }

  var existingSh = ss.getSheetByName(IBS_DASH.UPAH);
  var savedManual = {};
  if (existingSh && existingSh.getLastRow() >= 4) {
    var oldVals = existingSh.getRange(4, 1, existingSh.getLastRow() - 3, 10).getValues();
    oldVals.forEach(function (r) {
      var noWo = String(r[0] || '').trim();
      var penjahit = String(r[3] || '').trim();
      if (noWo && penjahit && noWo !== 'TOTAL KESELURUHAN') {
        var key = noWo.toUpperCase() + '||' + penjahit.toLowerCase();
        savedManual[key] = {
          tarif: Number(r[7]) || 0,
          status: String(r[9] || 'BELUM DIBAYAR').trim() || 'BELUM DIBAYAR'
        };
      }
    });
  }

  var woInfoMap = {};
  var woSh = ss.getSheetByName('WO_JAHIT');
  var itemSh = ss.getSheetByName('WO_JAHIT_ITEM');
  var itemRows = (itemSh && itemSh.getLastRow() >= 2) ? itemSh.getRange(2, 1, itemSh.getLastRow() - 1, 5).getValues() : [];
  if (woSh && woSh.getLastRow() >= 2) {
    var woRows = woSh.getRange(2, 1, woSh.getLastRow() - 1, 10).getValues();
    woRows.forEach(function (w) {
      var woId = String(w[0] || '').trim();
      var noWo = String(w[1] || '').trim();
      var matchingItems = itemRows.filter(function (it) { return String(it[1] || '').trim() === woId; });
      var targetTotal = matchingItems.reduce(function (s, it) { return s + (Number(it[3]) || 0); }, 0);
      var desc = matchingItems.map(function (it) { return String(it[2] || '').trim(); }).filter(Boolean).join(', ') || String(w[3] || '').trim();
      var info = { target: targetTotal, desc: desc, status: String(w[2] || '').trim() };
      if (woId) woInfoMap[woId] = info;
      if (noWo) woInfoMap[noWo.toUpperCase()] = info;
    });
  }

  var now = new Date();
  var cutoff = new Date();
  if (mode === 'MINGGUAN') {
    cutoff.setDate(now.getDate() - 7);
  } else {
    cutoff = new Date(2000, 0, 1);
  }

  var lapCols = Math.max(lapSh.getLastColumn(), 18);
  var lapRows = lapSh.getRange(2, 1, lapSh.getLastRow() - 1, lapCols).getValues();

  var aggMap = {};
  var tailorSet = {};

  lapRows.forEach(function (r) {
    var status = String(r[7] || '').trim().toUpperCase();
    if (status !== 'APPROVED') return;

    var waktuQc = r[17] instanceof Date ? r[17] : (r[8] instanceof Date ? r[8] : new Date(r[17] || r[8]));
    if (mode === 'MINGGUAN' && !isNaN(waktuQc.getTime()) && waktuQc < cutoff) return;

    var woId = String(r[2] || '').trim();
    var noWo = String(r[3] || '').trim() || woId || 'TANPA-WO';
    var penjahit = String(r[4] || '').trim() || 'Tanpa Nama';
    var qtyBagus = Number(r[5]) || 0;
    var qtyLolos = Number(r[10]) || 0;
    var qtyGagal = Number(r[11]) || 0;

    var key = noWo.toUpperCase() + '||' + penjahit.toLowerCase();
    if (!aggMap[key]) {
      var info = woInfoMap[woId] || woInfoMap[noWo.toUpperCase()] || { target: 0, desc: '-' };
      aggMap[key] = {
        noWo: noWo,
        desc: info.desc,
        target: info.target,
        penjahit: penjahit,
        bagus: 0,
        lolos: 0,
        gagal: 0
      };
    }
    aggMap[key].bagus += qtyBagus;
    aggMap[key].lolos += qtyLolos;
    aggMap[key].gagal += qtyGagal;
    tailorSet[penjahit] = true;
  });

  var keys = Object.keys(aggMap).sort(function (a, b) {
    if (aggMap[a].noWo !== aggMap[b].noWo) return aggMap[b].noWo.localeCompare(aggMap[a].noWo);
    return aggMap[a].penjahit.localeCompare(aggMap[b].penjahit);
  });

  var outSh = existingSh || ss.insertSheet(IBS_DASH.UPAH);
  outSh.clear();
  outSh.clearConditionalFormatRules();
  outSh.getDataRange().clearDataValidations();

  var labelPeriode = mode === 'MINGGUAN' ? '7 HARI TERAKHIR' : 'SEMUA PERIODE';
  outSh.getRange('A1:J1').merge()
    .setValue('REKAPITULASI AKUMULASI HASIL JAHIT & UPAH MANUAL (' + labelPeriode + ') - INA BASIC STORE')
    .setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
  outSh.getRange('A2:J2').merge()
    .setValue('Petunjuk: Kolom H berwarna kuning diisi manual oleh Owner (Rp/pcs). Kolom I otomatis menghitung Lolos QC x Tarif Upah.')
    .setBackground('#1e293b').setFontColor('#cbd5e1').setFontSize(9);

  outSh.getRange('A3:J3').setValues([[
    'Nomor WO', 'Model / Item WO', 'Target WO (Pcs)', 'Nama Penjahit',
    'Setor Bagus (Pcs)', 'Lolos QC (Pcs)', 'Gagal QC (Pcs)',
    'Tarif Upah / Pcs (Input Manual)', 'Total Upah (Rp)', 'Status Bayar'
  ]]).setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');

  if (keys.length === 0) {
    outSh.getRange('A4:J4').merge()
      .setValue('Belum ada laporan jahit berstatus APPROVED (Lolos QC) pada periode ini.')
      .setHorizontalAlignment('center').setFontColor('#64748b');
    ss.setActiveSheet(outSh);
    SpreadsheetApp.flush();
    return;
  }

  var rowsData = [];
  keys.forEach(function (k, idx) {
    var item = aggMap[k];
    var rNum = 4 + idx;
    var prev = savedManual[k] || { tarif: '', status: 'BELUM DIBAYAR' };
    rowsData.push([
      item.noWo,
      item.desc,
      item.target,
      item.penjahit,
      item.bagus,
      item.lolos,
      item.gagal,
      prev.tarif > 0 ? prev.tarif : '',
      '=IF(H' + rNum + '>0, F' + rNum + '*H' + rNum + ', 0)',
      prev.status
    ]);
  });

  var lastDataRow = 3 + rowsData.length;
  outSh.getRange(4, 1, rowsData.length, 10).setValues(rowsData);

  outSh.getRange(4, 3, rowsData.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
  outSh.getRange(4, 5, rowsData.length, 3).setNumberFormat('#,##0').setHorizontalAlignment('center');
  outSh.getRange(4, 8, rowsData.length, 1).setBackground('#fef9c3').setNumberFormat('"Rp "#,##0').setFontWeight('bold');
  outSh.getRange(4, 9, rowsData.length, 1).setNumberFormat('"Rp "#,##0').setFontWeight('bold');
  outSh.getRange(4, 10, rowsData.length, 1).setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['BELUM DIBAYAR', 'SUDAH DIBAYAR'], true).build());

  var totRow = lastDataRow + 1;
  outSh.getRange(totRow, 1, 1, 4).merge().setValue('TOTAL KESELURUHAN').setFontWeight('bold').setHorizontalAlignment('right');
  outSh.getRange(totRow, 5).setFormula('=SUM(E4:E' + lastDataRow + ')').setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('center');
  outSh.getRange(totRow, 6).setFormula('=SUM(F4:F' + lastDataRow + ')').setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('center');
  outSh.getRange(totRow, 7).setFormula('=SUM(G4:G' + lastDataRow + ')').setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('center');
  outSh.getRange(totRow, 8).setValue('-').setHorizontalAlignment('center');
  outSh.getRange(totRow, 9).setFormula('=SUM(I4:I' + lastDataRow + ')').setNumberFormat('"Rp "#,##0').setFontWeight('bold');
  outSh.getRange(totRow, 1, 1, 10).setBackground('#dbeafe');

  var sumStart = totRow + 3;
  outSh.getRange(sumStart - 1, 1, 1, 5).merge()
    .setValue('RINGKASAN TOTAL GAJI PER PENJAHIT (OTOMATIS DARI TABEL DI ATAS)')
    .setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
  outSh.getRange(sumStart, 1, 1, 5).setValues([[
    'Nama Penjahit', 'Total Lolos QC (Pcs)', 'Total Upah (Rp)', 'Sudah Dibayar (Rp)', 'Sisa Belum Dibayar (Rp)'
  ]]).setBackground('#e2e8f0').setFontWeight('bold').setHorizontalAlignment('center');

  var tailors = Object.keys(tailorSet).sort();
  var summaryRows = tailors.map(function (name, i) {
    var sr = sumStart + 1 + i;
    return [
      name,
      '=SUMIF($D$4:$D$' + lastDataRow + ', A' + sr + ', $F$4:$F$' + lastDataRow + ')',
      '=SUMIF($D$4:$D$' + lastDataRow + ', A' + sr + ', $I$4:$I$' + lastDataRow + ')',
      '=SUMIFS($I$4:$I$' + lastDataRow + ', $D$4:$D$' + lastDataRow + ', A' + sr + ', $J$4:$J$' + lastDataRow + ', "SUDAH DIBAYAR")',
      '=C' + sr + '-D' + sr
    ];
  });

  if (summaryRows.length > 0) {
    outSh.getRange(sumStart + 1, 1, summaryRows.length, 5).setValues(summaryRows);
    outSh.getRange(sumStart + 1, 2, summaryRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    outSh.getRange(sumStart + 1, 3, summaryRows.length, 3).setNumberFormat('"Rp "#,##0').setFontWeight('bold');
  }

  var paidRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SUDAH DIBAYAR').setBackground('#dcfce7').setFontColor('#166534').setRanges([outSh.getRange(4, 10, rowsData.length, 1)]).build();
  var unpaidRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('BELUM DIBAYAR').setBackground('#fee2e2').setFontColor('#991b1b').setRanges([outSh.getRange(4, 10, rowsData.length, 1)]).build();
  outSh.setConditionalFormatRules([paidRule, unpaidRule]);

  for (var col = 1; col <= 10; col++) outSh.autoResizeColumn(col);
  ss.setActiveSheet(outSh);
  SpreadsheetApp.flush();
}

function tarikLaporanHarian() { generatePeriodicReport_('HARIAN'); }
function tarikLaporanMingguan() { generatePeriodicReport_('MINGGUAN'); }
function tarikLaporanBulanan() { generatePeriodicReport_('BULANAN'); }

function generatePeriodicReport_(mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSh = ss.getSheetByName('OPERATION_LOG');
  if (!logSh || logSh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Belum ada data transaksi di OPERATION_LOG.');
    return;
  }

  var values = logSh.getRange(2, 1, logSh.getLastRow() - 1, Math.max(logSh.getLastColumn(), 18)).getValues();
  var tz = IBS_DASH.TZ;
  var now = new Date();
  var cutoff = new Date();
  if (mode === 'HARIAN') cutoff.setDate(now.getDate() - 30);
  else if (mode === 'MINGGUAN') cutoff.setDate(now.getDate() - 84);
  else cutoff.setMonth(now.getMonth() - 12);

  var buckets = {};
  values.forEach(function (row) {
    var ts = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (isNaN(ts.getTime()) || ts < cutoff) return;

    var act = String(row[2] || '').trim();
    var qty = Number(row[6]) || 0;
    var hpp = Number(row[14]) || 0;
    var omzet = Number(row[16]) || 0;
    var profit = Number(row[17]) || 0;

    var key = '';
    if (mode === 'HARIAN') {
      key = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    } else if (mode === 'MINGGUAN') {
      var d = new Date(ts);
      var day = d.getDay();
      var diff = d.getDate() - day + (day === 0 ? -6 : 1);
      d.setDate(diff);
      key = 'Minggu ' + Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    } else {
      key = Utilities.formatDate(ts, tz, 'yyyy-MM');
    }

    if (!buckets[key]) {
      buckets[key] = { periode: key, masuk: 0, keluar: 0, totalHpp: 0, omzet: 0, profit: 0 };
    }
    if (act === 'Masuk') buckets[key].masuk += qty;
    if (act === 'Keluar') {
      buckets[key].keluar += qty;
      buckets[key].totalHpp += (qty * hpp);
      buckets[key].omzet += omzet;
      buckets[key].profit += profit;
    }
  });

  var sortedKeys = Object.keys(buckets).sort().reverse();
  var outRows = sortedKeys.map(function (k) {
    var b = buckets[k];
    var margin = b.omzet > 0 ? (b.profit / b.omzet) : 0;
    return [b.periode, b.masuk, b.keluar, b.totalHpp, b.omzet, b.profit, margin];
  });

  var repSh = ss.getSheetByName(IBS_DASH.REKAP) || ss.insertSheet(IBS_DASH.REKAP);
  repSh.clear();
  repSh.getRange('A1:G1').merge().setValue('REKAPITULASI PENJUALAN & PROFIT (' + mode + ') - INA BASIC STORE')
    .setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);
  repSh.getRange('A2:G2').setValues([['Periode', 'Barang Masuk (Pcs)', 'Terjual / Keluar (Pcs)', 'Total HPP Terjual (Rp)', 'Total Omzet (Rp)', 'Profit Bersih (Rp)', 'Margin (%)']])
    .setBackground('#e2e8f0').setFontWeight('bold');

  if (outRows.length > 0) {
    repSh.getRange(3, 1, outRows.length, 7).setValues(outRows);
    repSh.getRange(3, 2, outRows.length, 2).setNumberFormat('#,##0');
    repSh.getRange(3, 4, outRows.length, 3).setNumberFormat('#,##0');
    repSh.getRange(3, 7, outRows.length, 1).setNumberFormat('0.0%');

    var totRow = 3 + outRows.length;
    repSh.getRange(totRow, 1).setValue('TOTAL').setFontWeight('bold');
    repSh.getRange(totRow, 2).setFormula('=SUM(B3:B' + (totRow - 1) + ')').setNumberFormat('#,##0').setFontWeight('bold');
    repSh.getRange(totRow, 3).setFormula('=SUM(C3:C' + (totRow - 1) + ')').setNumberFormat('#,##0').setFontWeight('bold');
    repSh.getRange(totRow, 4).setFormula('=SUM(D3:D' + (totRow - 1) + ')').setNumberFormat('#,##0').setFontWeight('bold');
    repSh.getRange(totRow, 5).setFormula('=SUM(E3:E' + (totRow - 1) + ')').setNumberFormat('#,##0').setFontWeight('bold');
    repSh.getRange(totRow, 6).setFormula('=SUM(F3:F' + (totRow - 1) + ')').setNumberFormat('#,##0').setFontWeight('bold');
    repSh.getRange(totRow, 7).setFormula('=IFERROR(F' + totRow + '/E' + totRow + ',0)').setNumberFormat('0.0%').setFontWeight('bold');
    repSh.getRange(totRow, 1, 1, 7).setBackground('#fef9c3');
  }

  for (var col = 1; col <= 7; col++) repSh.autoResizeColumn(col);
  ss.setActiveSheet(repSh);
  SpreadsheetApp.flush();
}