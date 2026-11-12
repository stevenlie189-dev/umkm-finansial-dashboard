/**
 * ROTI RONA — Sistem Pembukuan & Dashboard Keuangan UMKM (Google Sheets)
 * ---------------------------------------------------------------------
 * Menjalankan setup() akan membuat spreadsheet baru berisi 4 sheet:
 *   DASHBOARD, TRANSAKSI, KATEGORI, PETUNJUK
 * lengkap dengan formula, dropdown, conditional formatting, grafik,
 * dan data transaksi SIMULASI (fiktif).
 *
 * Tidak ada API key / password / token di dalam script ini.
 * Script hanya meminta izin standar Google (Spreadsheet) saat pertama dijalankan.
 *
 * Fungsi utama:
 *   setup()                    -> buat spreadsheet baru (disarankan)
 *   setupDiSpreadsheetIni()    -> bangun di spreadsheet aktif (harus kosong)
 *   hapusDataSimulasi()        -> kosongkan transaksi demo, siap input data asli
 */

// ============================== KONFIGURASI ==============================
var CFG = {
  NAMA_FILE: 'ROTI RONA - Pembukuan & Dashboard Keuangan',
  AWAL: 5,             // baris pertama data transaksi
  BARIS: 5000,         // jumlah baris yang sudah disiapkan formula & validasi
  SALDO_AWAL: 15000000,
  TAHUN_DEFAULT: 2026,
  TAHUN_MULAI: 2024,   // daftar tahun filter: TAHUN_MULAI .. +11
  KATEGORI_SLOT: 60,   // kapasitas daftar kategori (baris 5..64)
  C: {
    navy: '#0B2545', navy2: '#13315C', mid: '#1F4E8C', soft: '#E8EEF7',
    bg: '#F3F5F8', white: '#FFFFFF', line: '#D9DEE5', text: '#1B2430', muted: '#5B6675',
    green: '#1E8449', greenBg: '#E6F4EC', red: '#C0392B', redBg: '#FBE9E7',
    amber: '#9A6700', amberBg: '#FFF4DB', input: '#FFFBEA'
  }
};

var KATEGORI_AWAL = [
  ['Penjualan Produk', 'Pendapatan', 'Roti, pastry, kue pesanan'],
  ['Pendapatan Lainnya', 'Pendapatan', 'Katering, barang bekas, lain-lain'],
  ['Pembelian Bahan Baku', 'Pengeluaran', 'Tepung, gula, mentega, telur, dll.'],
  ['Kemasan', 'Pengeluaran', 'Box, plastik, label'],
  ['Gaji', 'Pengeluaran', 'Gaji & THR karyawan'],
  ['Listrik dan Air', 'Pengeluaran', 'Token listrik, PDAM'],
  ['Transportasi', 'Pengeluaran', 'BBM, ongkos kirim, kurir'],
  ['Sewa', 'Pengeluaran', 'Sewa tempat usaha'],
  ['Marketing', 'Pengeluaran', 'Iklan, endorse, promosi'],
  ['Operasional', 'Pengeluaran', 'Gas, kebersihan, internet, servis alat'],
  ['Pengeluaran Lainnya', 'Pengeluaran', 'Biaya tak terduga']
];
var METODE = ['Tunai', 'Transfer Bank', 'QRIS', 'E-Wallet', 'Lainnya'];
var BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
var BULAN_PENDEK = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

var FMT_RP = '"Rp "#,##0';
var FMT_RP_NEG = '"Rp "#,##0;[Red]-"Rp "#,##0';
var FMT_TGL = 'dd mmm yyyy';

// ============================== ENTRY POINT ==============================
/** Membuat spreadsheet BARU di Google Drive akun yang menjalankan script. */
function setup() {
  var ss = SpreadsheetApp.create(CFG.NAMA_FILE);
  bangun_(ss, true);
  Logger.log('Spreadsheet dibuat: ' + ss.getUrl());
  return ss.getUrl();
}

/** Membangun sistem di spreadsheet yang sedang aktif (harus belum punya sheet bernama sama). */
function setupDiSpreadsheetIni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Tidak ada spreadsheet aktif. Gunakan setup() untuk membuat file baru.');
  ['DASHBOARD', 'TRANSAKSI', 'KATEGORI', 'PETUNJUK'].forEach(function (n) {
    if (ss.getSheetByName(n)) throw new Error('Sheet "' + n + '" sudah ada. Setup dibatalkan agar data Anda tidak tertimpa.');
  });
  bangun_(ss, false);
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('ROTI RONA')
      .addItem('Hapus data simulasi (mulai input data asli)', 'hapusDataSimulasi')
      .addToUi();
  } catch (e) { /* konteks tanpa UI */ }
}

/** Mengosongkan input transaksi demo; formula & dropdown tetap ada. */
function hapusDataSimulasi() {
  var ui = SpreadsheetApp.getUi();
  var jawab = ui.alert('Hapus data simulasi?',
    'Semua baris transaksi (kolom Tanggal, Jenis, Kategori, Keterangan, Metode, Nominal) akan dikosongkan. ' +
    'Saldo Awal di DASHBOARD tidak diubah — sesuaikan sendiri. Lanjutkan?', ui.ButtonSet.YES_NO);
  if (jawab !== ui.Button.YES) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TRANSAKSI');
  sh.getRange(CFG.AWAL, 1, CFG.BARIS, 1).clearContent();
  sh.getRange(CFG.AWAL, 3, CFG.BARIS, 5).clearContent();
}

function bangun_(ss, ganti) {
  ss.setSpreadsheetTimeZone('Asia/Jakarta');
  var dash, trx, kat, pet;
  if (ganti) {
    dash = ss.getSheets()[0];
    dash.setName('DASHBOARD');
    trx = ss.insertSheet('TRANSAKSI');
    kat = ss.insertSheet('KATEGORI');
    pet = ss.insertSheet('PETUNJUK');
  } else {
    dash = ss.insertSheet('DASHBOARD');
    trx = ss.insertSheet('TRANSAKSI');
    kat = ss.insertSheet('KATEGORI');
    pet = ss.insertSheet('PETUNJUK');
  }
  bangunKategori_(kat);
  bangunTransaksi_(trx);
  bangunDashboard_(dash);
  bangunPetunjuk_(pet);
  // PENTING: locale diatur TERAKHIR. Jika diatur sebelum formula ditulis, locale id_ID
  // (pemisah argumen titik koma) membuat seluruh formula berpemisah koma menjadi #ERROR!.
  ss.setSpreadsheetLocale('id_ID');
  ss.setActiveSheet(dash);
}

// ============================== HELPER UI ==============================
/** Isi & format satu range. opt: merge, bg, color, bold, italic, size, ha, va, nf, wrap, border */
function put_(sh, a1, value, opt) {
  opt = opt || {};
  var r = sh.getRange(a1);
  if (opt.merge) r.merge();
  var first = sh.getRange(a1.split(':')[0]);
  if (value !== null && value !== undefined) {
    if (typeof value === 'string' && value.charAt(0) === '=') first.setFormula(value);
    else first.setValue(value);
  }
  if (opt.bg) r.setBackground(opt.bg);
  if (opt.color) r.setFontColor(opt.color);
  if (opt.bold) r.setFontWeight('bold');
  if (opt.italic) r.setFontStyle('italic');
  if (opt.size) r.setFontSize(opt.size);
  if (opt.ha) r.setHorizontalAlignment(opt.ha);
  r.setVerticalAlignment(opt.va || 'middle');
  if (opt.nf) r.setNumberFormat(opt.nf);
  if (opt.wrap) r.setWrap(true);
  if (opt.border) r.setBorder(true, true, true, true, null, null, CFG.C.line, SpreadsheetApp.BorderStyle.SOLID);
  return r;
}

function aturan_(sh, a1, formula, bg, fc, bold) {
  var b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setRanges([sh.getRange(a1)]);
  if (bg) b.setBackground(bg);
  if (fc) b.setFontColor(fc);
  if (bold) b.setBold(true);
  return b.build();
}

// ============================== SHEET: KATEGORI ==============================
function bangunKategori_(sh) {
  var C = CFG.C;
  sh.setTabColor(C.mid);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:M70').setFontFamily('Arial').setFontSize(10).setFontColor(C.text);
  sh.setColumnWidth(1, 45);
  sh.setColumnWidth(2, 210); sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 260); sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 24); sh.setColumnWidth(7, 150); sh.setColumnWidth(8, 24);
  sh.setColumnWidth(9, 70); sh.setColumnWidth(10, 24); sh.setColumnWidth(11, 120); sh.setColumnWidth(12, 90);

  put_(sh, 'A1:L1', 'KATEGORI & DAFTAR PILIHAN', { merge: true, bg: C.navy, color: C.white, bold: true, size: 14 });
  sh.setRowHeight(1, 34);
  put_(sh, 'A2:L2', 'Sumber dropdown untuk sheet TRANSAKSI & DASHBOARD. Tambah kategori baru di baris kosong tabel kiri (kolom Kategori + Jenis).',
    { merge: true, color: C.muted, italic: true });

  var head = { bg: C.navy2, color: C.white, bold: true, ha: 'center' };
  put_(sh, 'A4', 'No', head); put_(sh, 'B4', 'Kategori', head); put_(sh, 'C4', 'Jenis', head);
  put_(sh, 'D4', 'Contoh isi / catatan', head); put_(sh, 'E4', 'Urutan biaya (otomatis)', { bg: C.navy2, color: C.white, bold: true, ha: 'center', wrap: true });
  put_(sh, 'G4', 'Metode Pembayaran', head);
  put_(sh, 'I4', 'Tahun', head);
  put_(sh, 'K4', 'Bulan (filter)', head); put_(sh, 'L4', 'Singkatan', head);
  sh.setRowHeight(4, 34);

  var n = CFG.KATEGORI_SLOT, a = CFG.AWAL, z = a + n - 1;
  // Tabel kategori
  var nomor = [], urut = [], i;
  for (i = 0; i < n; i++) {
    var r = a + i;
    nomor.push(['=IF(B' + r + '="","",ROW()-' + (a - 1) + ')']);
    urut.push(['=IF(C' + r + '="Pengeluaran",COUNTIF(C$' + a + ':C' + r + ',"Pengeluaran"),"")']);
  }
  sh.getRange(a, 1, n, 1).setFormulas(nomor);
  sh.getRange(a, 5, n, 1).setFormulas(urut);
  sh.getRange(a, 2, KATEGORI_AWAL.length, 3).setValues(KATEGORI_AWAL);
  sh.getRange(a, 1, n, 5).setBorder(true, true, true, true, true, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(a, 2, n, 3).setBackground(C.white);
  sh.getRange(a, 1, n, 1).setBackground(C.bg).setHorizontalAlignment('center');
  sh.getRange(a, 5, n, 1).setBackground(C.bg).setHorizontalAlignment('center').setFontColor(C.muted);
  sh.getRange(a, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendapatan', 'Pengeluaran'], true).setAllowInvalid(false)
    .setHelpText('Pilih Pendapatan atau Pengeluaran').build());

  // Metode pembayaran
  sh.getRange(a, 7, 8, 1).setBorder(true, true, true, true, true, true, C.line, SpreadsheetApp.BorderStyle.SOLID).setBackground(C.white);
  sh.getRange(a, 7, METODE.length, 1).setValues(METODE.map(function (m) { return [m]; }));
  // Tahun
  var thn = [];
  for (i = 0; i < 12; i++) thn.push([CFG.TAHUN_MULAI + i]);
  sh.getRange(a, 9, 12, 1).setValues(thn).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, C.line, SpreadsheetApp.BorderStyle.SOLID).setBackground(C.white);
  // Bulan
  var bl = [['Semua Bulan', '']];
  for (i = 0; i < 12; i++) bl.push([BULAN[i], BULAN_PENDEK[i]]);
  sh.getRange(a, 11, 13, 2).setValues(bl)
    .setBorder(true, true, true, true, true, true, C.line, SpreadsheetApp.BorderStyle.SOLID).setBackground(C.white);

  // Warna jenis
  var rules = [
    aturan_(sh, 'C' + a + ':C' + z, '=$C' + a + '="Pendapatan"', C.greenBg, C.green, true),
    aturan_(sh, 'C' + a + ':C' + z, '=$C' + a + '="Pengeluaran"', C.redBg, C.red, true),
    aturan_(sh, 'B' + a + ':C' + z, '=AND($B' + a + '<>"",COUNTIF($B$' + a + ':$B$' + z + ',$B' + a + ')>1)', C.amberBg, C.amber, true)
  ];
  sh.setConditionalFormatRules(rules);
  sh.setFrozenRows(4);
}

// ============================== SHEET: TRANSAKSI ==============================
function bangunTransaksi_(sh) {
  var C = CFG.C, a = CFG.AWAL, n = CFG.BARIS, z = a + n - 1;
  sh.setTabColor(C.green);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:J' + z).setFontFamily('Arial').setFontSize(10).setFontColor(C.text).setVerticalAlignment('middle');
  var w = [105, 95, 115, 175, 300, 130, 130, 130, 110, 150];
  for (var i = 0; i < w.length; i++) sh.setColumnWidth(i + 1, w[i]);

  put_(sh, 'A1:J1', 'TRANSAKSI HARIAN — ROTI RONA', { merge: true, bg: C.navy, color: C.white, bold: true, size: 14 });
  sh.setRowHeight(1, 34);
  put_(sh, 'A2:J2', 'DATA SIMULASI (fiktif) untuk demonstrasi. Isi kolom berlatar putih: Tanggal, Jenis, Kategori, Keterangan, Metode, Nominal. ' +
    'Kolom abu-abu (ID, Arus Kas, Bulan, Cek Data) terisi otomatis — jangan ditimpa.',
    { merge: true, color: C.amber, bg: C.amberBg, wrap: true });
  sh.setRowHeight(2, 34);
  put_(sh, 'A3:J3', 'Baris siap pakai: ' + a + ' s.d. ' + z + ' (dashboard membaca seluruh kolom, jadi aman bila data melebihi baris ini — salin formula kolom B, H, I, J ke bawah).',
    { merge: true, color: C.muted, italic: true });

  var heads = ['Tanggal', 'ID Transaksi', 'Jenis Transaksi', 'Kategori', 'Keterangan', 'Metode Pembayaran', 'Nominal (Rp)', 'Arus Kas (Rp)', 'Bulan', 'Cek Data'];
  sh.getRange(4, 1, 1, 10).setValues([heads]).setBackground(C.navy2).setFontColor(C.white).setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(4, 32);
  sh.setFrozenRows(4);

  // Formula otomatis
  var fB = [], fH = [], fI = [], fJ = [];
  for (var k = 0; k < n; k++) {
    var r = a + k;
    fB.push(['=IF(A' + r + '="","","TRX-"&TEXT(ROW()-' + (a - 1) + ',"00000"))']);
    fH.push(['=IF(OR(C' + r + '="",G' + r + '=""),"",IF(C' + r + '="Pendapatan",G' + r + ',IF(C' + r + '="Pengeluaran",-G' + r + ',"")))']);
    fI.push(['=IF(A' + r + '="","",DATE(YEAR(A' + r + '),MONTH(A' + r + '),1))']);
    fJ.push(['=IF(A' + r + '&C' + r + '&D' + r + '&G' + r + '="","",IF(OR(A' + r + '="",C' + r + '="",D' + r + '="",G' + r + '=""),"Lengkapi",' +
      'IF(COUNTIFS(KATEGORI!$B$5:$B$64,D' + r + ',KATEGORI!$C$5:$C$64,C' + r + ')=0,"Kategori tidak cocok","OK")))']);
  }
  sh.getRange(a, 2, n, 1).setFormulas(fB);
  sh.getRange(a, 8, n, 1).setFormulas(fH);
  sh.getRange(a, 9, n, 1).setFormulas(fI);
  sh.getRange(a, 10, n, 1).setFormulas(fJ);

  // Format
  sh.getRange(a, 1, n, 10).setBorder(true, true, true, true, true, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(a, 1, n, 1).setNumberFormat(FMT_TGL).setHorizontalAlignment('center');
  sh.getRange(a, 2, n, 1).setBackground(C.bg).setFontColor(C.muted).setHorizontalAlignment('center');
  sh.getRange(a, 3, n, 1).setHorizontalAlignment('center');
  sh.getRange(a, 7, n, 1).setNumberFormat(FMT_RP).setHorizontalAlignment('right');
  sh.getRange(a, 8, n, 1).setNumberFormat(FMT_RP_NEG).setBackground(C.bg).setHorizontalAlignment('right');
  sh.getRange(a, 9, n, 1).setNumberFormat('mmmm yyyy').setBackground(C.bg).setHorizontalAlignment('center');
  sh.getRange(a, 10, n, 1).setBackground(C.bg).setHorizontalAlignment('center').setFontSize(9);

  // Data validation
  var rA = sh.getRange(a, 1, n, 1), rC = sh.getRange(a, 3, n, 1), rD = sh.getRange(a, 4, n, 1),
      rF = sh.getRange(a, 6, n, 1), rG = sh.getRange(a, 7, n, 1);
  rA.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
    .setHelpText('Isi tanggal yang valid, mis. 15/03/2026').build());
  rC.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pendapatan', 'Pengeluaran'], true)
    .setAllowInvalid(false).setHelpText('Pilih Pendapatan atau Pengeluaran').build());
  rD.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getParent().getSheetByName('KATEGORI').getRange('B5:B64'), true)
    .setAllowInvalid(false).setHelpText('Pilih kategori dari sheet KATEGORI (sesuaikan dengan Jenis Transaksi)').build());
  rF.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getParent().getSheetByName('KATEGORI').getRange('G5:G12'), true)
    .setAllowInvalid(false).setHelpText('Pilih metode pembayaran').build());
  rG.setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
    .setHelpText('Nominal tidak boleh negatif. Isi angka positif; tanda +/- ditentukan oleh Jenis Transaksi.').build());

  // Conditional formatting
  var rules = [
    aturan_(sh, 'C' + a + ':C' + z, '=$C' + a + '="Pendapatan"', C.greenBg, C.green, true),
    aturan_(sh, 'C' + a + ':C' + z, '=$C' + a + '="Pengeluaran"', C.redBg, C.red, true),
    aturan_(sh, 'H' + a + ':H' + z, '=AND(ISNUMBER($H' + a + '),$H' + a + '>0)', null, C.green, true),
    aturan_(sh, 'H' + a + ':H' + z, '=AND(ISNUMBER($H' + a + '),$H' + a + '<0)', null, C.red, true),
    aturan_(sh, 'D' + a + ':D' + z, '=$J' + a + '="Kategori tidak cocok"', C.redBg, C.red, true),
    aturan_(sh, 'J' + a + ':J' + z, '=$J' + a + '="OK"', null, C.green, false),
    aturan_(sh, 'J' + a + ':J' + z, '=AND($J' + a + '<>"",$J' + a + '<>"OK")', C.amberBg, C.red, true)
  ];
  sh.setConditionalFormatRules(rules);

  // Data simulasi
  var d = buatDataSimulasi_(), m = d.length;
  var kol = function (idx) { return d.map(function (x) { return [x[idx]]; }); };
  sh.getRange(a, 1, m, 1).setValues(kol(0));
  sh.getRange(a, 3, m, 1).setValues(kol(1));
  sh.getRange(a, 4, m, 1).setValues(kol(2));
  sh.getRange(a, 5, m, 1).setValues(kol(3));
  sh.getRange(a, 6, m, 1).setValues(kol(4));
  sh.getRange(a, 7, m, 1).setValues(kol(5));
}

// ============================== SHEET: DASHBOARD ==============================
function sumif_(jenis, dari, sampai, kat) {
  var f = 'SUMIFS(TRANSAKSI!$G:$G,TRANSAKSI!$C:$C,"' + jenis + '"';
  if (kat) f += ',TRANSAKSI!$D:$D,' + kat;
  if (dari) f += ',TRANSAKSI!$A:$A,">="&' + dari;
  if (sampai) f += ',TRANSAKSI!$A:$A,"<="&' + sampai;
  return f + ')';
}

function bangunDashboard_(sh) {
  var C = CFG.C;
  sh.setTabColor(C.navy);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:R72').setBackground(C.bg).setFontFamily('Arial').setFontSize(10).setFontColor(C.text).setVerticalAlignment('middle');
  sh.setColumnWidth(1, 16);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 108);
  sh.setColumnWidth(14, 16); sh.setColumnWidth(15, 16); sh.setColumnWidth(16, 110); sh.setColumnWidth(17, 110);
  sh.setRowHeight(1, 8);

  // Judul
  put_(sh, 'B2:M2', 'ROTI RONA  |  Dashboard Keuangan', { merge: true, bg: C.navy, color: C.white, bold: true, size: 20 });
  sh.setRowHeight(2, 44);
  put_(sh, 'B3:M3', 'Bakery & Pastry  •  Usaha kecil  •  Cash basis  •  Mata uang IDR  •  DATA SIMULASI (fiktif, untuk demonstrasi)',
    { merge: true, bg: C.navy2, color: '#C9D6E8', size: 10 });
  sh.setRowHeight(3, 24);
  sh.setRowHeight(4, 10);

  // Filter
  var lab = { color: C.muted, bold: true, ha: 'right', size: 10 };
  var inp = { bg: C.input, bold: true, ha: 'center', border: true, color: C.navy, size: 11 };
  put_(sh, 'B5', 'Tahun', lab);
  put_(sh, 'C5', CFG.TAHUN_DEFAULT, inp);
  put_(sh, 'D5', 'Bulan', lab);
  put_(sh, 'E5:F5', 'Semua Bulan', { merge: true, bg: C.input, bold: true, ha: 'center', border: true, color: C.navy, size: 11 });
  put_(sh, 'G5:H5', 'Saldo Awal (Rp)', { merge: true, color: C.muted, bold: true, ha: 'right' });
  put_(sh, 'I5:J5', CFG.SALDO_AWAL, { merge: true, bg: C.input, bold: true, ha: 'center', border: true, color: C.navy, size: 11, nf: FMT_RP });
  put_(sh, 'K5', '=$Q$3', { ha: 'right', nf: FMT_TGL, color: C.muted, bold: true });
  put_(sh, 'L5', 's/d', { ha: 'center', color: C.muted });
  put_(sh, 'M5', '=$Q$4', { ha: 'left', nf: FMT_TGL, color: C.muted, bold: true });
  sh.setRowHeight(5, 32);
  sh.setRowHeight(6, 10);
  sh.getRange('C5').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getParent().getSheetByName('KATEGORI').getRange('I5:I16'), true).setAllowInvalid(false)
    .setHelpText('Pilih tahun laporan').build());
  sh.getRange('E5').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getParent().getSheetByName('KATEGORI').getRange('K5:K17'), true).setAllowInvalid(false)
    .setHelpText('Pilih bulan, atau "Semua Bulan" untuk satu tahun penuh').build());
  sh.getRange('I5').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false).setHelpText('Kas yang dimiliki sebelum transaksi pertama dicatat').build());

  // Area perhitungan (kolom P:Q)
  put_(sh, 'P2:Q2', 'Area perhitungan (jangan diubah)', { merge: true, color: C.muted, italic: true, size: 8, wrap: true });
  put_(sh, 'P3', 'Tanggal mulai', { color: C.muted, size: 8 });
  put_(sh, 'Q3', '=IF($E$5="Semua Bulan",DATE($C$5,1,1),DATE($C$5,MATCH($E$5,KATEGORI!$K$6:$K$17,0),1))', { nf: 'yyyy-mm-dd', color: C.muted, size: 8 });
  put_(sh, 'P4', 'Tanggal akhir', { color: C.muted, size: 8 });
  put_(sh, 'Q4', '=IF($E$5="Semua Bulan",DATE($C$5,12,31),EOMONTH($Q$3,0))', { nf: 'yyyy-mm-dd', color: C.muted, size: 8 });

  var Q3 = '$Q$3', Q4 = '$Q$4';

  // KPI cards
  var kpi = [
    { col: 'B', end: 'D', t: 'TOTAL PENDAPATAN', f: '=' + sumif_('Pendapatan', Q3, Q4),
      sub: '=COUNTIFS(TRANSAKSI!$C:$C,"Pendapatan",TRANSAKSI!$A:$A,">="&' + Q3 + ',TRANSAKSI!$A:$A,"<="&' + Q4 + ')&" transaksi pemasukan"' },
    { col: 'E', end: 'G', t: 'TOTAL PENGELUARAN', f: '=' + sumif_('Pengeluaran', Q3, Q4),
      sub: '=COUNTIFS(TRANSAKSI!$C:$C,"Pengeluaran",TRANSAKSI!$A:$A,">="&' + Q3 + ',TRANSAKSI!$A:$A,"<="&' + Q4 + ')&" transaksi pengeluaran"' },
    { col: 'H', end: 'J', t: 'LABA / RUGI BERSIH', f: '=$B$8-$E$8',
      sub: '=IF($B$8=0,"Belum ada pendapatan","Margin laba bersih "&ROUND($E$13*100,1)&"%")' },
    { col: 'K', end: 'M', t: 'SALDO KAS AKHIR',
      f: '=$I$5+' + sumif_('Pendapatan', null, Q4) + '-' + sumif_('Pengeluaran', null, Q4),
      sub: 'Saldo awal + arus kas s.d. akhir periode' }
  ];
  kpi.forEach(function (k) {
    put_(sh, k.col + '7:' + k.end + '7', k.t, { merge: true, bg: C.white, color: C.muted, bold: true, size: 9, ha: 'left' });
    put_(sh, k.col + '8:' + k.end + '9', k.f, { merge: true, bg: C.white, color: C.navy, bold: true, size: 22, ha: 'left', nf: FMT_RP_NEG });
    put_(sh, k.col + '10:' + k.end + '10', k.sub, { merge: true, bg: C.white, color: C.muted, size: 9, ha: 'left' });
    sh.getRange(k.col + '7:' + k.end + '10').setBorder(true, true, true, true, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
  });
  // aksen warna atas kartu
  sh.getRange('B7:D7').setBorder(true, null, null, null, null, null, C.mid, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange('E7:G7').setBorder(true, null, null, null, null, null, C.red, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange('H7:J7').setBorder(true, null, null, null, null, null, C.green, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange('K7:M7').setBorder(true, null, null, null, null, null, C.navy, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(7, 26); sh.setRowHeight(8, 24); sh.setRowHeight(9, 24); sh.setRowHeight(10, 24); sh.setRowHeight(11, 10);

  // Indikator kondisi
  var lbl = { bg: C.white, color: C.muted, bold: true, size: 9, ha: 'left' };
  put_(sh, 'B12:D12', 'KONDISI KEUANGAN', Object.assign({ merge: true }, lbl));
  put_(sh, 'E12:G12', 'MARGIN LABA BERSIH', Object.assign({ merge: true }, lbl));
  put_(sh, 'H12:J12', 'RASIO PENGELUARAN / PENDAPATAN', Object.assign({ merge: true }, lbl));
  put_(sh, 'K12:M12', 'TRANSAKSI PERLU DICEK', Object.assign({ merge: true }, lbl));
  put_(sh, 'B13:D13', '=IF(AND($B$8=0,$E$8=0),"○ BELUM ADA DATA",IF($H$8<0,"● RUGI",IF($E$13<0.1,"● WASPADA","● SEHAT")))',
    { merge: true, bg: C.white, bold: true, size: 14, ha: 'left' });
  put_(sh, 'E13:G13', '=IF($B$8=0,0,$H$8/$B$8)', { merge: true, bg: C.white, bold: true, size: 14, ha: 'left', nf: '0.0%;[Red]-0.0%', color: C.navy });
  put_(sh, 'H13:J13', '=IF($B$8=0,0,$E$8/$B$8)', { merge: true, bg: C.white, bold: true, size: 14, ha: 'left', nf: '0.0%', color: C.navy });
  put_(sh, 'K13:M13', '=COUNTIF(TRANSAKSI!$J:$J,"Lengkapi")+COUNTIF(TRANSAKSI!$J:$J,"Kategori tidak cocok")',
    { merge: true, bg: C.white, bold: true, size: 14, ha: 'left', nf: '0" baris"', color: C.navy });
  sh.getRange('B12:M13').setBorder(true, true, true, true, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(12, 22); sh.setRowHeight(13, 30); sh.setRowHeight(14, 12);

  // Bagian grafik
  var sec = { bg: C.navy2, color: C.white, bold: true, size: 11, ha: 'left' };
  put_(sh, 'B15:H15', '="  Pendapatan vs Pengeluaran per Bulan — "&$C$5', Object.assign({ merge: true }, sec));
  put_(sh, 'I15:M15', '  Komposisi Pengeluaran (periode terpilih)', Object.assign({ merge: true }, sec));
  sh.setRowHeight(15, 26);

  // Tabel ringkasan bulanan
  var r0 = 36, rEnd = 47, rTot = 48, i;
  put_(sh, 'B34:G34', '="  Ringkasan Performa Bulanan — "&$C$5', Object.assign({ merge: true }, sec));
  put_(sh, 'H34:M34', '  Komposisi Pengeluaran per Kategori', Object.assign({ merge: true }, sec));
  sh.setRowHeight(34, 26);
  var th = { bg: C.soft, color: C.navy, bold: true, ha: 'center', size: 9, border: true };
  ['Bulan', 'Pendapatan', 'Pengeluaran', 'Laba/Rugi', 'Margin', 'Status'].forEach(function (h, k) {
    put_(sh, String.fromCharCode(66 + k) + '35', h, th);
  });
  put_(sh, 'H35:I35', 'Kategori', Object.assign({ merge: true }, th));
  put_(sh, 'J35', 'Pengeluaran', th); put_(sh, 'K35', '% Porsi', th);
  put_(sh, 'L35:M35', 'Proporsi', Object.assign({ merge: true }, th));
  sh.setRowHeight(35, 24);

  for (i = 1; i <= 12; i++) {
    var r = r0 + i - 1;
    var dari = 'DATE($C$5,' + i + ',1)', sampai = 'EOMONTH(DATE($C$5,' + i + ',1),0)';
    var cell = { bg: C.white, border: true };
    put_(sh, 'B' + r, '=INDEX(KATEGORI!$L$6:$L$17,' + i + ')', Object.assign({ ha: 'center' }, cell));
    put_(sh, 'C' + r, '=' + sumif_('Pendapatan', dari, sampai), Object.assign({ nf: FMT_RP, ha: 'right' }, cell));
    put_(sh, 'D' + r, '=' + sumif_('Pengeluaran', dari, sampai), Object.assign({ nf: FMT_RP, ha: 'right' }, cell));
    put_(sh, 'E' + r, '=IF(AND(C' + r + '=0,D' + r + '=0),"",C' + r + '-D' + r + ')', Object.assign({ nf: FMT_RP_NEG, ha: 'right' }, cell));
    put_(sh, 'F' + r, '=IF(OR(E' + r + '="",C' + r + '=0),"",E' + r + '/C' + r + ')', Object.assign({ nf: '0.0%;[Red]-0.0%', ha: 'center' }, cell));
    put_(sh, 'G' + r, '=IF(E' + r + '="","–",IF(E' + r + '<0,"Rugi",IF(F' + r + '<0.1,"Waspada","Sehat")))', Object.assign({ ha: 'center', bold: true }, cell));

    put_(sh, 'H' + r + ':I' + r, '=IFERROR(INDEX(KATEGORI!$B$5:$B$64,MATCH(' + i + ',KATEGORI!$E$5:$E$64,0)),"")', { merge: true, bg: C.white, border: true });
    put_(sh, 'J' + r, '=IF(H' + r + '="","",' + sumif_('Pengeluaran', Q3, Q4, '$H' + r) + ')', Object.assign({ nf: FMT_RP, ha: 'right' }, cell));
    put_(sh, 'K' + r, '=IF(H' + r + '="","",IF($J$' + rTot + '=0,0,J' + r + '/$J$' + rTot + '))', Object.assign({ nf: '0.0%', ha: 'center' }, cell));
    put_(sh, 'L' + r + ':M' + r, '=IF(OR(H' + r + '="",K' + r + '=""),"",REPT("█",ROUND(K' + r + '*20,0)))', { merge: true, bg: C.white, border: true, color: C.mid, size: 8 });
  }
  var tot = { bg: C.soft, bold: true, color: C.navy, border: true };
  put_(sh, 'B' + rTot, 'TOTAL', Object.assign({ ha: 'center' }, tot));
  put_(sh, 'C' + rTot, '=SUM(C' + r0 + ':C' + rEnd + ')', Object.assign({ nf: FMT_RP, ha: 'right' }, tot));
  put_(sh, 'D' + rTot, '=SUM(D' + r0 + ':D' + rEnd + ')', Object.assign({ nf: FMT_RP, ha: 'right' }, tot));
  put_(sh, 'E' + rTot, '=C' + rTot + '-D' + rTot, Object.assign({ nf: FMT_RP_NEG, ha: 'right' }, tot));
  put_(sh, 'F' + rTot, '=IF(C' + rTot + '=0,"",E' + rTot + '/C' + rTot + ')', Object.assign({ nf: '0.0%;[Red]-0.0%', ha: 'center' }, tot));
  put_(sh, 'G' + rTot, '=IF(E' + rTot + '<0,"Rugi",IF(C' + rTot + '=0,"–","Laba"))', Object.assign({ ha: 'center' }, tot));
  put_(sh, 'H' + rTot + ':I' + rTot, 'TOTAL', Object.assign({ merge: true, ha: 'center' }, tot));
  put_(sh, 'J' + rTot, '=SUM(J' + r0 + ':J' + rEnd + ')', Object.assign({ nf: FMT_RP, ha: 'right' }, tot));
  put_(sh, 'K' + rTot, '=IF(J' + rTot + '=0,"",1)', Object.assign({ nf: '0.0%', ha: 'center' }, tot));
  put_(sh, 'L' + rTot + ':M' + rTot, '=IF(ABS($J$' + rTot + '-$E$8)<0.5,"✓ cocok dgn KPI","⚠ ada kategori tak terdaftar")',
    Object.assign({ merge: true, ha: 'center', size: 8 }, tot));
  sh.getRange('B' + r0 + ':M' + rTot).setFontSize(9);

  // Perbandingan
  put_(sh, 'B50:G50', '="  Perbandingan Pendapatan vs Pengeluaran — Periode Terpilih"', Object.assign({ merge: true }, sec));
  put_(sh, 'H50:M50', '  Rincian & Sorotan Otomatis', Object.assign({ merge: true }, sec));
  sh.setRowHeight(50, 26);
  put_(sh, 'H51:I51', 'Komponen', Object.assign({ merge: true }, th));
  put_(sh, 'J51', 'Nominal', th); put_(sh, 'K51', '% Pendapatan', th);
  put_(sh, 'L51:M51', 'Keterangan', Object.assign({ merge: true }, th));
  var cmp = [['Pendapatan', '$B$8', 'Uang masuk'], ['Pengeluaran', '$E$8', 'Uang keluar'], ['Laba/Rugi', '$H$8', 'Selisih']];
  cmp.forEach(function (x, k) {
    var r = 52 + k;
    put_(sh, 'H' + r + ':I' + r, x[0], { merge: true, bg: C.white, border: true, bold: true });
    put_(sh, 'J' + r, '=' + x[1], { bg: C.white, border: true, nf: FMT_RP_NEG, ha: 'right' });
    put_(sh, 'K' + r, '=IF($B$8=0,0,J' + r + '/$B$8)', { bg: C.white, border: true, nf: '0.0%;[Red]-0.0%', ha: 'center' });
    put_(sh, 'L' + r + ':M' + r, x[2], { merge: true, bg: C.white, border: true, color: C.muted, ha: 'center' });
  });
  var ins = { merge: true, bg: C.white, border: true, size: 9 };
  put_(sh, 'H56:I56', 'Sorotan', Object.assign({ merge: true }, th));
  put_(sh, 'J56:K56', 'Hasil', Object.assign({ merge: true }, th));
  put_(sh, 'L56:M56', 'Nilai', Object.assign({ merge: true }, th));
  put_(sh, 'H57:I57', 'Bulan laba tertinggi', ins);
  put_(sh, 'J57:K57', '=IF(COUNT($E$' + r0 + ':$E$' + rEnd + ')=0,"–",INDEX($B$' + r0 + ':$B$' + rEnd + ',MATCH(MAX($E$' + r0 + ':$E$' + rEnd + '),$E$' + r0 + ':$E$' + rEnd + ',0)))', Object.assign({ ha: 'center', bold: true }, ins));
  put_(sh, 'L57:M57', '=IF(COUNT($E$' + r0 + ':$E$' + rEnd + ')=0,0,MAX($E$' + r0 + ':$E$' + rEnd + '))', Object.assign({ nf: FMT_RP_NEG, ha: 'right' }, ins));
  put_(sh, 'H58:I58', 'Bulan laba terendah', ins);
  put_(sh, 'J58:K58', '=IF(COUNT($E$' + r0 + ':$E$' + rEnd + ')=0,"–",INDEX($B$' + r0 + ':$B$' + rEnd + ',MATCH(MIN($E$' + r0 + ':$E$' + rEnd + '),$E$' + r0 + ':$E$' + rEnd + ',0)))', Object.assign({ ha: 'center', bold: true }, ins));
  put_(sh, 'L58:M58', '=IF(COUNT($E$' + r0 + ':$E$' + rEnd + ')=0,0,MIN($E$' + r0 + ':$E$' + rEnd + '))', Object.assign({ nf: FMT_RP_NEG, ha: 'right' }, ins));
  put_(sh, 'H59:I59', 'Biaya terbesar', ins);
  put_(sh, 'J59:K59', '=IF($J$' + rTot + '=0,"–",INDEX($H$' + r0 + ':$H$' + rEnd + ',MATCH(MAX($J$' + r0 + ':$J$' + rEnd + '),$J$' + r0 + ':$J$' + rEnd + ',0)))', Object.assign({ ha: 'center', bold: true }, ins));
  put_(sh, 'L59:M59', '=IF($J$' + rTot + '=0,0,MAX($J$' + r0 + ':$J$' + rEnd + '))', Object.assign({ nf: FMT_RP, ha: 'right' }, ins));
  put_(sh, 'H60:I60', 'Rata-rata laba/bulan', ins);
  put_(sh, 'J60:K60', 'bulan aktif', Object.assign({ ha: 'center', color: C.muted }, ins));
  put_(sh, 'L60:M60', '=IFERROR(AVERAGE($E$' + r0 + ':$E$' + rEnd + '),0)', Object.assign({ nf: FMT_RP_NEG, ha: 'right' }, ins));
  put_(sh, 'H61:I61', 'Jumlah transaksi', ins);
  put_(sh, 'J61:K61', 'dalam periode', Object.assign({ ha: 'center', color: C.muted }, ins));
  put_(sh, 'L61:M61', '=COUNTIFS(TRANSAKSI!$A:$A,">="&' + Q3 + ',TRANSAKSI!$A:$A,"<="&' + Q4 + ')', Object.assign({ nf: '#,##0', ha: 'right' }, ins));
  put_(sh, 'H63:M65', '="Catatan: Saldo Kas Akhir = Saldo Awal + seluruh pendapatan − seluruh pengeluaran sejak awal pencatatan s.d. "&TEXT($Q$4,"dd/mm/yyyy")&". Data bersifat SIMULASI."',
    { merge: true, color: C.muted, italic: true, size: 9, wrap: true, ha: 'left', va: 'top' });

  put_(sh, 'B68:M68', 'ROTI RONA adalah usaha fiktif. Seluruh transaksi bersifat simulasi untuk portofolio Data Analytics & Spreadsheet Automation.',
    { merge: true, color: C.muted, italic: true, size: 8, ha: 'center' });

  // Conditional formatting dashboard
  sh.setConditionalFormatRules([
    aturan_(sh, 'H8:J9', '=$H$8<0', null, C.red, true),
    aturan_(sh, 'H8:J9', '=$H$8>=0', null, C.green, true),
    aturan_(sh, 'K8:M9', '=$K$8<0', null, C.red, true),
    aturan_(sh, 'B13:D13', '=ISNUMBER(SEARCH("SEHAT",$B$13))', C.greenBg, C.green, true),
    aturan_(sh, 'B13:D13', '=ISNUMBER(SEARCH("RUGI",$B$13))', C.redBg, C.red, true),
    aturan_(sh, 'B13:D13', '=ISNUMBER(SEARCH("WASPADA",$B$13))', C.amberBg, C.amber, true),
    aturan_(sh, 'K13:M13', '=$K$13>0', C.amberBg, C.red, true),
    aturan_(sh, 'G' + r0 + ':G' + rTot, '=OR(G' + r0 + '="Sehat",G' + r0 + '="Laba")', C.greenBg, C.green, true),
    aturan_(sh, 'G' + r0 + ':G' + rTot, '=G' + r0 + '="Rugi"', C.redBg, C.red, true),
    aturan_(sh, 'G' + r0 + ':G' + rTot, '=G' + r0 + '="Waspada"', C.amberBg, C.amber, true),
    aturan_(sh, 'L' + rTot + ':M' + rTot, '=ISNUMBER(SEARCH("⚠",$L$' + rTot + '))', C.redBg, C.red, true),
    aturan_(sh, 'L' + rTot + ':M' + rTot, '=ISNUMBER(SEARCH("✓",$L$' + rTot + '))', C.greenBg, C.green, true),
    aturan_(sh, 'J52:J54', '=J52<0', null, C.red, true)
  ]);

  // Grafik
  var c1 = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('B35:D47'))
    .setNumHeaders(1)
    .setPosition(16, 2, 4, 4)
    .setOption('width', 7 * 108 - 8).setOption('height', 17 * 21 - 6)
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', [C.mid, C.red])
    .setOption('vAxis', { format: 'short', gridlines: { color: '#E6E9EE' } })
    .setOption('backgroundColor', C.white)
    .build();
  sh.insertChart(c1);

  var c2 = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('H35:H47')).addRange(sh.getRange('J35:J47'))
    .setNumHeaders(1)
    .setPosition(16, 9, 4, 4)
    .setOption('width', 5 * 108 - 8).setOption('height', 17 * 21 - 6)
    .setOption('pieHole', 0.5)
    .setOption('legend', { position: 'right', textStyle: { fontSize: 9 } })
    .setOption('pieSliceText', 'none')
    .setOption('colors', ['#0B2545', '#1D3E6E', '#2F5D9E', '#4C7FBF', '#7BA3D1', '#A9C1E0', '#C9D6E8', '#E07B7B', '#E8A33D', '#6FAF8E', '#8E8E93', '#BFBFC4'])
    .setOption('backgroundColor', C.white)
    .build();
  sh.insertChart(c2);

  var c3 = sh.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('H51:H54')).addRange(sh.getRange('J51:J54'))
    .setNumHeaders(1)
    .setPosition(51, 2, 4, 4)
    .setOption('width', 6 * 108 - 8).setOption('height', 15 * 21 - 6)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [C.mid])
    .setOption('hAxis', { format: 'short' })
    .setOption('backgroundColor', C.white)
    .build();
  sh.insertChart(c3);

  sh.setFrozenRows(3);
}

// ============================== SHEET: PETUNJUK ==============================
function bangunPetunjuk_(sh) {
  var C = CFG.C;
  sh.setTabColor('#8E8E93');
  sh.setHiddenGridlines(true);
  sh.getRange('A1:C80').setFontFamily('Arial').setFontSize(10).setFontColor(C.text).setVerticalAlignment('top');
  sh.setColumnWidth(1, 16); sh.setColumnWidth(2, 230); sh.setColumnWidth(3, 640);
  put_(sh, 'B1:C1', 'PETUNJUK PENGGUNAAN — ROTI RONA', { merge: true, bg: C.navy, color: C.white, bold: true, size: 14, va: 'middle' });
  sh.setRowHeight(1, 36);
  put_(sh, 'B2:C2', 'Sistem pembukuan cash basis: transaksi dicatat saat uang benar-benar diterima atau dibayar. Semua angka di DASHBOARD dihitung otomatis dari sheet TRANSAKSI.',
    { merge: true, color: C.muted, italic: true, wrap: true, va: 'middle' });
  sh.setRowHeight(2, 34);

  var isi = [
    ['A. MENAMBAH TRANSAKSI', null],
    ['Langkah 1', 'Buka sheet TRANSAKSI. Cari baris kosong pertama di bawah data terakhir (kolom Tanggal kosong).'],
    ['Langkah 2', 'Isi Tanggal (mis. 15/03/2026), lalu pilih Jenis Transaksi: Pendapatan (uang masuk) atau Pengeluaran (uang keluar).'],
    ['Langkah 3', 'Pilih Kategori, isi Keterangan singkat, pilih Metode Pembayaran, lalu isi Nominal sebagai angka positif tanpa titik/"Rp".'],
    ['Otomatis', 'ID Transaksi, Arus Kas, Bulan, dan Cek Data terisi sendiri. Jangan ditimpa. Dashboard langsung ikut berubah.'],
    ['Catatan', 'Nominal tidak boleh negatif. Tanda minus pada Arus Kas ditentukan sistem dari Jenis Transaksi.'],
    ['B. MEMILIH KATEGORI', null],
    ['Aturan', 'Kategori Pendapatan: Penjualan Produk, Pendapatan Lainnya. Kategori lain adalah Pengeluaran.'],
    ['Cek kecocokan', 'Bila Kategori tidak sesuai Jenis (mis. "Sewa" pada Pendapatan), sel Kategori berwarna merah dan kolom Cek Data berisi "Kategori tidak cocok". Perbaiki sampai kolom itu "OK".'],
    ['C. MEMBACA DASHBOARD', null],
    ['Kartu KPI', 'Total Pendapatan, Total Pengeluaran, Laba/Rugi Bersih (hijau = untung, merah = rugi), dan Saldo Kas Akhir untuk periode terpilih.'],
    ['Kondisi Keuangan', 'SEHAT = laba dengan margin ≥ 10%. WASPADA = laba tetapi margin < 10%. RUGI = pengeluaran melebihi pendapatan.'],
    ['Grafik', 'Kiri: pendapatan vs pengeluaran tiap bulan pada tahun terpilih. Kanan: komposisi pengeluaran per kategori. Bawah: perbandingan total periode.'],
    ['Tabel bulanan', 'Margin = Laba ÷ Pendapatan. Status per bulan: Sehat / Waspada / Rugi. Bulan tanpa transaksi ditampilkan "–".'],
    ['Transaksi perlu dicek', 'Jumlah baris TRANSAKSI yang belum lengkap atau kategorinya tidak cocok. Targetnya 0.'],
    ['D. MENGUBAH PERIODE LAPORAN', null],
    ['Tahun', 'Di DASHBOARD pilih Tahun pada sel kuning (kiri atas). KPI, tabel, dan grafik mengikuti tahun tersebut.'],
    ['Bulan', 'Pilih satu bulan, atau "Semua Bulan" untuk satu tahun penuh. Grafik bulanan tetap menampilkan 12 bulan pada tahun terpilih.'],
    ['Saldo Awal', 'Ubah sel Saldo Awal (kuning) sesuai kas yang dimiliki sebelum transaksi pertama dicatat. Saldo Kas Akhir menghitung seluruh transaksi sejak awal s.d. akhir periode.'],
    ['E. MENAMBAH KATEGORI BARU', null],
    ['Langkah', 'Buka sheet KATEGORI, isi baris kosong pertama: kolom Kategori (nama) dan Jenis (Pendapatan/Pengeluaran). Kategori baru langsung muncul di dropdown TRANSAKSI.'],
    ['Batas', 'Dashboard menampilkan hingga 12 kategori pengeluaran. Jika lebih, kotak "TOTAL" di tabel kategori menandai ⚠ — gabungkan kategori yang jarang dipakai.'],
    ['Jangan', 'Jangan mengganti nama kategori yang sudah dipakai transaksi; transaksi lama tidak ikut berubah dan akan ditandai "Kategori tidak cocok".'],
    ['F. MEMPERBAIKI KESALAHAN PENCATATAN', null],
    ['Salah nominal/kategori', 'Cukup timpa nilai pada sel yang salah di baris tersebut. Semua KPI dan grafik ikut diperbarui.'],
    ['Transaksi dobel / batal', 'Kosongkan kolom Tanggal, Jenis, Kategori, Keterangan, Metode, dan Nominal pada baris itu (tekan Delete). Hindari menghapus seluruh baris agar formula tidak hilang; ID Transaksi bergantung pada nomor baris.'],
    ['Salah jenis', 'Ubah kolom Jenis Transaksi lalu sesuaikan Kategorinya.'],
    ['Mulai data asli', 'Menu ROTI RONA → "Hapus data simulasi", atau kosongkan data demo secara manual. Setelah itu atur Saldo Awal.'],
    ['G. KETERBATASAN', null],
    ['Ringkas', 'Cash basis (tanpa piutang/utang/persediaan/penyusutan). Bukan pengganti laporan akuntansi resmi atau pelaporan pajak.']
  ];
  var r = 4;
  isi.forEach(function (x) {
    if (x[1] === null) {
      put_(sh, 'B' + r + ':C' + r, x[0], { merge: true, bg: C.navy2, color: C.white, bold: true, va: 'middle' });
      sh.setRowHeight(r, 26);
    } else {
      put_(sh, 'B' + r, x[0], { bold: true, color: C.navy, bg: C.white, border: true, wrap: true });
      put_(sh, 'C' + r, x[1], { bg: C.white, border: true, wrap: true });
      sh.setRowHeight(r, x[1].length > 110 ? 48 : 30);
    }
    r++;
  });
  sh.getRange('A1:C' + r).setHorizontalAlignment('left');
}

// ============================== DATA SIMULASI ==============================
/** PRNG deterministik (mulberry32) agar data demo selalu sama. */
function mulberry32_(seed) {
  return function () {
    seed = (seed + 0x6D2B79F5) | 0;
    var t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Menghasilkan transaksi fiktif Jan 2025 – Okt 2026. Return: [Tanggal, Jenis, Kategori, Keterangan, Metode, Nominal][] */
function buatDataSimulasi_() {
  var rand = mulberry32_(20250101);
  var rows = [];
  var r = function (a, b) { return a + rand() * (b - a); };
  var rd = function (n, step) { return Math.round(n / step) * step; };
  var add = function (d, jenis, kat, ket, metode, nominal) {
    rows.push([new Date(d.getFullYear(), d.getMonth(), d.getDate()), jenis, kat, '[SIMULASI] ' + ket, metode, Math.round(nominal)]);
  };
  // Pengali musiman per bulan (idx 0 = Jan 2025 ... 21 = Okt 2026)
  var SEAS = [0.68, 0.85, 1.25, 1.05, 0.98, 0.95, 0.97, 1.0, 1.0, 1.03, 1.05, 1.22,
              0.92, 1.12, 1.22, 1.0, 1.02, 0.98, 1.0, 1.04, 1.05, 1.08];
  var BAHAN = { 1: 'Tepung, gula & ragi', 3: 'Mentega, susu & telur', 5: 'Cokelat, topping & selai' };
  var akhir = new Date(2026, 9, 31);

  for (var d = new Date(2025, 0, 1); d <= akhir; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
    var y = d.getFullYear(), m = d.getMonth(), tgl = d.getDate(), dow = d.getDay();
    var idx = (y - 2025) * 12 + m;
    var mult = SEAS[idx] * (1 + 0.008 * idx);
    var base = dow === 0 ? 1950000 : (dow === 6 ? 1750000 : 1350000);

    // Penjualan harian
    var u = rand();
    add(d, 'Pendapatan', 'Penjualan Produk', 'Penjualan harian toko (kasir)',
      u < 0.45 ? 'Tunai' : (u < 0.85 ? 'QRIS' : 'Transfer Bank'), rd(base * mult * r(0.88, 1.12), 1000));
    // Pesanan katering / kue ulang tahun
    if ((tgl === 7 || tgl === 16 || tgl === 25) && rand() < 0.7) {
      add(d, 'Pendapatan', 'Pendapatan Lainnya', 'Pesanan kue ulang tahun / katering', 'Transfer Bank',
        rd(r(350000, 1800000) * (mult > 1.15 ? 1.4 : 1), 5000));
    }
    if (tgl === 28) add(d, 'Pendapatan', 'Pendapatan Lainnya', 'Penjualan dus & karung bekas', 'Tunai', rd(r(60000, 150000), 5000));
    if ((y === 2025 && m === 2 && tgl === 20) || (y === 2026 && m === 2 && tgl === 10)) {
      add(d, 'Pendapatan', 'Penjualan Produk', 'Pesanan parcel kue kering Lebaran', 'Transfer Bank', rd(r(4500000, 5500000), 50000));
    }

    // Bahan baku (Sen/Rab/Jum)
    if (BAHAN[dow]) {
      add(d, 'Pengeluaran', 'Pembelian Bahan Baku', 'Belanja ' + BAHAN[dow], rand() < 0.6 ? 'Tunai' : 'Transfer Bank',
        rd(0.40 * 10450000 * mult / 3 * r(0.9, 1.1), 5000));
    }
    // Kemasan
    if (tgl === 5 || tgl === 20) add(d, 'Pengeluaran', 'Kemasan', 'Box kue, plastik & label', 'Transfer Bank', rd(r(550000, 900000) * (mult > 1.1 ? 1.3 : 1), 5000));
    // Gaji
    if (tgl === 28) {
      var f = y === 2026 ? 1.05 : 1;
      add(d, 'Pengeluaran', 'Gaji', 'Gaji karyawan produksi (2 orang)', 'Transfer Bank', 5600000 * f);
      add(d, 'Pengeluaran', 'Gaji', 'Gaji kasir & penjualan (1 orang)', 'Transfer Bank', 2800000 * f);
    }
    if ((y === 2025 && m === 2 && tgl === 24) || (y === 2026 && m === 2 && tgl === 13)) {
      add(d, 'Pengeluaran', 'Gaji', 'THR karyawan', 'Transfer Bank', y === 2026 ? 8820000 : 8400000);
    }
    // Sewa
    if (tgl === 1) add(d, 'Pengeluaran', 'Sewa', 'Sewa tempat usaha bulanan', 'Transfer Bank', y === 2026 ? 4750000 : 4500000);
    // Listrik & air
    if (tgl === 8) add(d, 'Pengeluaran', 'Listrik dan Air', 'Token listrik oven & toko', 'E-Wallet', rd(r(1100000, 1500000), 5000));
    if (tgl === 12) add(d, 'Pengeluaran', 'Listrik dan Air', 'Tagihan air PDAM', 'Transfer Bank', rd(r(250000, 340000), 5000));
    // Transportasi (Sabtu)
    if (dow === 6) add(d, 'Pengeluaran', 'Transportasi', 'BBM & ongkos kirim pesanan', 'Tunai', rd(r(250000, 400000), 5000));
    // Marketing
    if (tgl === 10) add(d, 'Pengeluaran', 'Marketing', 'Iklan Instagram / Meta Ads', 'E-Wallet', rd(r(400000, 900000), 5000));
    if (tgl === 22 && rand() < 0.3) add(d, 'Pengeluaran', 'Marketing', 'Endorse / konten influencer lokal', 'Transfer Bank', 500000);
    // Operasional
    if (tgl === 3 || tgl === 17) add(d, 'Pengeluaran', 'Operasional', 'Gas LPG 12 kg (2 tabung)', 'Tunai', 380000);
    if (tgl === 15) add(d, 'Pengeluaran', 'Operasional', 'Alat kebersihan & sabun', 'Tunai', rd(r(150000, 250000), 5000));
    if (tgl === 21) add(d, 'Pengeluaran', 'Operasional', 'Internet & pulsa', 'Transfer Bank', 350000);
    // Lain-lain
    if (tgl === 13 && rand() < 0.6) add(d, 'Pengeluaran', 'Pengeluaran Lainnya', 'Biaya tak terduga', 'Tunai', rd(r(50000, 400000), 5000));
    if (y === 2025 && m === 0 && tgl === 22) add(d, 'Pengeluaran', 'Pengeluaran Lainnya', 'Ganti mixer rusak', 'Transfer Bank', 3800000);
    if (y === 2025 && m === 4 && tgl === 14) add(d, 'Pengeluaran', 'Operasional', 'Servis oven', 'Transfer Bank', 4500000);
  }
  return rows;
}
