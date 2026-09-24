/**
 * Aplikasi Database Pelanggaran Siswa
 * UPT SPF SMP Negeri 2 Sunggal T.P. 2026/2027
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SIKAD - UPT SPF SMPN 2 Sunggal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Mengambil master data siswa dan jenis pelanggaran untuk dropdown
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Data Siswa
  const sheetSiswa = ss.getSheetByName('Master_Siswa');
  const rawSiswa = sheetSiswa.getRange(2, 1, sheetSiswa.getLastRow() - 1, 5).getValues();
  const listSiswa = rawSiswa.map(row => ({
    nis: String(row[0]),
    nama: row[1],
    kelas: row[2],
    lp: row[3],
    hp: row[4] || '-'
  }));

  // Data Pelanggaran
  const sheetPelanggaran = ss.getSheetByName('Master_Pelanggaran');
  const rawPelanggaran = sheetPelanggaran.getRange(2, 1, sheetPelanggaran.getLastRow() - 1, 3).getValues();
  const listPelanggaran = rawPelanggaran.map(row => ({
    id: row[0],
    jenis: row[1],
    skor: parseInt(row[2])
  }));

  return { siswa: listSiswa, pelanggaran: listPelanggaran };
}

// Fungsi Input Pelanggaran
function simpanPelanggaran(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLog = ss.getSheetByName('Log_Pelanggaran');
    
    const idLog = 'LOG-' + new Date().getTime();
    const waktuInput = new Date();
    
    sheetLog.appendRow([
      idLog,
      waktuInput,
      data.tanggal,
      data.nis,
      data.nama,
      data.kelas,
      data.idPelanggaran,
      data.jenisPelanggaran,
      parseInt(data.skor),
      data.catatan || '-',
      data.petugas || 'Guru Kesiswaan'
    ]);

    // Hitung total poin terbaru siswa tersebut
    const totalSkorTerbaru = hitungTotalSkorSiswa(data.nis);
    const statusTerbaru = dapatkanStatusSanksi(totalSkorTerbaru);

    return {
      success: true,
      message: 'Pelanggaran berhasil dicatat!',
      totalSkor: totalSkorTerbaru,
      status: statusTerbaru
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Utility: Hitung Total Poin Siswa
function hitungTotalSkorSiswa(nis) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName('Log_Pelanggaran');
  if (sheetLog.getLastRow() < 2) return 0;
  
  const data = sheetLog.getRange(2, 4, sheetLog.getLastRow() - 1, 6).getValues(); // NIS di col 4, Skor di col 9
  let total = 0;
  data.forEach(row => {
    if (String(row[0]) === String(nis)) {
      total += parseInt(row[5] || 0);
    }
  });
  return total;
}

// Utility: Tentukan Status Sanksi
function dapatkanStatusSanksi(skor) {
  if (skor >= 100) return 'Sanksi Maksimal (Dikembalikan ke Orang Tua)';
  if (skor >= 75) return 'SP 3 & Skorsing 1 Minggu (Peringatan Keras)';
  if (skor >= 50) return 'Panggilan Orang Tua II (Surat Perjanjian II)';
  if (skor >= 25) return 'Panggilan Orang Tua I (Surat Perjanjian I)';
  return 'Aman / Pembinaan Normal';
}

// Mengambil Data Rekap untuk Dashboard Web
function getRekapData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRekap = ss.getSheetByName('Rekap_Siswa');
  if (sheetRekap.getLastRow() < 2) return [];
  
  const values = sheetRekap.getRange(2, 1, sheetRekap.getLastRow() - 1, 6).getValues();
  return values.map(row => ({
    nis: String(row[0]),
    nama: row[1],
    kelas: row[2],
    lp: row[3],
    totalSkor: parseInt(row[4] || 0),
    status: row[5]
  })).filter(item => item.totalSkor > 0); // Hanya tampilkan yang memiliki poin
}
