/**
 * File: Code.gs
 * Deskripsi: Backend untuk "Spendubaya Edutest Report" menggunakan Google Apps Script.
 * Script ini berfungsi sebagai API untuk mengambil data hasil ujian dari Google Sheet.
 *
 * PENTING:
 * 1. Pastikan nama Sheet data Anda adalah "HasilUjian"
 * 2. Pastikan struktur kolom sesuai:
 * A: No_Peserta, B: Nama_Siswa, C: Kelas_Siswa, D: Nama_WaliKelas
 * E-O: 11 Mata Pelajaran
 */

// Peta Mata Pelajaran (sesuai urutan kolom E sampai O, index 4 sampai 14)
const SUBJECT_MAP = {
  4: "Bahasa_Indonesia",
  5: "Matematika",
  6: "IPA",
  7: "IPS",
  8: "PAI",
  9: "PPKN",
  10: "PJOK",
  11: "Bahasa_Inggris",
  12: "Seni_Budaya",
  13: "Informatika",
  14: "Bahasa_Madura",
};

/**
 * =================================================================
 * FUNGSI SETUP AWAL: Membuat sheet dan header secara otomatis
 * =================================================================
 */
function Setup_awal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "HasilUjian";

  // 1. Cek dan hapus sheet lama jika ada
  const existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
    Logger.log(`Sheet lama "${sheetName}" dihapus.`);
  }

  // 2. Buat sheet baru
  const newSheet = ss.insertSheet(sheetName, 0); // Buat di posisi pertama

  // 3. Tentukan data header
  const headers = [
    "No_Peserta",
    "Nama_Siswa",
    "Kelas_Siswa",
    "Nama_WaliKelas",
    // Mata Pelajaran (sesuai urutan E-O)
    "Bahasa Indonesia",
    "Matematika",
    "Ilmu Pengetahuan Alam",
    "Ilmu Pengetahuan Sosial",
    "Pendidikan Agama Islam dan Budi Pekerti",
    "Pendidikan Pancasila dan Kewarganegaraan",
    "Pendidikan Jasmani, Olahraga, dan Kesehatan",
    "Bahasa Inggris",
    "Seni Budaya",
    "Informatika",
    "Bahasa Madura",
  ];

  // 4. Masukkan header ke Baris 1
  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 5. Formatting header (opsional, agar lebih rapi)
  const headerRange = newSheet.getRange("A1:O1");
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#d9ead3"); // Warna hijau muda
  headerRange.setHorizontalAlignment("center");

  // Auto-resize kolom (opsional)
  newSheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert(
    `Sheet data "${sheetName}" berhasil dibuat dengan struktur header yang benar.`
  );
}

/**
 * Fungsi utama untuk menangani permintaan HTTP GET.
 * Berfungsi sebagai router API.
 * @param {Object} e Event object dari permintaan HTTP.
 * @returns {GoogleAppsScript.Content.TextOutput} Output JSON.
 */
function doGet(e) {
  const action = e.parameter.action;

  // Menambahkan pengamanan dasar jika sheet belum ada
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HasilUjian");
  if (!sheet) {
    return createJsonOutput({
      status: "error",
      message:
        "Sheet 'HasilUjian' tidak ditemukan. Jalankan fungsi Setup_awal terlebih dahulu.",
    });
  }

  try {
    let result = {};
    if (action === "getStudentReport") {
      const noPeserta = e.parameter.noPeserta;
      result = getStudentReport(noPeserta);
    } else if (action === "getRecapReport") {
      const kelas = e.parameter.kelas;
      const mapel = e.parameter.mapel;
      result = getRecapReport(kelas, mapel);
    } else {
      return createJsonOutput({
        status: "error",
        message: "Aksi tidak dikenal.",
      });
    }

    return createJsonOutput(result);
  } catch (error) {
    Logger.log("Kesalahan dalam doGet: " + error.toString());
    return createJsonOutput({
      status: "error",
      message: "Kesalahan server: " + error.message,
    });
  }
}

/**
 * Mengambil data siswa berdasarkan No_Peserta.
 * @param {string} noPeserta No. Peserta Siswa.
 * @returns {Object} Data siswa atau not_found status.
 */
function getStudentReport(noPeserta) {
  if (!noPeserta) {
    return { status: "error", message: "Parameter No_Peserta harus diisi." };
  }

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HasilUjian");

  // Ambil semua data kecuali header (baris 1)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    // Hanya ada header atau kosong
    return { status: "not_found", message: "Sheet data kosong." };
  }

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();

  // Kolom No_Peserta berada di index 0 (Kolom A)
  const studentData = data.find(
    (row) => row[0].toString().trim() === noPeserta.toString().trim()
  );

  if (!studentData) {
    return { status: "not_found", message: "Data siswa tidak ditemukan." };
  }

  const report = {
    No_Peserta: studentData[0],
    Nama_Siswa: studentData[1],
    Kelas_Siswa: studentData[2],
    Nama_WaliKelas: studentData[3],
  };

  // Loop untuk nilai mata pelajaran (index 4 sampai 14)
  for (let i = 4; i <= 14; i++) {
    const subjectKey = SUBJECT_MAP[i];
    // Pastikan nilai dikonversi ke Number jika memungkinkan
    const rawValue = studentData[i];
    const score =
      typeof rawValue === "number"
        ? rawValue
        : rawValue
        ? parseFloat(rawValue)
        : null;
    report[subjectKey] = score !== null && !isNaN(score) ? score : ""; // Atur nilai kosong ke string kosong
  }

  return { status: "success", data: report };
}

/**
 * Mengambil rekap nilai untuk mata pelajaran dan kelas tertentu.
 * @param {string} kelas Kelas Siswa.
 * @param {string} mapel Kunci Mata Pelajaran (e.g., Bahasa_Indonesia).
 * @returns {Object} Array data rekap atau not_found status.
 */
function getRecapReport(kelas, mapel) {
  if (!kelas || !mapel) {
    return {
      status: "error",
      message: "Parameter Kelas dan Mapel harus diisi.",
    };
  }

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HasilUjian");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    // Hanya ada header atau kosong
    return { status: "not_found", message: "Sheet data kosong." };
  }

  // Ambil semua data kecuali header (baris 1)
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();

  // Tentukan index kolom mata pelajaran
  const subjectIndex = Object.keys(SUBJECT_MAP).find(
    (key) => SUBJECT_MAP[key] === mapel
  );
  if (!subjectIndex) {
    return { status: "error", message: "Kunci mata pelajaran tidak valid." };
  }
  const mapelColIndex = parseInt(subjectIndex); // Konversi kembali ke integer

  const recapData = [];

  // Kolom Kelas berada di index 2 (Kolom C)
  data.forEach((row) => {
    if (
      row[2] &&
      row[2].toString().trim().toUpperCase() ===
        kelas.toString().trim().toUpperCase()
    ) {
      const rawValue = row[mapelColIndex];
      const score =
        typeof rawValue === "number"
          ? rawValue
          : rawValue
          ? parseFloat(rawValue)
          : null;

      recapData.push({
        No_Peserta: row[0],
        Nama_Siswa: row[1],
        Kelas_Siswa: row[2],
        [mapel]: score !== null && !isNaN(score) ? score : "", // Atur nilai kosong ke string kosong
      });
    }
  });

  if (recapData.length === 0) {
    return {
      status: "not_found",
      message: "Tidak ada data siswa ditemukan untuk kelas ini.",
    };
  }

  return { status: "success", data: recapData };
}

/**
 * Fungsi pembantu untuk membuat output JSON dengan header yang tepat.
 * @param {Object} data Objek data yang akan dikonversi menjadi JSON.
 * @returns {GoogleAppsScript.Content.TextOutput} Output JSON.
 */
function createJsonOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON
  );
}
