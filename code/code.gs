const SUBJECT_MAP = {
  // Gunakan nama kolom yang disingkat (tanpa spasi) sebagai kunci di JS
  // Angka indeks kolom 4-14 tidak perlu di sini, tapi di fungsi getStudentReport
  Bahasa_Indonesia: "Bahasa Indonesia",
  Matematika: "Matematika",
  IPA: "Ilmu Pengetahuan Alam",
  IPS: "Ilmu Pengetahuan Sosial",
  PAI: "Pendidikan Agama Islam dan Budi Pekerti",
  PPKN: "Pendidikan Pancasila dan Kewarganegaraan",
  PJOK: "Pendidikan Jasmani, Olahraga, dan Kesehatan",
  Bahasa_Inggris: "Bahasa Inggris",
  Seni_Budaya: "Seni Budaya",
  Informatika: "Informatika",
  Bahasa_Madura: "Bahasa Madura",
};

// Peta konstanta yang menghubungkan kunci singkat (JS) ke Index Kolom (Sheet HasilUjian)
// Kolom A=0, B=1, C=2, D=3, E=4, dst.
const COL_INDEX_MAP = {
  No_Peserta: 0,
  Nama_Siswa: 1,
  Kelas_Siswa: 2,
  Nama_WaliKelas: 3,
  // Mata Pelajaran (E-O = Index 4-14)
  Bahasa_Indonesia: 4,
  Matematika: 5,
  IPA: 6,
  IPS: 7,
  PAI: 8,
  PPKN: 9,
  PJOK: 10,
  Bahasa_Inggris: 11,
  Seni_Budaya: 12,
  Informatika: 13,
  Bahasa_Madura: 14,
};

//-------------------------------------------------------------
// FUNGSI BARU: Mengambil data Wali Kelas dari DataBase_Utama
//-------------------------------------------------------------
/**
 * Mengambil pasangan Kelas dan Wali Kelas dari sheet DataBase_Utama.
 * @returns {Object} Peta (Map) dengan kunci: Nama Kelas, Nilai: Nama Wali Kelas.
 */
function getWaliKelasMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DataBase_Utama");

  if (!sheet) {
    Logger.log("Sheet 'DataBase_Utama' tidak ditemukan.");
    return {};
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  // Asumsi header di DataBase_Utama adalah Kolom A: Kelas, Kolom B: Wali_Kelas
  // Ambil data (A2:B terakhir)
  const range = sheet.getRange(2, 1, lastRow - 1, 2);
  const data = range.getValues();

  const waliKelasMap = {};
  data.forEach((row) => {
    const kelas = row[0] ? row[0].toString().trim().toUpperCase() : "";
    const waliKelas = row[1] ? row[1].toString().trim() : "";

    if (kelas && waliKelas) {
      waliKelasMap[kelas] = waliKelas;
    }
  });

  return waliKelasMap;
}

//-------------------------------------------------------------
// FUNGSI UTAMA / SETUP
//-------------------------------------------------------------

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
    "Nama_WaliKelas", // <-- Kolom ini akan diisi di Sheet HasilUjian jika ada data
    // Mata Pelajaran (sesuai urutan E-O)
    "Bahasa_Indonesia",
    "Matematika",
    "IPA",
    "IPS",
    "PAI",
    "PPKN",
    "PJOK",
    "Bahasa_Inggris",
    "Seni_Budaya",
    "Informatika",
    "Bahasa_Madura",
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

function doGet(e) {
  // ... (tidak ada perubahan signifikan pada doGet) ...
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

//-------------------------------------------------------------
// FUNGSI getStudentReport
//-------------------------------------------------------------

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
    (row) =>
      row[COL_INDEX_MAP.No_Peserta].toString().trim() ===
      noPeserta.toString().trim()
  );

  if (!studentData) {
    return { status: "not_found", message: "Data siswa tidak ditemukan." };
  }

  const report = {
    No_Peserta: studentData[COL_INDEX_MAP.No_Peserta],
    Nama_Siswa: studentData[COL_INDEX_MAP.Nama_Siswa],
    Kelas_Siswa: studentData[COL_INDEX_MAP.Kelas_Siswa],
    Nama_WaliKelas: studentData[COL_INDEX_MAP.Nama_WaliKelas], // Ambil dari kolom D
  };

  // Loop untuk nilai mata pelajaran (index 4 sampai 14)
  for (const subjectKey in SUBJECT_MAP) {
    const index = COL_INDEX_MAP[subjectKey];
    if (index !== undefined) {
      const rawValue = studentData[index];
      const score =
        typeof rawValue === "number"
          ? rawValue
          : rawValue
          ? parseFloat(rawValue)
          : null;
      report[subjectKey] = score !== null && !isNaN(score) ? score : "";
    }
  }

  return { status: "success", data: report };
}

//-------------------------------------------------------------
// FUNGSI getRecapReport (DIPERBAIKI)
//-------------------------------------------------------------
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

  // 🟢 PERBAIKAN 1: Ambil map Wali Kelas dari sheet DataBase_Utama
  const waliKelasMap = getWaliKelasMap();
  const kelasUpper = kelas.toString().trim().toUpperCase();

  // 🟢 PERBAIKAN 2: Ambil Nama Wali Kelas berdasarkan Kelas yang difilter
  // Ini akan digunakan untuk TTD Rekap
  const namaWaliKelas =
    waliKelasMap[kelasUpper] || "Nama Wali Kelas Belum Ditetapkan";

  // Tentukan index kolom mata pelajaran
  const mapelColIndex = COL_INDEX_MAP[mapel];
  if (mapelColIndex === undefined) {
    return { status: "error", message: "Kunci mata pelajaran tidak valid." };
  }

  const recapData = [];

  // Kolom Kelas berada di index 2 (Kolom C)
  data.forEach((row) => {
    if (
      row[COL_INDEX_MAP.Kelas_Siswa] &&
      row[COL_INDEX_MAP.Kelas_Siswa].toString().trim().toUpperCase() ===
        kelasUpper
    ) {
      const rawValue = row[mapelColIndex];
      const score =
        typeof rawValue === "number"
          ? rawValue
          : rawValue
          ? parseFloat(rawValue)
          : null;

      recapData.push({
        No_Peserta: row[COL_INDEX_MAP.No_Peserta],
        Nama_Siswa: row[COL_INDEX_MAP.Nama_Siswa],
        Kelas_Siswa: row[COL_INDEX_MAP.Kelas_Siswa],
        // 🟢 PERBAIKAN 3: Sisipkan Nama Wali Kelas dari map
        Nama_WaliKelas: namaWaliKelas,
        [mapel]: score !== null && !isNaN(score) ? score : "",
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

function createJsonOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON
  );
}
