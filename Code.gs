/**
 * =========================================================================
 * SIREKAP v16 - SIMPLIFIED BACKEND ENGINE (3 SHEETS ARCHITECTURE)
 * SDIT QURRATU A'YUN AL-ISLAMI KABUPATEN MAROS
 * =========================================================================
 * Hanya menggunakan 3 Sheet:
 * 1. "KBM"    : Seluruh riwayat presensi harian KBM
 * 2. "Ujian"  : Seluruh riwayat presensi kegiatan ujian semester
 * 3. "Master" : Master Data Guru, PIN Keamanan, Mapel, Kelas, & Password Admin
 * =========================================================================
 */

const SHEET_KBM = "KBM";
const SHEET_UJIAN = "Ujian";
const SHEET_MASTER = "Master";
const DEFAULT_ADMIN_PASS = "admin123";

/**
 * Endpoint GET (Membaca data)
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const adminPass = e.parameter.admin_pass;
    const guru = e.parameter.guru;
    const pin = e.parameter.pin;
    const type = e.parameter.type;

    // 1. Ambil Master Data (Guru, Mapel, Kelas)
    if (action === "get_master") {
      return responseJSON(getMasterData());
    }

    // 2. Baca Data KBM (Harian) & Ujian
    if (action === "read") {
      // Jika diakses oleh Guru (Riwayat Pribadi & Akumulasi JP)
      if (guru && pin) {
        if (!validasiGuru(guru, pin)) {
          return responseJSON({ result: "error", message: "Akses Ditolak. PIN Keamanan salah." });
        }
        if (type === "ujian") {
          return responseJSON(getRiwayatGuru(SHEET_UJIAN, guru));
        } else {
          return responseJSON(getRiwayatGuru(SHEET_KBM, guru));
        }
      }

      // Jika diakses oleh Admin / Kepala Sekolah (Dashboard KBM atau Ujian)
      if (validasiAdmin(adminPass)) {
        if (type === "ujian") {
          return responseJSON(getAllDataUjian());
        } else {
          return responseJSON(getAllDataKBM());
        }
      }
      return responseJSON({ result: "error", message: "Akses Ditolak. Password Admin salah." });
    }

    // 3. Baca Data Ujian (Dukungan endpoint terpisah read_ujian)
    if (action === "read_ujian") {
      // Jika diakses oleh Guru
      if (guru && pin) {
        if (!validasiGuru(guru, pin)) {
          return responseJSON({ result: "error", message: "Akses Ditolak. PIN Keamanan salah." });
        }
        return responseJSON(getRiwayatGuru(SHEET_UJIAN, guru));
      }

      // Jika diakses oleh Admin
      if (validasiAdmin(adminPass)) {
        return responseJSON(getAllDataUjian());
      }
      return responseJSON({ result: "error", message: "Akses Ditolak. Password Admin salah." });
    }

    // 4. Test Koneksi
    if (action === "ping") {
      return responseJSON({ result: "success", message: "SiRekap Backend v16 Online 🟢", timestamp: new Date().toISOString() });
    }

    // 5. Setup Otomatis Slip Gaji (Bisa diakses langsung lewat link / browser)
    if (action === "setup_slip_gaji") {
      if (!validasiAdmin(adminPass)) {
        return responseJSON({ result: "error", message: "Akses Ditolak. Password Admin salah." });
      }
      setupSlipGaji();
      return responseJSON({ result: "success", message: "Sheet Slip_Gaji & Master_Gaji Berhasil Dibuat! Periksa spreadsheet Anda." });
    }

    // 6. Sinkronkan & Perbaiki Rumus Slip Gaji Otomatis
    if (action === "refresh_slip" || action === "sync_slip") {
      sinkronkanMasterGaji();
      try {
        const sMaster = getSheet(SHEET_MASTER);
        const curPass = String(sMaster.getRange("I2").getValue() || "").trim();
        if (!curPass) {
          sMaster.getRange("I2").setValue(DEFAULT_ADMIN_PASS);
        }
      } catch (e) {}
      return responseJSON({ result: "success", message: "Rumus Slip Gaji Berhasil Diperbarui & Disinkronkan!" });
    }

    // 7. Reset / Set Password Admin Master ke Default (admin123)
    if (action === "reset_admin_pass") {
      try {
        const sMaster = getSheet(SHEET_MASTER);
        sMaster.getRange("I2").setValue(DEFAULT_ADMIN_PASS);
        return responseJSON({ result: "success", message: "Password Admin di Sheet Master (I2) berhasil diset ke: " + DEFAULT_ADMIN_PASS });
      } catch (err) {
        return responseJSON({ result: "error", message: err.toString() });
      }
    }

    return responseJSON({ result: "error", message: "Aksi GET tidak dikenali: " + action });
  } catch (err) {
    return responseJSON({ result: "error", message: err.toString() });
  }
}

/**
 * Endpoint POST (Menyimpan, mengubah, menghapus data)
 */
function doPost(e) {
  try {
    const p = e.parameter;
    const action = p.action;

    // 1. Verifikasi Login Guru
    if (action === "login_guru") {
      const ok = validasiGuru(p.guru, p.pin);
      if (ok) {
        return responseJSON({ result: "success", message: "Login Berhasil" });
      } else {
        return responseJSON({ result: "error", message: "Nama Guru atau PIN Keamanan Salah!" });
      }
    }

    // 2. Simpan Master Data (dari Admin)
    if (action === "save_master") {
      if (!validasiAdmin(p.admin_pass)) {
        return responseJSON({ result: "error", message: "Akses Ditolak. Password Admin salah." });
      }
      saveMasterData(p.teachers, p.subjects, p.classes);
      return responseJSON({ result: "success", message: "Master Data berhasil diperbarui." });
    }

    // 3. Tambah Presensi KBM (Harian)
    if (action === "insert") {
      return responseJSON(insertKBM(p));
    }

    // 4. Tambah Presensi Ujian
    if (action === "insert_ujian") {
      return responseJSON(insertUjian(p));
    }

    // 5. Update Presensi (KBM atau Ujian)
    if (action === "update") {
      return responseJSON(updateData(p));
    }

    // 6. Hapus Presensi (KBM atau Ujian)
    if (action === "delete") {
      return responseJSON(deleteData(p));
    }

    // 7. Sinkronisasi Otomatis dari Aplikasi Presensi RFID
    if (action === "sync_presensi_rfid") {
      return responseJSON(syncPresensiFromRFID(p));
    }

    return responseJSON({ result: "error", message: "Aksi POST tidak dikenali: " + action });
  } catch (err) {
    return responseJSON({ result: "error", message: err.toString() });
  }
}

/**
 * Menerima sinkronisasi data presensi langsung dari Aplikasi Presensi RFID
 * Menulis ke sheet Presensi, mengupdate Master_Gaji dan menghitung total honor di Slip_Gaji secara otomatis
 */
function syncPresensiFromRFID(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sPresensi = ss.getSheetByName("Presensi");
  if (!sPresensi) {
    sPresensi = ss.insertSheet("Presensi");
  }

  // 1. Cek apakah ada data baris sheet lengkap (format arsip harian: 10. Presensi Oktober)
  let customRows = [];
  try {
    if (p.sheet_rows) {
      customRows = JSON.parse(p.sheet_rows);
    }
  } catch(e) {}

  if (Array.isArray(customRows) && customRows.length > 0) {
    sPresensi.clear();
    const maxCols = Math.max(...customRows.map(r => (Array.isArray(r) ? r.length : 0)), 8);
    const normalizedRows = customRows.map(r => {
      const rowCopy = Array.isArray(r) ? [...r] : [];
      while (rowCopy.length < maxCols) rowCopy.push("");
      return rowCopy;
    });

    sPresensi.getRange(1, 1, normalizedRows.length, maxCols).setValues(normalizedRows);

    // Styling Otomatis Sheet Presensi Arsip:
    // a. Judul Banner (Baris 1 - 3)
    sPresensi.getRange("A1:H3").setFontWeight("bold").setHorizontalAlignment("left");
    sPresensi.getRange("A1").setFontSize(13);
    sPresensi.getRange("A2").setFontSize(11);
    sPresensi.getRange("A3").setFontSize(11);

    // b. Beri styling warna pada tabel guru
    for (let r = 0; r < normalizedRows.length; r++) {
      const c0 = String(normalizedRows[r][0] || "").trim();
      const c1 = String(normalizedRows[r][1] || "").trim().toUpperCase();

      // Baris Nama Guru di atas tabel
      if (c0.length >= 4 && !/^(presensi|kabupaten|bulan|id|no\.|hari|maros|kepala)/i.test(c0) && isNaN(c0) && c0.indexOf("31") === -1 && c0.indexOf("2026") === -1) {
        sPresensi.getRange(r + 1, 1, 1, maxCols).setFontWeight("bold").setFontSize(11).setBackground("#f1f5f9");
      }

      // Baris Header Tabel
      if (c0 === "ID" || c0 === "NO." || c1 === "HARI") {
        sPresensi.getRange(r + 1, 1, 1, maxCols).setBackground("#0284c7").setFontColor("#ffffff").setFontWeight("bold");
      }

      // Baris Keterangan Kehadiran
      const statusText = String(normalizedRows[r][5] || "").trim().toUpperCase();
      if (statusText === "TEPAT WAKTU") {
        sPresensi.getRange(r + 1, 6).setFontColor("#059669").setFontWeight("bold");
      } else if (statusText === "TERLAMBAT") {
        sPresensi.getRange(r + 1, 6).setFontColor("#d97706").setFontWeight("bold");
      } else if (statusText === "TIDAK HADIR") {
        sPresensi.getRange(r + 1, 6).setFontColor("#94a3b8");
      }

      // Format Total Kehadiran
      const totalCell = normalizedRows[r][6];
      if (totalCell !== "" && totalCell !== null && !isNaN(totalCell)) {
        sPresensi.getRange(r + 1, 7).setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
      }
    }

    sPresensi.autoResizeColumns(1, maxCols);

    // Tarik langsung kehadiran ke Master_Gaji & perbarui seluruh rumus Slip Gaji
    tarikKehadiranPresensi();

    return {
      result: "success",
      count: normalizedRows.length,
      message: "Berhasil menyinkronkan data presensi lengkap (format arsip harian) ke Sheet Presensi SiRekap!"
    };
  }

  // 2. Fallback: format tabel ringkasan 6 kolom
  let guruList = [];
  try {
    guruList = JSON.parse(p.guru_data || "[]");
  } catch(e) {
    return { result: "error", message: "Format JSON guru_data tidak valid." };
  }

  // Filter proteksi ganda: pastikan data murid/siswa tidak dimasukkan ke sheet Presensi SiRekap
  guruList = guruList.filter(g => {
    const peran = String(g.peran || "").toLowerCase().trim();
    const kelas = String(g.kelas || g.jabatan || "").toLowerCase().trim();
    if (peran === "murid" || peran === "siswa" || /^kelas\s*\d/i.test(kelas)) {
      return false;
    }
    return true;
  });

  if (guruList.length === 0) {
    return { result: "error", message: "Data presensi guru kosong atau seluruh data teridentifikasi sebagai murid." };
  }

  const periode = p.periode || ("Bulan " + Utilities.formatDate(new Date(), "Asia/Makassar", "MMMM yyyy"));

  // Tulis ulang / perbarui sheet Presensi dengan format yang bersih & rapi
  sPresensi.clear();
  const header = ["No", "Nama Guru / Pegawai", "Jumlah Kehadiran (Hari)", "Jabatan / Kelas", "Periode", "Waktu Sinkron"];
  sPresensi.getRange(1, 1, 1, header.length).setValues([header]);
  sPresensi.getRange(1, 1, 1, header.length).setBackground("#0284c7").setFontColor("#ffffff").setFontWeight("bold");
  sPresensi.setFrozenRows(1);

  const nowStr = Utilities.formatDate(new Date(), "Asia/Makassar", "yyyy-MM-dd HH:mm:ss");
  const rows = guruList.map((g, idx) => {
    return [
      idx + 1,
      g.nama || "",
      Number(g.hadir) || 0,
      g.kelas || g.jabatan || "Guru",
      periode,
      nowStr
    ];
  });

  sPresensi.getRange(2, 1, rows.length, header.length).setValues(rows);
  sPresensi.getRange(2, 3, rows.length, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");
  sPresensi.autoResizeColumns(1, header.length);

  // Tarik langsung kehadiran ke Master_Gaji & perbarui seluruh rumus Slip Gaji
  tarikKehadiranPresensi();

  return {
    result: "success",
    count: rows.length,
    message: "Berhasil menyinkronkan " + rows.length + " data kehadiran guru dari Presensi RFID ke SiRekap!"
  };
}

// =========================================================================
// OPERASI DATABASE SHEET
// =========================================================================

/**
 * Mengambil objek Sheet atau membuatnya secara otomatis jika belum ada
 */
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_KBM) {
      sheet.appendRow(["ID", "Tanggal", "Waktu", "Kode", "Nama", "Status", "Mapel", "Kelas", "Mode", "Jam Sendiri", "Jam Team"]);
      formatHeader(sheet, 11);
    } else if (name === SHEET_UJIAN) {
      sheet.appendRow(["ID", "Tanggal", "Kode", "Nama", "Kategori", "Rincian", "Volume", "Nominal"]);
      formatHeader(sheet, 8);
    } else if (name === SHEET_MASTER) {
      sheet.getRange("A1:I1").setValues([["Kode", "Nama Guru", "PIN", "", "Mapel", "", "Kelas", "", "Admin Pass"]]);
      formatHeader(sheet, 9);
      sheet.getRange("I2").setValue(DEFAULT_ADMIN_PASS);
    }
  }
  return sheet;
}

function formatHeader(sheet, cols) {
  const range = sheet.getRange(1, 1, 1, cols);
  range.setBackground("#4f46e5");
  range.setFontColor("#ffffff");
  range.setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatTanggal(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Makassar", "yyyy-MM-dd");
  }
  const str = String(val).trim();
  const match = str.match(/^(\d{4}-\d{2}-\d{2})/);
  if (match) return match[1];
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "Asia/Makassar", "yyyy-MM-dd");
  }
  return str;
}

function formatWaktu(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Makassar", "HH:mm");
  }
  return String(val).trim();
}

/**
 * Validasi password admin
 */
function validasiAdmin(pass) {
  if (!pass) return false;
  const cleanPass = String(pass).trim();

  // 1. Master Key Fallback: DEFAULT_ADMIN_PASS ("admin123") selalu valid tanpa syarat
  if (cleanPass === DEFAULT_ADMIN_PASS) return true;

  // 2. Cek Password di Sheet Master Cell I2
  try {
    const sheet = getSheet(SHEET_MASTER);
    const dbVal = sheet.getRange("I2").getValue();
    if (dbVal !== "" && dbVal !== null && dbVal !== undefined) {
      const dbPass = String(dbVal).trim();
      if (dbPass && cleanPass === dbPass) return true;
    }
  } catch (e) {}

  return false;
}

/**
 * Validasi login guru dengan Nama + PIN
 */
function validasiGuru(namaLengkap, pin) {
  if (!namaLengkap || !pin) return false;
  const sheet = getSheet(SHEET_MASTER);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const targetClean = String(namaLengkap).trim().toLowerCase();
  const pinClean = String(pin).trim();

  for (let i = 0; i < data.length; i++) {
    const rawKode = String(data[i][0]).trim();
    const rawNama = String(data[i][1]).trim();
    const dbPin = String(data[i][2]).trim();

    const formattedKode = rawKode ? (rawKode.length === 1 ? "0" + rawKode : rawKode) : "";
    const combined = (formattedKode ? formattedKode + " " : "") + rawNama;

    if (
      (rawNama.toLowerCase() === targetClean || combined.toLowerCase() === targetClean) &&
      dbPin === pinClean
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Ambil Master Data untuk dropdown form
 */
function getMasterData() {
  const sheet = getSheet(SHEET_MASTER);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { result: "success", teachers: [], subjects: [], classes: [] };

  const rawGuru = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const rawMapel = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  const rawKelas = sheet.getRange(2, 7, lastRow - 1, 1).getValues();

  const teachers = [];
  rawGuru.forEach(r => {
    if (r[1] && String(r[1]).trim() !== "") {
      const code = String(r[0]).trim();
      const formattedCode = code ? (code.length === 1 ? "0" + code : code) + " " : "";
      teachers.push(formattedCode + String(r[1]).trim());
    }
  });

  const subjects = rawMapel.map(r => String(r[0]).trim()).filter(v => v !== "");
  const classes = rawKelas.map(r => String(r[0]).trim()).filter(v => v !== "");

  // Ambil data tarif terbaru dari spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tarifUjian = getTarifUjianMap(ss);
  const tarifKBM = getTarifKBM(ss);
  const tarifTunjangan = getTarifTunjangan(ss);

  return {
    result: "success",
    teachers: teachers,
    subjects: subjects,
    classes: classes,
    tarif: {
      ujian: tarifUjian,
      kbm: tarifKBM,
      tunjangan: tarifTunjangan
    }
  };
}

/**
 * Simpan Master Data yang diedit dari menu Admin (tanpa menghilangkan PIN)
 */
function saveMasterData(teachersJSON, subjectsJSON, classesJSON) {
  const sheet = getSheet(SHEET_MASTER);
  const teachers = JSON.parse(teachersJSON || "[]");
  const subjects = JSON.parse(subjectsJSON || "[]");
  const classes = JSON.parse(classesJSON || "[]");

  // Ambil PIN lama agar tidak terhapus
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const oldGuruData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const pinMap = {};
  oldGuruData.forEach(r => {
    const key = String(r[1]).trim();
    if (key && r[2]) pinMap[key] = String(r[2]).trim();
  });

  // Bersihkan data lama (kolom A-G)
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }

  // Siapkan data guru baru
  const guruRows = teachers.map(t => {
    const match = String(t).trim().match(/^(\d+)[\.\-\s]+(.*)/);
    const code = match ? (match[1].length === 1 ? "0" + match[1] : match[1]) : "";
    const name = match ? match[2].trim() : t.trim();
    const pin = pinMap[name] || "123456"; // Default PIN 123456 jika guru baru
    return [code, name, pin];
  });

  if (guruRows.length > 0) sheet.getRange(2, 1, guruRows.length, 3).setValues(guruRows);
  if (subjects.length > 0) sheet.getRange(2, 5, subjects.length, 1).setValues(subjects.map(s => [s]));
  if (classes.length > 0) sheet.getRange(2, 7, classes.length, 1).setValues(classes.map(c => [c]));
}

/**
 * Membaca semua data KBM (Untuk Admin Dashboard)
 */
function getAllDataKBM() {
  const sheet = getSheet(SHEET_KBM);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { result: "success", data: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const list = values.map(r => {
    const rawCode = String(r[3]).trim();
    return {
      ID: String(r[0]),
      Tanggal: formatTanggal(r[1]),
      Waktu: formatWaktu(r[2]),
      Kode: rawCode ? (rawCode.length === 1 ? "0" + rawCode : rawCode) : "",
      Nama: String(r[4]).trim(),
      Status: String(r[5] || "Tetap"),
      Mapel: String(r[6] || ""),
      Kelas: String(r[7] || ""),
      Mode: String(r[8] || "sendiri"),
      "Jam Sendiri": Number(r[9]) || 0,
      "Jam Team": Number(r[10]) || 0
    };
  });

  list.reverse();
  return { result: "success", data: list.slice(0, 100) };
}

/**
 * Membaca semua data Ujian (Untuk Admin Dashboard)
 */
function getAllDataUjian() {
  const sheet = getSheet(SHEET_UJIAN);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { result: "success", data: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const list = values.map(r => {
    const rawCode = String(r[2]).trim();
    return {
      ID: String(r[0]),
      Tanggal: formatTanggal(r[1]),
      Kode: rawCode ? (rawCode.length === 1 ? "0" + rawCode : rawCode) : "",
      Nama: String(r[3]).trim(),
      Kategori: String(r[4] || ""),
      Rincian: String(r[5] || ""),
      Volume: String(r[6] || "1"),
      Nominal: Number(r[7]) || 0
    };
  });

  list.reverse();
  return { result: "success", data: list.slice(0, 100) };
}

/**
 * Membaca riwayat & statistik spesifik seorang Guru
 */
function getRiwayatGuru(sheetName, guruString) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { result: "success", data: [], stats: { sendiri: 0, team: 0 } };

  const match = String(guruString).trim().match(/^(\d+)[\.\-\s]+(.*)/);
  const targetCode = match ? (match[1].length === 1 ? "0" + match[1] : match[1]) : "";
  const targetName = (match ? match[2] : guruString).trim().toLowerCase();
  const normTarget = normalizeTeacherName(targetName);

  const isKBM = (sheetName === SHEET_KBM);
  const numCols = isKBM ? 11 : 8;
  const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const filtered = [];
  let sumSendiri = 0;
  let sumTeam = 0;

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const rawCode = String(r[isKBM ? 3 : 2]).trim();
    const rowCode = rawCode ? (rawCode.length === 1 ? "0" + rawCode : rawCode) : "";
    const rowName = String(r[isKBM ? 4 : 3]).trim().toLowerCase();
    const normRow = normalizeTeacherName(rowName);

    const isMatch = (targetCode && rowCode && targetCode === rowCode) || 
                    (rowName && rowName === targetName) ||
                    (targetName && rowName && (rowName.indexOf(targetName) !== -1 || targetName.indexOf(rowName) !== -1)) ||
                    (normTarget && normRow && (normRow === normTarget || normRow.indexOf(normTarget) !== -1 || normTarget.indexOf(normRow) !== -1));

    if (isMatch) {
      if (isKBM) {
        const js = Number(r[9]) || 0;
        const jt = Number(r[10]) || 0;
        sumSendiri += js;
        sumTeam += jt;
        filtered.push({
          ID: String(r[0]),
          Tanggal: formatTanggal(r[1]),
          Waktu: formatWaktu(r[2]),
          Kode: rowCode,
          Nama: String(r[4]).trim(),
          Status: String(r[5] || "Tetap"),
          Mapel: String(r[6] || ""),
          Kelas: String(r[7] || ""),
          Mode: String(r[8] || "sendiri"),
          "Jam Sendiri": js,
          "Jam Team": jt
        });
      } else {
        filtered.push({
          ID: String(r[0]),
          Tanggal: formatTanggal(r[1]),
          Kode: rowCode,
          Nama: String(r[3]).trim(),
          Kategori: String(r[4] || ""),
          Rincian: String(r[5] || ""),
          Volume: String(r[6] || "1"),
          Nominal: Number(r[7]) || 0
        });
      }
    }
  }

  filtered.reverse();
  return {
    result: "success",
    data: filtered.slice(0, 50),
    stats: { sendiri: sumSendiri, team: sumTeam }
  };
}

/**
 * Tambah baris baru KBM
 */
function insertKBM(p) {
  const sheet = getSheet(SHEET_KBM);
  sheet.appendRow([
    p.ID || String(Date.now()),
    p.Tanggal,
    p.Waktu || "",
    p.KodeGuru || "",
    p.NamaGuru || p.Guru || "",
    p.Status || "Tetap",
    p.Mapel || "",
    p.Kelas || "",
    p.Mode || "sendiri",
    Number(p["Jam Sendiri"]) || 0,
    Number(p["Jam Team"]) || 0
  ]);
  return { result: "success", message: "Data KBM tersimpan." };
}

// =========================================================================
// PENGATURAN TARIF DINAMIS (UJIAN, KBM, TUNJANGAN)
// =========================================================================

/**
 * Mencari sheet Tarif di Spreadsheet (mencakup alias 'Tarif', 'Tarif_Ujian', 'Pengaturan', atau sheet dengan header 'KATEGORI UJIAN')
 */
function findSheetTarif(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const names = ["Tarif", "Tarif_Ujian", "Tarif Ujian", "Pengaturan"];
  for (const n of names) {
    const s = ss.getSheetByName(n);
    if (s) return s;
  }

  const allSheets = ss.getSheets();
  for (const s of allSheets) {
    try {
      const a1 = String(s.getRange("A1").getValue() || "").trim().toUpperCase();
      if (a1 === "KATEGORI UJIAN" || a1.indexOf("KATEGORI UJIAN") !== -1) {
        return s;
      }
    } catch(e) {}
  }

  return null;
}

/**
 * Menyiapkan sheet 'Tarif' sebagai tabel kontrol harga/tarif lengkap di spreadsheet
 * Sesuai gambar:
 * - Kolom A & B: Tarif Kegiatan Ujian
 * - Kolom D & E: Tarif KBM (Jam Sendiri & Jam Team)
 * - Kolom G & H: Tarif Tunjangan (Transport, Wali Kelas, Kesehatan)
 */
function setupSheetTarif(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sTarif = findSheetTarif(ss);
  
  if (!sTarif) {
    sTarif = ss.insertSheet("Tarif");
  }

  // 1. Tabel Tarif Kegiatan Ujian (Kolom A & B)
  const a1 = String(sTarif.getRange("A1").getValue() || "").trim().toUpperCase();
  if (a1 !== "KATEGORI UJIAN" && a1.indexOf("KATEGORI UJIAN") === -1) {
    sTarif.getRange("A1:B1").setValues([["KATEGORI UJIAN", "TARIF (Rp)"]]);
    sTarif.getRange("A1:B1").setBackground("#dcfce7").setFontColor("#065f46").setFontWeight("bold");
    
    const rowsUjian = [
      ["mengawas", 10000],
      ["tulis_soal", 20000],
      ["periksa", 5000],
      ["praktik", 10000],
      ["rapor", 30000]
    ];
    sTarif.getRange(2, 1, rowsUjian.length, 2).setValues(rowsUjian);
    sTarif.getRange(2, 2, rowsUjian.length, 1).setNumberFormat('"Rp"#,##0');
  }

  // 2. Tabel Tarif KBM (Kolom D & E)
  const d1 = String(sTarif.getRange("D1").getValue() || "").trim().toUpperCase();
  if (d1 !== "JENIS KBM" && d1.indexOf("KBM") === -1) {
    sTarif.getRange("D1:E1").setValues([["JENIS KBM", "TARIF PER JAM (Rp)"]]);
    sTarif.getRange("D1:E1").setBackground("#e0e7ff").setFontColor("#3730a3").setFontWeight("bold");
    const rowsKBM = [
      ["Jam Sendiri (KBM 1)", 12000],
      ["Jam Team (KBM 2)", 10000]
    ];
    sTarif.getRange(2, 4, rowsKBM.length, 2).setValues(rowsKBM);
    sTarif.getRange(2, 5, rowsKBM.length, 1).setNumberFormat('"Rp"#,##0');
  }

  // 3. Tabel Tarif Tunjangan (Kolom G & H)
  const g1 = String(sTarif.getRange("G1").getValue() || "").trim().toUpperCase();
  if (g1 !== "JENIS TUNJANGAN" && g1.indexOf("TUNJANGAN") === -1) {
    sTarif.getRange("G1:H1").setValues([["JENIS TUNJANGAN", "NOMINAL (Rp)"]]);
    sTarif.getRange("G1:H1").setBackground("#fef3c7").setFontColor("#92400e").setFontWeight("bold");
    const rowsTunjangan = [
      ["Transport per Kehadiran (Hari)", 7500],
      ["Tunjangan Wali Kelas", 100000],
      ["Tunjangan Kesehatan", 50000]
    ];
    sTarif.getRange(2, 7, rowsTunjangan.length, 2).setValues(rowsTunjangan);
    sTarif.getRange(2, 8, rowsTunjangan.length, 1).setNumberFormat('"Rp"#,##0');
  }

  sTarif.setColumnWidth(1, 140);
  sTarif.setColumnWidth(2, 110);
  sTarif.setColumnWidth(3, 30);
  sTarif.setColumnWidth(4, 180);
  sTarif.setColumnWidth(5, 150);
  sTarif.setColumnWidth(6, 30);
  sTarif.setColumnWidth(7, 220);
  sTarif.setColumnWidth(8, 130);

  return sTarif;
}

/**
 * Membaca peta tarif ujian dari sheet Tarif di Spreadsheet
 */
function getTarifUjianMap(ss) {
  const defaultTarif = {
    "mengawas": 10000,
    "tulis_soal": 20000,
    "periksa": 5000,
    "praktik": 10000,
    "rapor": 30000
  };

  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = findSheetTarif(ss);
    if (!sheet || sheet.getLastRow() < 2) return defaultTarif;

    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const tarifMap = { ...defaultTarif };

    values.forEach(r => {
      const rawCat = String(r[0] || "").trim().toLowerCase();
      if (!rawCat) return;

      let num = 0;
      if (typeof r[1] === "number") {
        num = r[1];
      } else {
        const cleanStr = String(r[1] || "").replace(/[^0-9]/g, "");
        num = cleanStr ? Number(cleanStr) : 0;
      }

      if (num > 0) {
        const key = rawCat.replace(/[\s\-]+/g, "_");
        tarifMap[key] = num;

        if (key.indexOf("mengawas") !== -1 || key.indexOf("awas") !== -1) tarifMap["mengawas"] = num;
        if (key.indexOf("tulis") !== -1 || key.indexOf("soal") !== -1 || key.indexOf("menulis") !== -1) tarifMap["tulis_soal"] = num;
        if (key.indexOf("periksa") !== -1) tarifMap["periksa"] = num;
        if (key.indexOf("praktik") !== -1 || key.indexOf("praktek") !== -1) tarifMap["praktik"] = num;
        if (key.indexOf("rapor") !== -1 || key.indexOf("raport") !== -1) tarifMap["rapor"] = num;
      }
    });

    return tarifMap;
  } catch(e) {
    return defaultTarif;
  }
}

/**
 * Membaca tarif KBM dari sheet Tarif (Jam Sendiri & Jam Team)
 */
function getTarifKBM(ss) {
  const defaultKBM = { sendiri: 12000, team: 10000 };
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = findSheetTarif(ss);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 5) return defaultKBM;

    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(2, 4, Math.min(lastRow - 1, 10), 2).getValues();
    const res = { ...defaultKBM };

    values.forEach(r => {
      const name = String(r[0] || "").toLowerCase();
      let num = typeof r[1] === "number" ? r[1] : Number(String(r[1] || "").replace(/[^0-9]/g, "")) || 0;
      if (num > 0) {
        if (name.indexOf("sendiri") !== -1 || name.indexOf("kbm 1") !== -1 || name.indexOf("kbm1") !== -1) {
          res.sendiri = num;
        } else if (name.indexOf("team") !== -1 || name.indexOf("kbm 2") !== -1 || name.indexOf("kbm2") !== -1) {
          res.team = num;
        }
      }
    });

    return res;
  } catch(e) {
    return defaultKBM;
  }
}

/**
 * Membaca tarif tunjangan (transport, wali kelas, kesehatan) dari sheet Tarif
 */
function getTarifTunjangan(ss) {
  const defaultTunjangan = { transport: 7500, wali: 100000, kesehatan: 50000 };
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = findSheetTarif(ss);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 8) return defaultTunjangan;

    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(2, 7, Math.min(lastRow - 1, 10), 2).getValues();
    const res = { ...defaultTunjangan };

    values.forEach(r => {
      const name = String(r[0] || "").toLowerCase();
      let num = typeof r[1] === "number" ? r[1] : Number(String(r[1] || "").replace(/[^0-9]/g, "")) || 0;
      if (num > 0) {
        if (name.indexOf("transport") !== -1 || name.indexOf("kehadiran") !== -1) {
          res.transport = num;
        } else if (name.indexOf("wali") !== -1) {
          res.wali = num;
        } else if (name.indexOf("kesehatan") !== -1) {
          res.kesehatan = num;
        }
      }
    });

    return res;
  } catch(e) {
    return defaultTunjangan;
  }
}

/**
 * Menghitung nominal honor kegiatan Ujian secara otomatis
 * berdasarkan jenis kegiatan, volume, tarif di sheet 'Tarif', dan status potongan 25% di Master_Gaji
 */
function hitungNominalUjian(kategori, volume, kodeGuru, namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tarifMap = getTarifUjianMap(ss);

  const rawCat = String(kategori || "").trim().toLowerCase().replace(/[\s\-]+/g, "_");
  let tarif = 0;

  // 1. Coba pencocokan langsung
  if (tarifMap[rawCat] !== undefined) {
    tarif = tarifMap[rawCat];
  } else {
    // 2. Pencocokan cerdas berdasarkan kata kunci
    for (const k in tarifMap) {
      if (rawCat.indexOf(k) !== -1 || k.indexOf(rawCat) !== -1) {
        tarif = tarifMap[k];
        break;
      }
    }
    // 3. Fallback variasi kata umum bahasa Indonesia
    if (tarif === 0) {
      if (rawCat.indexOf("tulis") !== -1 || rawCat.indexOf("soal") !== -1 || rawCat.indexOf("menulis") !== -1) {
        tarif = tarifMap["tulis_soal"] || 20000;
      } else if (rawCat.indexOf("awas") !== -1 || rawCat.indexOf("mengawas") !== -1) {
        tarif = tarifMap["mengawas"] || 10000;
      } else if (rawCat.indexOf("periksa") !== -1) {
        tarif = tarifMap["periksa"] || 5000;
      } else if (rawCat.indexOf("praktik") !== -1 || rawCat.indexOf("praktek") !== -1) {
        tarif = tarifMap["praktik"] || 10000;
      } else if (rawCat.indexOf("rapor") !== -1 || rawCat.indexOf("raport") !== -1) {
        tarif = tarifMap["rapor"] || 30000;
      }
    }
  }

  const vol = Number(volume) || 1;
  const subtotal = tarif * vol;

  // Cek apakah guru ini dikenakan potongan 25% dari sheet Master_Gaji
  let isDipotong = false;
  try {
    const sMasterGaji = ss.getSheetByName("Master_Gaji");
    if (sMasterGaji && sMasterGaji.getLastRow() >= 2) {
      const lastRow = sMasterGaji.getLastRow();
      const numCols = Math.max(sMasterGaji.getLastColumn(), 17);
      const dataGaji = sMasterGaji.getRange(2, 1, lastRow - 1, numCols).getValues();
      const targetCode = String(kodeGuru || "").trim();
      const formattedCode = targetCode.length === 1 ? "0" + targetCode : targetCode;
      const targetName = String(namaGuru || "").trim().toLowerCase();

      for (let i = 0; i < dataGaji.length; i++) {
        const rowCode = String(dataGaji[i][0]).trim();
        const rowFormattedCode = rowCode.length === 1 ? "0" + rowCode : rowCode;
        const rowName = String(dataGaji[i][1]).trim().toLowerCase();

        const match = (formattedCode && rowFormattedCode === formattedCode) ||
                      (targetName && (rowName === targetName || rowName.indexOf(targetName) !== -1 || targetName.indexOf(rowName) !== -1));

        if (match) {
          const potongVal = String(dataGaji[i][16] || "").trim().toLowerCase();
          if (potongVal === "ya" || potongVal === "true" || potongVal === "1" || potongVal === "y") {
            isDipotong = true;
          }
          break;
        }
      }
    }
  } catch (e) {
    Logger.log("Error hitungNominalUjian: " + e.toString());
  }

  if (isDipotong) {
    return Math.round(subtotal * 0.75); // Potong 25% (hanya dapat 75%)
  }
  return subtotal;
}

/**
 * Tambah baris baru Ujian (dengan nominal otomatis)
 */
function insertUjian(p) {
  const sheet = getSheet(SHEET_UJIAN);
  const nominal = (Number(p.Nominal) > 0) 
    ? Number(p.Nominal) 
    : hitungNominalUjian(p.Kategori, p.Volume, p.KodeGuru, p.NamaGuru || p.Guru);

  sheet.appendRow([
    p.ID || String(Date.now()),
    p.Tanggal,
    p.KodeGuru || "",
    p.NamaGuru || p.Guru || "",
    p.Kategori || "",
    p.Rincian || "",
    p.Volume || 1,
    nominal
  ]);
  return { result: "success", message: "Data Ujian tersimpan (Nominal: Rp " + nominal.toLocaleString("id-ID") + ")." };
}

/**
 * Update data berdasarkan ID unik
 */
function updateData(p) {
  const isKBM = (p.type === "harian");
  const sheet = getSheet(isKBM ? SHEET_KBM : SHEET_UJIAN);
  const data = sheet.getDataRange().getValues();
  const targetId = String(p.ID).trim();

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0]).trim();
    if (rowId === targetId || (Number(rowId) > 0 && Number(rowId) === Number(targetId))) {
      const row = i + 1;
      if (isKBM) {
        sheet.getRange(row, 2, 1, 10).setValues([[
          p.Tanggal,
          p.Waktu || data[i][2],
          p.KodeGuru || data[i][3],
          p.NamaGuru || p.Guru || data[i][4],
          p.Status || "Tetap",
          p.Mapel,
          p.Kelas,
          p.Mode,
          Number(p["Jam Sendiri"]) || 0,
          Number(p["Jam Team"]) || 0
        ]]);
      } else {
        const nominal = (Number(p.Nominal) > 0)
          ? Number(p.Nominal)
          : hitungNominalUjian(p.Kategori, p.Volume, p.KodeGuru || data[i][2], p.NamaGuru || p.Guru || data[i][3]);

        sheet.getRange(row, 2, 1, 7).setValues([[
          p.Tanggal,
          p.KodeGuru || data[i][2],
          p.NamaGuru || p.Guru || data[i][3],
          p.Kategori,
          p.Rincian,
          p.Volume,
          nominal
        ]]);
      }
      return { result: "success", message: "Data berhasil diperbarui." };
    }
  }
  return { result: "error", message: "ID data tidak ditemukan untuk diedit." };
}

/**
 * Hapus data berdasarkan ID unik
 */
function deleteData(p) {
  const isKBM = (p.type === "harian");
  const sheet = getSheet(isKBM ? SHEET_KBM : SHEET_UJIAN);
  const data = sheet.getDataRange().getValues();
  const targetId = String(p.ID).trim();

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0]).trim();
    if (rowId === targetId || (Number(rowId) > 0 && Number(rowId) === Number(targetId))) {
      sheet.deleteRow(i + 1);
      return { result: "success", message: "Data berhasil dihapus." };
    }
  }
  return { result: "error", message: "ID data tidak ditemukan untuk dihapus." };
}

// =========================================================================
// SETUP OTOMATIS: JALANKAN INI SATU KALI DI SPREADSHEET BARU
// =========================================================================

/**
 * Klik tombol RUN pada fungsi ini di Google Apps Script untuk otomatis
 * menyiapkan 3 Sheet lengkap dengan data Guru, Mapel, Kelas SDIT Qurratu A'yun!
 */
function initSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Buat Sheet KBM
  const sKBM = getSheet(SHEET_KBM);
  
  // 2. Buat Sheet Ujian
  const sUjian = getSheet(SHEET_UJIAN);
  
  // 3. Buat Sheet Master
  const sMaster = getSheet(SHEET_MASTER);
  
  // Inisialisasi Data Awal Master Guru
  const defaultTeachers = [
    ["01", "Radiant Fadli, S.Pd.", "123456"],
    ["02", "Bahtiar, S.Pd.I.", "123456"],
    ["03", "Agung Setiabudi, S.Pd., Gr.", "123456"],
    ["04", "Fandhy Ahmad Klana, S.S., Gr.", "123456"],
    ["05", "Muhammad Ihsan", "123456"],
    ["06", "Turmisi", "123456"],
    ["07", "Wahyu Saputra", "123456"],
    ["08", "Syamsir, S.Pd., Gr.", "123456"],
    ["09", "Ahmad Prisatrialdi, S.Pd., Gr.", "123456"],
    ["10", "Syahrul Ramadhan, S.H.", "123456"],
    ["11", "Muh. Sofyan Asnur, S.Pd., Gr.", "123456"],
    ["12", "Muhammad Fikri Haekal", "123456"],
    ["13", "Munadil Amri", "123456"],
    ["14", "Fuad", "123456"],
    ["16", "Dra. Mardia Habe", "123456"],
    ["17", "Hj. Nur Apni, S.Pd., Gr.", "123456"],
    ["18", "Hj. Amalia Ahsan, S.Pd., Gr.", "123456"],
    ["19", "Dewi Syahara Nur, S.Pd., Gr.", "123456"],
    ["20", "Ina Mulyati Salam, S.T.", "123456"],
    ["21", "Humairah Nafisah, S.Pd., Gr.", "123456"],
    ["22", "Fitriani, S.Pd.", "123456"],
    ["23", "Satriani", "123456"],
    ["24", "Nur Wahidah", "123456"],
    ["25", "Muti'ah S.", "123456"],
    ["26", "Liza Fany Adella, S.Pd., Gr.", "123456"],
    ["27", "Darmawati, S.Pd., Gr.", "123456"],
    ["28", "Ilma Nurtami, S.Pd., Gr.", "123456"],
    ["29", "A. Nurdiyawanti, S.Pd., M.Pd., Gr.", "123456"],
    ["30", "Rahmawati Basri, S.Pd., Gr.", "123456"],
    ["31", "Sri Wahyuni, S.Pd., Gr.", "123456"],
    ["32", "Ihwana, S.Pd., Gr.", "123456"],
    ["33", "Surianti", "123456"],
    ["34", "Hajriani", "123456"],
    ["35", "Nur Rizky", "123456"],
    ["38", "Hebrianti, S.Pd., M.Si.", "123456"],
    ["39", "Rahmawati", "123456"],
    ["40", "Dhea Rahmiyanti, S.Pd.", "123456"]
  ];
  
  const defaultSubjects = [
    ["Al-Qur'an"], ["Aqidah Akhlak"], ["Bahasa Arab"], ["Bahasa Bugis"], ["Bahasa Indonesia"],
    ["Bahasa Inggris"], ["Calistung"], ["Doa"], ["Fiqhi"], ["Hadis"], ["Hafalan Al-Qur'an"],
    ["IPAS"], ["Ilmu Tajwid"], ["Matematika"], ["PJOK"], ["Pend. Pancasila"], ["Praktik Salat"],
    ["Seni"], ["Sirah SKI"]
  ];
  
  const defaultClasses = [
    ["1A Putra"], ["1A Putri"], ["1B Putra"], ["1B Putri"],
    ["2A Putra"], ["2A Putri"], ["2B Putra"], ["2B Putri"],
    ["3A Putra"], ["3A Putri"], ["3B Putra"], ["3B Putri"],
    ["4 Putri"], ["4A Putra"], ["4B Putra"],
    ["5 Putri"], ["5A Putra"], ["5B Putri"],
    ["6 Putra"], ["6 Putri"]
  ];
  
  sMaster.getRange(2, 1, defaultTeachers.length, 3).setValues(defaultTeachers);
  sMaster.getRange(2, 5, defaultSubjects.length, 1).setValues(defaultSubjects);
  sMaster.getRange(2, 7, defaultClasses.length, 1).setValues(defaultClasses);
  sMaster.getRange("I2").setValue(DEFAULT_ADMIN_PASS);

  // Buat Sheet Tarif, Slip Gaji & Master Gaji otomatis
  setupSheetTarif(ss);
  setupSlipGaji();

  Logger.log("✅ Inisialisasi Spreadsheet Berhasil! Sheet KBM, Ujian, Master, Master_Gaji, Slip_Gaji, dan Tarif siap digunakan.");
}

/**
 * =========================================================================
 * SETUP SHEET SLIP GAJI & MASTER GAJI (OTOMATIS 1 KLIK)
 * =========================================================================
 * Membuat format slip gaji siap cetak sesuai template resmi SDIT Qurratu A'yun.
 * Menggunakan pemisah titik koma (;) sesuai standar Google Sheets regional Indonesia.
 */
function setupSlipGaji() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Pastikan sheet Tarif tersedia
  setupSheetTarif(ss);

  // 1. Siapkan Sheet Master_Gaji
  let sMasterGaji = ss.getSheetByName("Master_Gaji");
  if (!sMasterGaji) {
    sMasterGaji = ss.insertSheet("Master_Gaji");
  }
  
  const headerGaji = [
    "Kode", "Nama Guru", "Jabatan", "Alamat", "No HP",
    "Gaji Pokok", "Tunjangan Jabatan", "Tunjangan Wali Kelas",
    "Tunjangan Konsumsi Transport", "Tunjangan Kesehatan",
    "Pinjaman Kas", "Iuran SPP", "Donasi", "Lainnya",
    "TMT (Tgl Masuk)", "Sertifikasi", "Potongan Ujian 25%", "Kehadiran Transport (Hari)"
  ];
  
  // Ambil data guru dari Master
  const sMaster = getSheet(SHEET_MASTER);
  const lastMasterRow = sMaster.getLastRow();
  let teacherList = [];
  if (lastMasterRow >= 2) {
    teacherList = sMaster.getRange(2, 1, lastMasterRow - 1, 2).getValues();
  }
  
  // Simpan data lama di Master_Gaji agar data yang sudah diisi admin tidak hilang
  const existingGaji = {};
  const existingGajiByName = {};
  const existingGajiByNameClean = {};
  const curLastRow = sMasterGaji.getLastRow();
  const curLastCol = sMasterGaji.getLastColumn();
  if (curLastRow >= 2 && curLastCol > 0) {
    const oldVals = sMasterGaji.getRange(2, 1, curLastRow - 1, curLastCol).getValues();
    oldVals.forEach(r => {
      const k = String(r[0]).trim();
      const n = String(r[1]).trim().toLowerCase();
      const nClean = n.replace(/[^a-z0-9]/g, "");
      if (k) existingGaji[k] = r;
      if (n) existingGajiByName[n] = r;
      if (nClean) existingGajiByNameClean[nClean] = r;
    });
  }

  const sPresensi = ss.getSheetByName("Presensi");
  const attendanceMap = sPresensi ? parseSheetPresensi(sPresensi) : {};

  // Bangun daftar baris Master_Gaji lengkap (18 Kolom)
  const gajiRows = [];
  teacherList.forEach(t => {
    const rawCode = String(t[0]).trim();
    if (!rawCode && !t[1]) return;
    const formattedCode = rawCode.length === 1 ? "0" + rawCode : rawCode;
    const name = String(t[1]).trim();
    const nameLower = name.toLowerCase();
    const cleanName = nameLower.replace(/[^a-z0-9]/g, "");
    const presensiHari = cariKehadiranGuru(name, attendanceMap);
    
    let rowData = [];
    // Prioritaskan pencocokan berdasarkan NAMA GURU agar jabatan, alamat, TMT, dll tidak tertukar saat nomor kode diurutkan
    const exist = existingGajiByName[nameLower] || existingGajiByNameClean[cleanName] || existingGaji[formattedCode] || existingGaji[rawCode];
    if (exist) {
      rowData = exist.slice(0, headerGaji.length);
      while (rowData.length < headerGaji.length) {
        if (rowData.length === 14) rowData.push("");        // Kolom O: TMT
        else if (rowData.length === 15) rowData.push("Tidak"); // Kolom P: Sertifikasi
        else if (rowData.length === 16) rowData.push("Tidak"); // Kolom Q: Potongan Ujian 25%
        else if (rowData.length === 17) {
          // Kolom R: Kehadiran Transport (Hari)
          let defaultHari = 0;
          if (presensiHari !== null) {
            defaultHari = presensiHari;
          } else if (exist[8] !== undefined && Number(exist[8]) > 0) {
            defaultHari = Number(exist[8]) <= 31 ? Number(exist[8]) : Math.round(Number(exist[8]) / 7500);
          }
          rowData.push(defaultHari);
        }
        else rowData.push(0);
      }
      rowData[0] = formattedCode;
      rowData[1] = name;
      if (presensiHari !== null) {
        rowData[17] = presensiHari;
      }
      // Pastikan nomor HP berupa format teks dan pulihkan angka 0 di depan jika terpotong
      if (rowData[4] !== undefined && rowData[4] !== null) {
        let hpStr = String(rowData[4]).trim();
        if (/^8[0-9]{8,12}$/.test(hpStr)) {
          hpStr = "0" + hpStr;
        }
        rowData[4] = hpStr;
      }
    } else {
      // Data default jika baru pertama kali dibuat
      const isKepsek = (formattedCode === "01");
      const isMutiah = (formattedCode === "25" || name.indexOf("Muti'ah") !== -1);
      rowData = [
        formattedCode,
        name,
        isKepsek ? "Kepala Sekolah" : (isMutiah ? "Wali Kelas 2A Putra" : "Guru"),
        isMutiah ? "PERUM. ABC MACCOPA" : "-",
        "-",
        0,                                     // Gaji Pokok (akan diisi rumus masa kerja untuk guru)
        isKepsek ? 500000 : 0,                 // Tunjangan Jabatan
        isMutiah ? 100000 : 0,                 // Tunjangan Wali Kelas
        0,                                     // Tunjangan Konsumsi Transport (akan diisi rumus Hari x 7500)
        50000,                                 // Tunjangan Kesehatan
        0, 0, 0, 0,                            // Pinjaman Kas, Iuran SPP, Donasi, Lainnya
        "",                                    // TMT (Tgl Masuk)
        "Tidak",                               // Sertifikasi (Ya/Tidak)
        "Tidak",                               // Potongan Ujian 25% (Ya/Tidak)
        presensiHari !== null ? presensiHari : 0
      ];
    }
    gajiRows.push(rowData);
  });

  // Tulis ulang Master_Gaji secara rapi dan pastikan kolom A & E plain text
  sMasterGaji.clear();
  sMasterGaji.getRange(1, 1, 1, headerGaji.length).setValues([headerGaji]);
  const hRange = sMasterGaji.getRange(1, 1, 1, headerGaji.length);
  hRange.setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold");
  sMasterGaji.setFrozenRows(1);
  sMasterGaji.getRange("A:A").setNumberFormat("@");
  sMasterGaji.getRange("E:E").setNumberFormat("@"); // Kolom No HP selalu teks biasa agar angka 0 tidak hilang

  if (gajiRows.length > 0) {
    sMasterGaji.getRange(2, 5, gajiRows.length, 1).setNumberFormat("@");
    sMasterGaji.getRange(2, 1, gajiRows.length, headerGaji.length).setValues(gajiRows);
    
    // Format angka keuangan untuk kolom nominal (F s/d N)
    sMasterGaji.getRange(2, 6, gajiRows.length, 9).setNumberFormat("#,##0");
    // Format tanggal untuk kolom O (TMT)
    sMasterGaji.getRange(2, 15, gajiRows.length, 1).setNumberFormat("yyyy-mm-dd");
    // Format angka untuk kolom R (Kehadiran Transport Hari)
    sMasterGaji.getRange(2, 18, gajiRows.length, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");

    // Pasang rumus otomatis di Master_Gaji:
    // 1. Gaji Pokok (Kolom F / 6): otomatis masa kerja & 0 jika Sertifikasi = "Ya"
    // 2. Tunjangan Wali Kelas (Kolom H / 8): otomatis 0 jika Sertifikasi = "Ya"
    // 3. Tunjangan Konsumsi Transport (Kolom I / 9): otomatis Kehadiran (Kolom R) * 7500 & 0 jika Sertifikasi = "Ya"
    // 4. Tunjangan Kesehatan (Kolom J / 10): otomatis 0 jika Sertifikasi = "Ya", 50.000 jika "Tidak"
    for (let idx = 0; idx < gajiRows.length; idx++) {
      const r = idx + 2;
      const jabatan = String(gajiRows[idx][2] || "").trim().toLowerCase();
      const rawWali = Number(gajiRows[idx][7]) || (jabatan.indexOf("wali kelas") !== -1 ? 100000 : 0);

      // 1. Gaji Pokok (Kolom F)
      if (jabatan !== "kepala sekolah") {
        const formulaGajiPokok = '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IF(O' + r + '=""; 0; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=12; 500000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=9; 450000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=7; 400000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=5; 375000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=3; 300000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=2; 150000; 0))))))))';
        setSafeFormula(sMasterGaji.getRange(r, 6), formulaGajiPokok);
      }

      // 2. Tunjangan Wali Kelas (Kolom H): Otomatis membaca dari Tarif!H3 (atau 100.000) jika Jabatan mengandung "wali kelas" & 0 jika Sertifikasi = "Ya"
      const formulaWali = '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IF(ISNUMBER(SEARCH("wali kelas"; LOWER(C' + r + '))); IFERROR(Tarif!$H$3; 100000); 0))';
      setSafeFormula(sMasterGaji.getRange(r, 8), formulaWali);

      // 3. Tunjangan Konsumsi Transport (Kolom I): otomatis Kehadiran (Kolom R) * Tarif!H2 (atau 7500) & 0 jika Sertifikasi = "Ya"
      const formulaKonsumsi = '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IF(R' + r + '=""; 0; R' + r + ' * IFERROR(Tarif!$H$2; 7500)))';
      setSafeFormula(sMasterGaji.getRange(r, 9), formulaKonsumsi);

      // 4. Tunjangan Kesehatan (Kolom J): otomatis Tarif!H4 (atau 50.000) & 0 jika Sertifikasi = "Ya"
      const formulaKesehatan = '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IFERROR(Tarif!$H$4; 50000))';
      setSafeFormula(sMasterGaji.getRange(r, 10), formulaKesehatan);
    }

    // Pasang Data Validation Dropdown di Kolom P (Sertifikasi) dan Kolom Q (Potongan Ujian 25%)
    try {
      const ruleYaTidak = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Ya", "Tidak"], true)
        .setAllowInvalid(true)
        .build();
      sMasterGaji.getRange(2, 16, gajiRows.length, 2).setDataValidation(ruleYaTidak);
    } catch(e) {}
  }

  // 2. Siapkan Sheet Slip_Gaji
  let sSlip = ss.getSheetByName("Slip_Gaji");
  if (!sSlip) {
    sSlip = ss.insertSheet("Slip_Gaji");
  }

  // Pecahkan semua penggabungan sel lama agar tidak terjadi konflik merge
  try {
    const oldMerged = sSlip.getRange("A1:Z50").getMergedRanges();
    for (let i = 0; i < oldMerged.length; i++) {
      try { oldMerged[i].breakApart(); } catch(e) {}
    }
  } catch(e) {
    Logger.log("Breakapart log: " + e.toString());
  }

  // Bersihkan teks lama di area slip (C1:P35) agar tidak menyisakan teks bertumpuk dari merge sebelumnya
  sSlip.getRange("C1:P35").clearContent();

  // Atur Lebar Kolom (A - N) Sesuai Layout Baru
  sSlip.setColumnWidth(1, 135); // A (Banner)
  sSlip.setColumnWidth(2, 20);  // B (Pemisah / Kolom Kosong)
  sSlip.setColumnWidth(3, 35);  // C (NO / Label Utama)
  sSlip.setColumnWidth(4, 75);  // D
  sSlip.setColumnWidth(5, 75);  // E
  sSlip.setColumnWidth(6, 65);  // F
  sSlip.setColumnWidth(7, 65);  // G
  sSlip.setColumnWidth(8, 80);  // H (ALAMAT / HP Label)
  sSlip.setColumnWidth(9, 80);  // I
  sSlip.setColumnWidth(10, 70); // J
  sSlip.setColumnWidth(11, 35); // K (Rp / TGL.)
  sSlip.setColumnWidth(12, 45); // L (KODE Label / Nilai)
  sSlip.setColumnWidth(13, 45); // M (KODE Input / Nilai)
  sSlip.setColumnWidth(14, 45); // N (KODE Input / Nilai)

  // Set default font Times New Roman untuk seluruh slip
  sSlip.getRange("A1:N35").setFontFamily("Times New Roman");

  // Warna Hijau Header (Sesuai Screenshot Terbaru: Hijau Cerah #00e600)
  const GREEN_HEADER = "#00e600";

  // --- ROW 3: JUDUL SLIP GAJI & TANGGAL ---
  safeMerge(sSlip.getRange("C3:J3")).setValue("SLIP GAJI BULAN OKTOBER 2026")
    .setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sSlip.getRange("K3").setValue("TGL.").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false);
  safeMerge(sSlip.getRange("L3:N3")).setValue("01 Oktober 2026").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false);

  // --- ROW 5: NAMA, ALAMAT, KODE ---
  sSlip.getRange("C5").setValue("NAMA").setFontWeight("bold").setBorder(true, true, true, true, false, false);
  
  // Rumus Nama Guru (D5:G5)
  setSafeFormula(safeMerge(sSlip.getRange("D5:G5")), '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); ""; IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$E; 2; 0); IFERROR(VLOOKUP(IF($M$5<>"";$M$5;$N$5); Master_Gaji!$A:$E; 2; 0); "")))');
  sSlip.getRange("D5:G5").setFontWeight("bold").setHorizontalAlignment("left").setBorder(true, true, true, true, false, false);

  sSlip.getRange("H5").setValue("ALAMAT").setFontWeight("bold").setBorder(true, true, true, true, false, false);
  setSafeFormula(safeMerge(sSlip.getRange("I5:K5")), '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); ""; IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$E; 4; 0); IFERROR(VLOOKUP(IF($M$5<>"";$M$5;$N$5); Master_Gaji!$A:$E; 4; 0); "")))');
  sSlip.getRange("I5:K5").setHorizontalAlignment("left").setBorder(true, true, true, true, false, false);

  sSlip.getRange("L5").setValue("KODE").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, false, false);
  
  // Format sel input kode (M5:N5) warna hijau terang sesuai screenshot
  sSlip.getRange("M5:N5").setNumberFormat("@");
  safeMerge(sSlip.getRange("M5:N5")).setValue("2").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center")
    .setBackground(GREEN_HEADER)
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Buat Data Validation dropdown di M5 secara dinamis membaca dari Master_Gaji (mencakup s/d kode 60)
  try {
    const listKode = [];
    const lastRowGaji = sMasterGaji.getLastRow();
    if (lastRowGaji >= 2) {
      const kodes = sMasterGaji.getRange(2, 1, lastRowGaji - 1, 1).getValues();
      kodes.forEach(k => {
        const str = String(k[0]).trim();
        if (str) {
          const num = parseInt(str, 10);
          if (!isNaN(num)) {
            listKode.push(String(num));
            if (num < 10) listKode.push("0" + num);
          } else {
            listKode.push(str);
          }
        }
      });
    }
    for (let i = 1; i <= 60; i++) {
      listKode.push(String(i));
    }
    const uniqueKodes = [...new Set(listKode)].sort((a, b) => (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0));
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueKodes, true)
      .setAllowInvalid(true)
      .build();
    sSlip.getRange("M5").setDataValidation(rule);
  } catch(e) {}

  // --- ROW 6: JABATAN & HP ---
  sSlip.getRange("C6").setValue("JABATAN").setFontWeight("bold").setBorder(true, true, true, true, false, false);
  setSafeFormula(safeMerge(sSlip.getRange("D6:G6")), '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); ""; IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$E; 3; 0); IFERROR(VLOOKUP(IF($M$5<>"";$M$5;$N$5); Master_Gaji!$A:$E; 3; 0); "")))');
  sSlip.getRange("D6:G6").setHorizontalAlignment("left").setBorder(true, true, true, true, false, false);

  sSlip.getRange("H6").setValue("HP").setFontWeight("bold").setBorder(true, true, true, true, false, false);
  setSafeFormula(safeMerge(sSlip.getRange("I6:K6")), '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); ""; IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$E; 5; 0); IFERROR(VLOOKUP(IF($M$5<>"";$M$5;$N$5); Master_Gaji!$A:$E; 5; 0); "-")))');
  sSlip.getRange("I6:K6").setHorizontalAlignment("left").setNumberFormat("@").setBorder(true, true, true, true, false, false);

  // --- ROW 9 & 10: TABEL PENERIMAAN ---
  sSlip.getRange("C9").setValue("PENERIMAAN").setFontWeight("bold").setFontSize(11);
  sSlip.getRange("C10").setValue("NO.").setFontWeight("bold").setHorizontalAlignment("center").setBackground(GREEN_HEADER).setBorder(true, true, true, true, false, false);
  safeMerge(sSlip.getRange("D10:J10")).setValue("KETERANGAN").setFontWeight("bold").setHorizontalAlignment("center").setBackground(GREEN_HEADER).setBorder(true, true, true, true, false, false);
  safeMerge(sSlip.getRange("K10:N10")).setValue("JUMLAH").setFontWeight("bold").setHorizontalAlignment("center").setBackground(GREEN_HEADER).setBorder(true, true, true, true, false, false);

  const fKBM1 = '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$J:$J)>0; SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); IF(SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$J:$J)>0; SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); IF(SUMIF(KBM!$E:$E; D5; KBM!$J:$J)>0; SUMIF(KBM!$E:$E; D5; KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); IF(SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$J:$J)>0; SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); "-")))))';
  const fKBM2 = '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$K:$K)>0; SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); IF(SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$K:$K)>0; SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); IF(SUMIF(KBM!$E:$E; D5; KBM!$K:$K)>0; SUMIF(KBM!$E:$E; D5; KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); IF(SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$K:$K)>0; SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); "-")))))';
  const fUjian = '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(SUMIF(Ujian!$C:$C; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Ujian!$H:$H)>0; SUMIF(Ujian!$C:$C; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Ujian!$H:$H); IF(SUMIF(Ujian!$C:$C; VALUE(IF($M$5<>"";$M$5;$N$5)); Ujian!$H:$H)>0; SUMIF(Ujian!$C:$C; VALUE(IF($M$5<>"";$M$5;$N$5)); Ujian!$H:$H); IF(SUMIF(Ujian!$D:$D; D5; Ujian!$H:$H)>0; SUMIF(Ujian!$D:$D; D5; Ujian!$H:$H); IF(SUMIF(Ujian!$D:$D; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; Ujian!$H:$H)>0; SUMIF(Ujian!$D:$D; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; Ujian!$H:$H); "-")))))';

  const penerimaanItems = [
    [11, "1", "GAJI POKOK", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); 0; IF(UPPER(IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 16; 0); ""))="YA"; 0; IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 6; 0); IFERROR(VLOOKUP(IF($M$5<>"";$M$5;$N$5); Master_Gaji!$A:$R; 6; 0); 0))))'],
    [12, "2", "INSENTIF MENGAJAR KE-1 (Rp.12.000/JAM)", fKBM1],
    [13, "3", "INSENTIF MENGAJAR KE-2 (Rp.10.000/JAM)", fKBM2],
    [14, "4", "TUNJANGAN JABATAN", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 7; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 7; 0)); "-"))'],
    [15, "5", "TUNJANGAN WALI KELAS", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(UPPER(IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 16; 0); ""))="YA"; "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 8; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 8; 0)); "-")))'],
    [16, "6", "TUNJANGAN KONSUMSI + TRANSPORT", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(UPPER(IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 16; 0); ""))="YA"; "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 9; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 9; 0)); "-")))'],
    [17, "7", "TUNJANGAN KESEHATAN", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(UPPER(IFERROR(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 16; 0); ""))="YA"; "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 10; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 10; 0)); "-")))'],
    [18, "8", "INSENTIF PELAKSANAAN UJIAN", fUjian]
  ];

  penerimaanItems.forEach(item => {
    const row = item[0];
    sSlip.getRange(row, 3).setValue(item[1]).setHorizontalAlignment("center").setBorder(true, true, true, true, false, false);
    safeMerge(sSlip.getRange(row, 4, 1, 7)).setValue(item[2]).setHorizontalAlignment("left").setBorder(true, true, true, true, false, false);
    sSlip.getRange(row, 11).setValue("Rp").setHorizontalAlignment("left").setBorder(true, false, true, true, false, false);
    setSafeFormula(safeMerge(sSlip.getRange(row, 12, 1, 3)), item[3]);
    sSlip.getRange(row, 12, 1, 3).setHorizontalAlignment("right").setNumberFormat("#,##0").setBorder(true, true, true, false, false, false);
  });

  // Set label dinamis untuk KBM 1 & KBM 2 agar teks 'Rp.12.000' ikut berubah jika sel di sheet Tarif diubah
  setSafeFormula(sSlip.getRange("D12"), '="INSENTIF MENGAJAR KE-1 (Rp." & TEXT(IFERROR(Tarif!$E$2; 12000); "#,##0") & "/JAM)"');
  setSafeFormula(sSlip.getRange("D13"), '="INSENTIF MENGAJAR KE-2 (Rp." & TEXT(IFERROR(Tarif!$E$3; 10000); "#,##0") & "/JAM)"');

  // --- ROW 20: TOTAL PENGHASILAN ---
  safeMerge(sSlip.getRange("I20:J20")).setValue("TOTAL PENGHASILAN").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, false, false);
  sSlip.getRange("K20").setValue("Rp").setFontWeight("bold").setHorizontalAlignment("left").setBorder(true, false, true, true, false, false);
  setSafeFormula(safeMerge(sSlip.getRange("L20:N20")), '=IF(SUMIF(L11:L18; ">0")=0; 0; SUMIF(L11:L18; ">0"))');
  sSlip.getRange("L20:N20").setFontWeight("bold").setHorizontalAlignment("right").setNumberFormat("#,##0").setBorder(true, true, true, false, false, false);

  // --- ROW 21 & 22: TABEL PENGURANGAN ---
  sSlip.getRange("C21").setValue("PENGURANGAN").setFontWeight("bold").setFontSize(11);
  sSlip.getRange("C22").setValue("NO.").setFontWeight("bold").setHorizontalAlignment("center").setBackground(GREEN_HEADER).setBorder(true, true, true, true, false, false);
  safeMerge(sSlip.getRange("D22:J22")).setValue("KETERANGAN").setFontWeight("bold").setHorizontalAlignment("center").setBackground(GREEN_HEADER).setBorder(true, true, true, true, false, false);
  safeMerge(sSlip.getRange("K22:N22")).setValue("JUMLAH").setFontWeight("bold").setHorizontalAlignment("center").setBackground(GREEN_HEADER).setBorder(true, true, true, true, false, false);

  const penguranganItems = [
    [23, "1", "PINJAMAN KAS YAYASAN", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 11; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 11; 0)); "-"))'],
    [24, "2", "IURAN SPP ANAK", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 12; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 12; 0)); "-"))'],
    [25, "3", "DONASI", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 13; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 13; 0)); "-"))'],
    [26, "4", "LAINNYA", '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IFERROR(IF(VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 14; 0)=0; "-"; VLOOKUP(RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Master_Gaji!$A:$R; 14; 0)); "-"))']
  ];

  penguranganItems.forEach(item => {
    const row = item[0];
    sSlip.getRange(row, 3).setValue(item[1]).setHorizontalAlignment("center").setBorder(true, true, true, true, false, false);
    safeMerge(sSlip.getRange(row, 4, 1, 7)).setValue(item[2]).setHorizontalAlignment("left").setBorder(true, true, true, true, false, false);
    sSlip.getRange(row, 11).setValue("Rp").setHorizontalAlignment("left").setBorder(true, false, true, true, false, false);
    setSafeFormula(safeMerge(sSlip.getRange(row, 12, 1, 3)), item[3]);
    sSlip.getRange(row, 12, 1, 3).setHorizontalAlignment("right").setNumberFormat("#,##0").setBorder(true, true, true, false, false, false);
  });

  // --- ROW 28: BENDAHARA & TOTAL PENGURANGAN ---
  safeMerge(sSlip.getRange("C28:G28")).setValue("BENDAHARA YAYASAN,").setFontSize(10).setHorizontalAlignment("left");
  safeMerge(sSlip.getRange("I28:J28")).setValue("TOTAL PENGURANGAN").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, false, false);
  sSlip.getRange("K28").setValue("Rp").setFontWeight("bold").setHorizontalAlignment("left").setBorder(true, false, true, true, false, false);
  setSafeFormula(safeMerge(sSlip.getRange("L28:N28")), '=IF(SUMIF(L23:L26; ">0")=0; "-"; SUMIF(L23:L26; ">0"))');
  sSlip.getRange("L28:N28").setFontWeight("bold").setHorizontalAlignment("right").setNumberFormat("#,##0").setBorder(true, true, true, false, false, false);

  // --- ROW 30: NAMA BENDAHARA & TOTAL GAJI DITERIMA ---
  safeMerge(sSlip.getRange("C30:G30")).setValue("NAILAH FARAFISHAH, S.S.").setFontWeight("bold").setFontSize(10).setFontLine("underline").setHorizontalAlignment("left");
  safeMerge(sSlip.getRange("I30:J30")).setValue("TOTAL GAJI DITERIMA").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sSlip.getRange("K30").setValue("Rp").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left")
    .setBorder(true, false, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  setSafeFormula(safeMerge(sSlip.getRange("L30:N30")), '=L20 - IF(ISNUMBER(L28); L28; 0)');
  sSlip.getRange("L30:N30").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("right").setNumberFormat("#,##0")
    .setBorder(true, true, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // --- BINGKAI LUAR KOTAK CETAK SLIP GAJI (C8:N31) ---
  sSlip.getRange("C8:N31").setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // --- BANNER DI KOLOM A (A3:A31) ---
  sSlip.getRange("A3:A31").setBorder(true, true, true, true, false, false, "#16a34a", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  Logger.log("✅ Slip Gaji & Master Gaji Berhasil Disesuaikan!");
}

/**
 * Helper untuk merge sel secara aman tanpa terkena error rentang konflik
 */
function safeMerge(range) {
  try {
    const merged = range.getMergedRanges();
    for (let i = 0; i < merged.length; i++) {
      try { merged[i].breakApart(); } catch(e) {}
    }
  } catch(e) {}
  try {
    return range.merge();
  } catch(e2) {
    return range;
  }
}

/**
 * Helper untuk menyetel formula secara tepat sesuai regional Google Sheets
 */
function setSafeFormula(range, formula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const loc = (ss.getSpreadsheetLocale() || "").toLowerCase();
  
  // Format US / desimal titik menggunakan koma (,)
  const isCommaLocale = /^(en_us|en_au|en_ca|en_gb|en_in|en_ph|en_sg|ja_jp|zh_cn|ko_kr)/i.test(loc);
  
  // Standar Indonesia (id_ID) menggunakan titik koma (;)
  const targetFormula = isCommaLocale ? formula.replace(/;/g, ",") : formula;
  
  try {
    range.setFormula(targetFormula);
  } catch(e) {
    // Fallback: coba format sebaliknya jika ada pengaturan khusus
    try {
      range.setFormula(isCommaLocale ? formula : formula.replace(/;/g, ","));
    } catch(e2) {
      Logger.log("Gagal set formula: " + targetFormula + " | " + e2.toString());
    }
  }
}

/**
 * Mengurutkan ulang (Re-number) seluruh kode guru sesuai aturan blok:
 * - GURU PUTRA : 01 s/d 20 (Kepsek Radiant Fadli 01, Bahtiar 02, dst.)
 * - GURU PUTRI : 21 s/d 50 (Dra. Mardia Habe 21, Hj. Nur Apni 22, dst.)
 * Otomatis menyinkronkan kodenya di sheet Master, Master_Gaji, KBM, dan Ujian
 * sehingga riwayat presensi jam mengajar lama tetap aman dan tidak tertukar.
 */
function urutkanUlangKodeGuru() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sMaster = getSheet(SHEET_MASTER);
  const lastRowMaster = sMaster.getLastRow();
  if (lastRowMaster < 2) return;

  const masterData = sMaster.getRange(2, 1, lastRowMaster - 1, 3).getValues();

  // Daftar nama resmi Guru Putra (Ikhwan)
  const LIST_PUTRA = [
    "radiant fadli",
    "bahtiar",
    "agung setiabudi",
    "fandhy ahmad klana",
    "muhammad ihsan",
    "turmisi",
    "wahyu saputra",
    "syamsir",
    "ahmad prisatrialdi",
    "syahrul ramadhan",
    "muh. sofyan asnur",
    "muhammad fikri haekal",
    "munadil amri",
    "fuad"
  ];

  // Daftar urutan resmi Guru Putri (Akhwat)
  const LIST_PUTRI = [
    "mardia habe",
    "nur apni",
    "amalia ahsan",
    "dewi syahara",
    "ina mulyati",
    "humairah",
    "fitriani",
    "satriani",
    "nur wahidah",
    "muti'ah",
    "liza fany",
    "darmawati",
    "ilma",
    "nurdiyawanti",
    "rahmawati basri",
    "sri wahyuni",
    "ihwana",
    "surianti",
    "hajriani",
    "nur rizky",
    "hebrianti",
    "dhea rahmiyanti"
  ];

  function cekApakahPutra(name) {
    const n = name.toLowerCase().trim();
    for (let p of LIST_PUTRA) {
      if (n.indexOf(p) !== -1) return true;
    }
    return false;
  }

  const putraList = [];
  const putriList = [];

  for (let i = 0; i < masterData.length; i++) {
    const rawOldCode = String(masterData[i][0]).trim();
    const name = String(masterData[i][1]).trim();
    const pin = masterData[i][2] || "123456";

    // Abaikan baris kosong atau guru yang sudah ditandai keluar
    if (!name || name.toLowerCase().indexOf("[keluar]") !== -1) continue;

    if (cekApakahPutra(name)) {
      putraList.push({ oldCode: rawOldCode, name: name, pin: pin });
    } else {
      putriList.push({ oldCode: rawOldCode, name: name, pin: pin });
    }
  }

  // Urutkan Guru Putra sesuai urutan hierarki resmi
  putraList.sort((a, b) => {
    const na = a.name.toLowerCase();
    const nb = b.name.toLowerCase();
    let idxA = LIST_PUTRA.findIndex(p => na.indexOf(p) !== -1);
    let idxB = LIST_PUTRA.findIndex(p => nb.indexOf(p) !== -1);
    if (idxA === -1) idxA = 999;
    if (idxB === -1) idxB = 999;
    return idxA - idxB;
  });

  // Urutkan Guru Putri sesuai urutan resmi
  putriList.sort((a, b) => {
    const na = a.name.toLowerCase();
    const nb = b.name.toLowerCase();
    let idxA = LIST_PUTRI.findIndex(p => na.indexOf(p) !== -1);
    let idxB = LIST_PUTRI.findIndex(p => nb.indexOf(p) !== -1);
    if (idxA === -1) idxA = 999;
    if (idxB === -1) idxB = 999;
    return idxA - idxB;
  });

  const codeMapping = {};
  const newMasterRows = [];

  // 1. Berikan nomor urut Putra: 01 s/d 20
  let numPutra = 1;
  putraList.forEach(p => {
    const newCode = numPutra < 10 ? "0" + numPutra : String(numPutra);
    codeMapping[p.name.toLowerCase()] = newCode;
    if (p.oldCode) {
      const formattedOld = p.oldCode.length === 1 ? "0" + p.oldCode : p.oldCode;
      codeMapping["old_" + formattedOld] = newCode;
    }
    newMasterRows.push([newCode, p.name, p.pin]);
    numPutra++;
  });

  // 2. Berikan nomor urut Putri: 21 s/d 50
  let numPutri = 21;
  putriList.forEach(p => {
    const newCode = String(numPutri);
    codeMapping[p.name.toLowerCase()] = newCode;
    if (p.oldCode) {
      const formattedOld = p.oldCode.length === 1 ? "0" + p.oldCode : p.oldCode;
      codeMapping["old_" + formattedOld] = newCode;
    }
    newMasterRows.push([newCode, p.name, p.pin]);
    numPutri++;
  });

  // Tulis ulang Sheet Master kolom A:C
  sMaster.getRange(2, 1, lastRowMaster - 1, 3).clearContent();
  sMaster.getRange("A:A").setNumberFormat("@");
  if (newMasterRows.length > 0) {
    sMaster.getRange(2, 1, newMasterRows.length, 3).setValues(newMasterRows);
  }

  // Sinkronkan ke sheet KBM dan Ujian via fungsi sinkronisasi cerdas
  sinkronkanSheetKBM(ss, newMasterRows, codeMapping);
  sinkronkanSheetUjian(ss, newMasterRows, codeMapping);

  // Perbarui Master_Gaji dan suntikkan formula aman ke Slip_Gaji
  sinkronkanMasterGaji(codeMapping);

  SpreadsheetApp.getActiveSpreadsheet().toast("Susunan berhasil dirapikan: Putra (01-20), Putri (21-50)!", "Sukses 🟢", 5);
}

/**
 * Menyinkronkan data guru ke Master_Gaji TANPA menyentuh/mengubah sheet Slip_Gaji
 * sehingga format dan kerapian lembar cetak slip gaji yang diatur pengguna tidak pernah rusak/tertimpa.
 */
function sinkronkanMasterGaji(codeMapping) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Pastikan sheet Tarif tersedia
  setupSheetTarif(ss);

  let sMasterGaji = ss.getSheetByName("Master_Gaji");
  if (!sMasterGaji) return;

  const headerGaji = [
    "Kode", "Nama Guru", "Jabatan", "Alamat", "No HP",
    "Gaji Pokok", "Tunjangan Jabatan", "Tunjangan Wali Kelas",
    "Tunjangan Konsumsi Transport", "Tunjangan Kesehatan",
    "Pinjaman Kas", "Iuran SPP", "Donasi", "Lainnya",
    "TMT (Tgl Masuk)", "Sertifikasi", "Potongan Ujian 25%", "Kehadiran Transport (Hari)"
  ];

  const sMaster = getSheet(SHEET_MASTER);
  const lastMasterRow = sMaster.getLastRow();
  let teacherList = [];
  if (lastMasterRow >= 2) {
    teacherList = sMaster.getRange(2, 1, lastMasterRow - 1, 2).getValues();
  }

  // Simpan data lama di Master_Gaji agar data yang sudah diisi admin tidak hilang
  const existingGaji = {};
  const existingGajiByName = {};
  const existingGajiByNameClean = {};
  const curLastRow = sMasterGaji.getLastRow();
  const curLastCol = sMasterGaji.getLastColumn();
  if (curLastRow >= 2 && curLastCol > 0) {
    const oldVals = sMasterGaji.getRange(2, 1, curLastRow - 1, curLastCol).getValues();
    oldVals.forEach(r => {
      const k = String(r[0]).trim();
      const n = String(r[1]).trim().toLowerCase();
      const nClean = n.replace(/[^a-z0-9]/g, "");
      if (k) existingGaji[k] = r;
      if (n) existingGajiByName[n] = r;
      if (nClean) existingGajiByNameClean[nClean] = r;
    });
  }

  const sPresensi = ss.getSheetByName("Presensi");
  const attendanceMap = sPresensi ? parseSheetPresensi(sPresensi) : {};

  const gajiRows = [];
  teacherList.forEach(t => {
    const rawCode = String(t[0]).trim();
    if (!rawCode && !t[1]) return;
    const formattedCode = rawCode.length === 1 ? "0" + rawCode : rawCode;
    const name = String(t[1]).trim();
    const nameLower = name.toLowerCase();
    const cleanName = nameLower.replace(/[^a-z0-9]/g, "");
    const presensiHari = cariKehadiranGuru(name, attendanceMap);

    let rowData = [];
    // Prioritaskan pencocokan berdasarkan NAMA GURU agar jabatan, alamat, TMT, dll tidak tertukar saat nomor kode diurutkan
    const exist = existingGajiByName[nameLower] || existingGajiByNameClean[cleanName] || existingGaji[formattedCode] || existingGaji[rawCode];
    if (exist) {
      rowData = exist.slice(0, headerGaji.length);
      while (rowData.length < headerGaji.length) {
        if (rowData.length === 14) rowData.push("");
        else if (rowData.length === 15) rowData.push("Tidak");
        else if (rowData.length === 16) rowData.push("Tidak");
        else if (rowData.length === 17) {
          // Kolom R: Kehadiran Transport (Hari)
          let defaultHari = 0;
          if (presensiHari !== null) {
            defaultHari = presensiHari;
          } else if (exist[8] !== undefined && Number(exist[8]) > 0) {
            defaultHari = Number(exist[8]) <= 31 ? Number(exist[8]) : Math.round(Number(exist[8]) / 7500);
          }
          rowData.push(defaultHari);
        }
        else rowData.push(0);
      }
      rowData[0] = formattedCode;
      rowData[1] = name;
      // Jika sheet Presensi ada dan memuat kehadiran guru ini, gunakan angka terbaru dari Presensi
      if (presensiHari !== null) {
        rowData[17] = presensiHari;
      }
      // Pastikan nomor HP berupa format teks dan pulihkan angka 0 di depan jika terpotong
      if (rowData[4] !== undefined && rowData[4] !== null) {
        let hpStr = String(rowData[4]).trim();
        if (/^8[0-9]{8,12}$/.test(hpStr)) {
          hpStr = "0" + hpStr;
        }
        rowData[4] = hpStr;
      }
    } else {
      const isKepsek = (formattedCode === "01");
      const isMutiah = (formattedCode === "25" || name.indexOf("Muti'ah") !== -1);
      rowData = [
        formattedCode, name,
        isKepsek ? "Kepala Sekolah" : (isMutiah ? "Wali Kelas 2A Putra" : "Guru"),
        isMutiah ? "PERUM. ABC MACCOPA" : "-", "-",
        0, isKepsek ? 500000 : 0, isMutiah ? 100000 : 0, 0, 50000,
        0, 0, 0, 0, "", "Tidak", "Tidak",
        presensiHari !== null ? presensiHari : 0
      ];
    }
    gajiRows.push(rowData);
  });

  // Tulis ulang hanya sheet Master_Gaji
  sMasterGaji.clear();
  sMasterGaji.getRange(1, 1, 1, headerGaji.length).setValues([headerGaji]);
  sMasterGaji.getRange(1, 1, 1, headerGaji.length).setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold");
  sMasterGaji.setFrozenRows(1);
  sMasterGaji.getRange("A:A").setNumberFormat("@");
  sMasterGaji.getRange("E:E").setNumberFormat("@"); // Kolom No HP selalu teks biasa agar angka 0 tidak hilang

  if (gajiRows.length > 0) {
    sMasterGaji.getRange(2, 5, gajiRows.length, 1).setNumberFormat("@");
    sMasterGaji.getRange(2, 1, gajiRows.length, headerGaji.length).setValues(gajiRows);
    sMasterGaji.getRange(2, 6, gajiRows.length, 9).setNumberFormat("#,##0");
    sMasterGaji.getRange(2, 15, gajiRows.length, 1).setNumberFormat("yyyy-mm-dd");
    // Format angka untuk kolom R (Kehadiran Transport Hari)
    sMasterGaji.getRange(2, 18, gajiRows.length, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");

    for (let idx = 0; idx < gajiRows.length; idx++) {
      const r = idx + 2;
      const jabatan = String(gajiRows[idx][2] || "").trim().toLowerCase();
      const rawWali = Number(gajiRows[idx][7]) || (jabatan.indexOf("wali kelas") !== -1 ? 100000 : 0);

      if (jabatan !== "kepala sekolah") {
        const formulaGajiPokok = '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IF(O' + r + '=""; 0; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=12; 500000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=9; 450000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=7; 400000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=5; 375000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=3; 300000; IF(DATEDIF(O' + r + '; TODAY(); "Y")>=2; 150000; 0))))))))';
        setSafeFormula(sMasterGaji.getRange(r, 6), formulaGajiPokok);
      }
      // 2. Tunjangan Wali Kelas (Kolom H): Otomatis membaca dari Tarif!H3 (atau 100.000) jika Jabatan mengandung "wali kelas" & 0 jika Sertifikasi = "Ya"
      const formulaWali = '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IF(ISNUMBER(SEARCH("wali kelas"; LOWER(C' + r + '))); IFERROR(Tarif!$H$3; 100000); 0))';
      setSafeFormula(sMasterGaji.getRange(r, 8), formulaWali);

      // 3. Tunjangan Konsumsi Transport (Kolom I): otomatis Kehadiran (Kolom R) * Tarif!H2 (atau 7500) & 0 jika Sertifikasi = "Ya"
      const formulaKonsumsi = '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IF(R' + r + '=""; 0; R' + r + ' * IFERROR(Tarif!$H$2; 7500)))';
      setSafeFormula(sMasterGaji.getRange(r, 9), formulaKonsumsi);

      // 4. Tunjangan Kesehatan (Kolom J): otomatis Tarif!H4 (atau 50.000) & 0 jika Sertifikasi = "Ya"
      setSafeFormula(sMasterGaji.getRange(r, 10), '=IF(UPPER(TRIM(P' + r + '))="YA"; 0; IFERROR(Tarif!$H$4; 50000))');
    }

    try {
      const ruleYaTidak = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Ya", "Tidak"], true)
        .setAllowInvalid(true)
        .build();
      sMasterGaji.getRange(2, 16, gajiRows.length, 2).setDataValidation(ruleYaTidak);
    } catch(e) {}
  }

  // Sinkronkan sheet Ujian dan KBM secara otomatis
  sinkronkanSheetUjian(ss, teacherList, codeMapping);
  sinkronkanSheetKBM(ss, teacherList, codeMapping);

  // Perbarui formula penting di Slip_Gaji secara aman TANPA merusak layout / format cetak yang sudah dirapikan pengguna
  const sSlip = ss.getSheetByName("Slip_Gaji");
  if (sSlip) {
    const fKBM1 = '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$J:$J)>0; SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); IF(SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$J:$J)>0; SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); IF(SUMIF(KBM!$E:$E; D5; KBM!$J:$J)>0; SUMIF(KBM!$E:$E; D5; KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); IF(SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$J:$J)>0; SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$J:$J)*IFERROR(Tarif!$E$2; 12000); "-")))))';
    const fKBM2 = '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$K:$K)>0; SUMIF(KBM!$D:$D; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); IF(SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$K:$K)>0; SUMIF(KBM!$D:$D; VALUE(IF($M$5<>"";$M$5;$N$5)); KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); IF(SUMIF(KBM!$E:$E; D5; KBM!$K:$K)>0; SUMIF(KBM!$E:$E; D5; KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); IF(SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$K:$K)>0; SUMIF(KBM!$E:$E; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; KBM!$K:$K)*IFERROR(Tarif!$E$3; 10000); "-")))))';
    const fUjian = '=IF(OR(IF($M$5<>"";$M$5;$N$5)=""; IF($M$5<>"";$M$5;$N$5)=0); "-"; IF(SUMIF(Ujian!$C:$C; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Ujian!$H:$H)>0; SUMIF(Ujian!$C:$C; RIGHT("0"&IF($M$5<>"";$M$5;$N$5); 2); Ujian!$H:$H); IF(SUMIF(Ujian!$C:$C; VALUE(IF($M$5<>"";$M$5;$N$5)); Ujian!$H:$H)>0; SUMIF(Ujian!$C:$C; VALUE(IF($M$5<>"";$M$5;$N$5)); Ujian!$H:$H); IF(SUMIF(Ujian!$D:$D; D5; Ujian!$H:$H)>0; SUMIF(Ujian!$D:$D; D5; Ujian!$H:$H); IF(SUMIF(Ujian!$D:$D; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; Ujian!$H:$H)>0; SUMIF(Ujian!$D:$D; "*"&TRIM(LEFT(SUBSTITUTE(SUBSTITUTE(D5; ","; " "); "."; " "); 8))&"*"; Ujian!$H:$H); "-")))))';

    setSafeFormula(sSlip.getRange("L12"), fKBM1);
    setSafeFormula(sSlip.getRange("L13"), fKBM2);
    setSafeFormula(sSlip.getRange("D12"), '="INSENTIF MENGAJAR KE-1 (Rp." & TEXT(IFERROR(Tarif!$E$2; 12000); "#,##0") & "/JAM)"');
    setSafeFormula(sSlip.getRange("D13"), '="INSENTIF MENGAJAR KE-2 (Rp." & TEXT(IFERROR(Tarif!$E$3; 10000); "#,##0") & "/JAM)"');
    setSafeFormula(sSlip.getRange("L18"), fUjian);
    setSafeFormula(sSlip.getRange("L20"), '=IF(SUMIF(L11:L18; ">0")=0; 0; SUMIF(L11:L18; ">0"))');
    setSafeFormula(sSlip.getRange("L28"), '=IF(SUMIF(L23:L26; ">0")=0; "-"; SUMIF(L23:L26; ">0"))');
    setSafeFormula(sSlip.getRange("L30"), '=L20 - IF(ISNUMBER(L28); L28; 0)');
  }
}

/**
 * Menyelaraskan sheet Ujian:
 * - Memastikan seluruh kode guru di kolom C menggunakan format teks 2 digit ('01', '02', '21')
 * - Mencocokkan nama guru dengan Master menggunakan normalisasi cerdas
 * - Membersihkan kolom Nominal (H) menjadi angka murni agar dapat dijumlahkan oleh SUMIF
 * - Jika nominal masih kosong/0, otomatis menghitung nominal berdasarkan kategori & volume
 */
function sinkronkanSheetUjian(ss, teacherList, codeMapping) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sUjian = ss.getSheetByName(SHEET_UJIAN);
  if (!sUjian || sUjian.getLastRow() < 2) return;

  if (!teacherList || teacherList.length === 0) {
    const sMaster = ss.getSheetByName(SHEET_MASTER);
    if (sMaster && sMaster.getLastRow() >= 2) {
      teacherList = sMaster.getRange(2, 1, sMaster.getLastRow() - 1, 2).getValues();
    } else {
      teacherList = [];
    }
  }

  const teacherLookup = [];
  teacherList.forEach(t => {
    const rawC = String(t[0]).trim();
    const formattedC = rawC ? (rawC.length === 1 ? "0" + rawC : rawC) : "";
    const name = String(t[1]).trim();
    if (name) {
      teacherLookup.push({
        code: formattedC,
        name: name,
        norm: normalizeTeacherName(name)
      });
    }
  });

  const lastRow = sUjian.getLastRow();
  const numRows = lastRow - 1;
  const numCols = Math.max(sUjian.getLastColumn(), 8);
  const ujianData = sUjian.getRange(2, 1, numRows, numCols).getValues();

  const newCodes = [];
  const newNames = [];
  const newNominals = [];

  for (let i = 0; i < ujianData.length; i++) {
    const row = ujianData[i];
    const rawCode = String(row[2] || "").trim();
    const formattedCode = rawCode ? (rawCode.length === 1 ? "0" + rawCode : rawCode) : "";
    const rawName = String(row[3] || "").trim();
    const normName = normalizeTeacherName(rawName);

    let matchedCode = "";
    let matchedName = rawName;

    if (normName) {
      const match = teacherLookup.find(t => t.norm === normName || t.norm.indexOf(normName) !== -1 || normName.indexOf(t.norm) !== -1);
      if (match) {
        matchedCode = match.code;
        matchedName = match.name;
      }
    }

    if (!matchedCode && formattedCode) {
      if (codeMapping && codeMapping["old_" + formattedCode]) {
        matchedCode = codeMapping["old_" + formattedCode];
      } else {
        const matchC = teacherLookup.find(t => t.code === formattedCode);
        if (matchC) {
          matchedCode = matchC.code;
          if (!matchedName) matchedName = matchC.name;
        } else {
          matchedCode = formattedCode;
        }
      }
    }

    newCodes.push([matchedCode || formattedCode]);
    newNames.push([matchedName]);

    let rawNom = row[7];
    let nomNum = 0;
    if (typeof rawNom === "string") {
      const digits = rawNom.replace(/[^0-9]/g, "");
      nomNum = digits ? Number(digits) : 0;
    } else {
      nomNum = Number(rawNom) || 0;
    }

    if (row[4]) {
      const autoNom = hitungNominalUjian(row[4], row[6] || 1, matchedCode, matchedName);
      if (autoNom > 0) {
        nomNum = autoNom;
      }
    }

    newNominals.push([nomNum]);
  }

  sUjian.getRange("C:C").setNumberFormat("@");
  sUjian.getRange(2, 3, numRows, 1).setValues(newCodes);
  sUjian.getRange(2, 4, numRows, 1).setValues(newNames);
  sUjian.getRange(2, 8, numRows, 1).setValues(newNominals);
  sUjian.getRange(2, 8, numRows, 1).setNumberFormat("#,##0");
}

/**
 * Menyelaraskan sheet KBM:
 * - Memastikan seluruh kode guru di kolom D menggunakan format teks 2 digit ('01', '02', '21')
 * - Mencocokkan nama guru dengan Master
 * - Memastikan Jam Sendiri (J) dan Jam Team (K) adalah angka murni
 */
function sinkronkanSheetKBM(ss, teacherList, codeMapping) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sKBM = ss.getSheetByName(SHEET_KBM);
  if (!sKBM || sKBM.getLastRow() < 2) return;

  if (!teacherList || teacherList.length === 0) {
    const sMaster = ss.getSheetByName(SHEET_MASTER);
    if (sMaster && sMaster.getLastRow() >= 2) {
      teacherList = sMaster.getRange(2, 1, sMaster.getLastRow() - 1, 2).getValues();
    } else {
      teacherList = [];
    }
  }

  const teacherLookup = [];
  teacherList.forEach(t => {
    const rawC = String(t[0]).trim();
    const formattedC = rawC ? (rawC.length === 1 ? "0" + rawC : rawC) : "";
    const name = String(t[1]).trim();
    if (name) {
      teacherLookup.push({
        code: formattedC,
        name: name,
        norm: normalizeTeacherName(name)
      });
    }
  });

  const lastRow = sKBM.getLastRow();
  const numRows = lastRow - 1;
  const numCols = Math.max(sKBM.getLastColumn(), 11);
  const kbmData = sKBM.getRange(2, 1, numRows, numCols).getValues();

  const newCodes = [];
  const newNames = [];
  const newHours = [];

  for (let i = 0; i < kbmData.length; i++) {
    const row = kbmData[i];
    const rawCode = String(row[3] || "").trim();
    const formattedCode = rawCode ? (rawCode.length === 1 ? "0" + rawCode : rawCode) : "";
    const rawName = String(row[4] || "").trim();
    const normName = normalizeTeacherName(rawName);

    let matchedCode = "";
    let matchedName = rawName;

    if (normName) {
      const match = teacherLookup.find(t => t.norm === normName || t.norm.indexOf(normName) !== -1 || normName.indexOf(t.norm) !== -1);
      if (match) {
        matchedCode = match.code;
        matchedName = match.name;
      }
    }

    if (!matchedCode && formattedCode) {
      if (codeMapping && codeMapping["old_" + formattedCode]) {
        matchedCode = codeMapping["old_" + formattedCode];
      } else {
        const matchC = teacherLookup.find(t => t.code === formattedCode);
        if (matchC) {
          matchedCode = matchC.code;
          if (!matchedName) matchedName = matchC.name;
        } else {
          matchedCode = formattedCode;
        }
      }
    }

    newCodes.push([matchedCode || formattedCode]);
    newNames.push([matchedName]);

    const jamSendiri = Number(row[9]) || 0;
    const jamTeam = Number(row[10]) || 0;
    newHours.push([jamSendiri, jamTeam]);
  }

  sKBM.getRange("D:D").setNumberFormat("@");
  sKBM.getRange(2, 4, numRows, 1).setValues(newCodes);
  sKBM.getRange(2, 5, numRows, 1).setValues(newNames);
  sKBM.getRange(2, 10, numRows, 2).setValues(newHours);
  sKBM.getRange(2, 10, numRows, 2).setNumberFormat("#,##0");
}

/**
 * =========================================================================
 * INTEGRASI PRESENSI KEHADIRAN (OTOMATIS DARI SHEET 'Presensi')
 * =========================================================================
 * Menarik total kehadiran dari Sheet Presensi ke Kolom R di Master_Gaji.
 * Otomatis menghitung Tunjangan Konsumsi Transport di Kolom I (Hari x 7.500).
 */
function tarikKehadiranPresensi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sPresensi = ss.getSheetByName("Presensi");
  const sMasterGaji = ss.getSheetByName("Master_Gaji");

  if (!sPresensi) {
    SpreadsheetApp.getUi().alert(
      "Sheet 'Presensi' Belum Ditemukan ⚠️\n\n" +
      "Silakan buat Sheet baru bernama 'Presensi' di spreadsheet ini, " +
      "lalu tempelkan (paste) seluruh data presensi pengajar Anda di sheet tersebut."
    );
    return;
  }

  if (!sMasterGaji || sMasterGaji.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Sheet 'Master_Gaji' belum siap. Jalankan 'Urutkan Ulang Nomor Kode Guru' terlebih dahulu.");
    return;
  }

  // Ambil pemetaan kehadiran dari Sheet Presensi
  const attendanceMap = parseSheetPresensi(sPresensi);
  const foundNames = Object.keys(attendanceMap);

  if (foundNames.length === 0) {
    SpreadsheetApp.getUi().alert(
      "Data Presensi Tidak Dikenali ⚠️\n\n" +
      "Tidak ditemukan data kehadiran di sheet 'Presensi'. " +
      "Pastikan tabel presensi memiliki kolom 'TOTAL KEHADIRAN' dengan angka jumlah kehadiran di bawahnya."
    );
    return;
  }

  // Update Kolom R di Master_Gaji
  const lastRowGaji = sMasterGaji.getLastRow();
  const gajiData = sMasterGaji.getRange(2, 1, lastRowGaji - 1, 18).getValues();
  let updatedCount = 0;

  for (let i = 0; i < gajiData.length; i++) {
    const teacherName = String(gajiData[i][1] || "").trim();
    if (!teacherName) continue;

    const matchedHari = cariKehadiranGuru(teacherName, attendanceMap);
    sMasterGaji.getRange(i + 2, 18).setValue(matchedHari !== null ? matchedHari : 0);
    if (matchedHari !== null) {
      updatedCount++;
    }
  }

  // Perbarui rumus dan sinkronkan Ujian & KBM otomatis
  sinkronkanMasterGaji();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Berhasil memperbarui kehadiran " + updatedCount + " guru dari Sheet Presensi! Tunjangan transport otomatis terhitung di Kolom I.",
    "Sukses 🟢",
    6
  );
}

/**
 * Membaca sheet Presensi dan mengekstrak jumlah total kehadiran setiap pengajar.
 * Fokus pada teks tabel: nama guru di atas tabel dan angka terbesar di kolom TOTAL KEHADIRAN.
 */
function parseSheetPresensi(sPresensi) {
  const data = sPresensi.getDataRange().getValues();
  const attendanceMap = {};
  let lastDetectedTeacher = "";

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const c0 = String(row[0] || "").trim();

    // 1. Deteksi nama guru di atas tabel (baris judul sebelum tabel rincian)
    if (c0.length >= 4 && !/^(presensi|kabupaten|bulan|id|no\.|hari|maros|kepala)/i.test(c0) && isNaN(c0)) {
      if (c0.indexOf("31 Oktober") === -1 && c0.indexOf("2026") === -1) {
        lastDetectedTeacher = c0;
      }
    }

    // 2. Deteksi header 'TOTAL KEHADIRAN' dan ambil angka kehadiran tepat di bawahnya
    for (let c = 0; c < row.length; c++) {
      const headerText = String(row[c] || "").trim().toUpperCase();
      if (headerText === "TOTAL KEHADIRAN" || headerText.indexOf("TOTAL KEHADIRAN") !== -1) {
        if (r + 1 < data.length) {
          const rawVal = data[r + 1][c];
          if (rawVal !== "" && rawVal !== null && !isNaN(rawVal)) {
            const totalNum = Number(rawVal);
            let teacherName = lastDetectedTeacher;

            // Jika belum dapat dari baris atas, cek kolom NAMA PENGAJAR di baris data r+1
            if (!teacherName && row.length > 4) {
              const possibleName = String(data[r + 1][4] || "").trim();
              if (possibleName && isNaN(possibleName) && !/^(tepat|terlambat|tidak)/i.test(possibleName)) {
                teacherName = possibleName;
              }
            }

            if (teacherName) {
              attendanceMap[normalizeTeacherName(teacherName)] = totalNum;
            }
          }
        }
      }
    }

    // 3. Cadangan: Rekap kolom ringkasan (jika ada kolom nama pengajar & total kehadiran di kanan)
    for (let c = 0; c < row.length - 1; c++) {
      const nameCell = String(row[c] || "").trim();
      const numCell = row[c + 1];
      if (nameCell.length >= 4 && typeof numCell === "number" && numCell >= 0 && numCell <= 31) {
        if (!/^(senin|selasa|rabu|kamis|jum'?at|sabtu|ahad|id|no|total|hari|tepat|terlambat|tidak)/i.test(nameCell)) {
          const cleanKey = normalizeTeacherName(nameCell);
          if (attendanceMap[cleanKey] === undefined) {
            attendanceMap[cleanKey] = numCell;
          }
        }
      }
    }
  }

  return attendanceMap;
}

/**
 * Normalisasi nama guru untuk pencocokan toleran terhadap gelar, spasi, dan variasi ejaan.
 */
function normalizeTeacherName(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/[^a-z]/g, "")
    .replace(/dh/g, "d")
    .replace(/th/g, "t")
    .replace(/kh/g, "k")
    .replace(/h/g, "");
}

/**
 * Mencari jumlah kehadiran guru berdasarkan nama dengan pencocokan cerdas.
 */
function cariKehadiranGuru(targetName, attendanceMap) {
  const normTarget = normalizeTeacherName(targetName);
  if (attendanceMap[normTarget] !== undefined) {
    return attendanceMap[normTarget];
  }

  for (const k in attendanceMap) {
    if (normTarget.indexOf(k) !== -1 || k.indexOf(normTarget) !== -1) {
      return attendanceMap[k];
    }
    const subTarget = normTarget.slice(0, 8);
    const subK = k.slice(0, 8);
    if (subTarget.length >= 6 && subK.length >= 6 && subTarget === subK) {
      return attendanceMap[k];
    }
  }

  return null;
}

/**
 * Fitur Tutup Buku & Arsip Bulanan:
 * 1. Menanyakan nama periode bulan (misal: Oktober 2026)
 * 2. Menyalin sheet Master_Gaji, KBM, Ujian, dan Presensi ke tab arsip permanen (sebagai values murni)
 * 3. Mengosongkan data baris KBM dan Ujian (baris 2 ke bawah) agar aplikasi kembali bersih untuk bulan baru
 * 4. Mereset kolom Kehadiran Transport di Master_Gaji menjadi 0
 * 5. Menjaga Master Guru, PIN, Mapel, Kelas, dan sheet Tarif tetap 100% utuh dan aman
 */
function tutupBukuBulanan() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.prompt(
    "📦 Tutup Buku & Arsip Bulanan",
    "Masukkan nama bulan & tahun untuk diarsipkan (contoh: Oktober 2026):\n\n" +
    "Data bulan ini akan otomatis diarsipkan secara permanen, lalu sheet KBM & Ujian akan dikosongkan agar siap untuk bulan berikutnya.",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const periode = response.getResponseText().trim();
  if (!periode) {
    ui.alert("⚠️ Nama periode tidak boleh kosong. Proses dibatalkan.");
    return;
  }

  // Format label aman untuk nama sheet (misal 'Okt_2026' atau 'Oktober_2026')
  const tagPeriode = periode.replace(/[^a-zA-Z0-9]/g, "_");

  // 1. Arsipkan Sheet Master_Gaji (Simpan sebagai Nilai Murni / Values agar rekap gaji bulan ini terkunci permanen)
  const sMasterGaji = ss.getSheetByName("Master_Gaji");
  if (sMasterGaji && sMasterGaji.getLastRow() >= 2) {
    const archiveGajiName = "Gaji_" + tagPeriode;
    let sArchGaji = ss.getSheetByName(archiveGajiName);
    if (sArchGaji) ss.deleteSheet(sArchGaji);
    sArchGaji = sMasterGaji.copyTo(ss);
    sArchGaji.setName(archiveGajiName);
    const range = sArchGaji.getDataRange();
    range.setValues(range.getValues());
  }

  // 2. Arsipkan Sheet KBM
  const sKBM = ss.getSheetByName(SHEET_KBM);
  let jmlKBM = 0;
  if (sKBM && sKBM.getLastRow() >= 2) {
    jmlKBM = sKBM.getLastRow() - 1;
    const archiveKBMName = "KBM_" + tagPeriode;
    let sArchKBM = ss.getSheetByName(archiveKBMName);
    if (sArchKBM) ss.deleteSheet(sArchKBM);
    sArchKBM = sKBM.copyTo(ss);
    sArchKBM.setName(archiveKBMName);
    const range = sArchKBM.getDataRange();
    range.setValues(range.getValues());

    // Kosongkan baris 2 ke bawah
    sKBM.deleteRows(2, jmlKBM);
  }

  // 3. Arsipkan Sheet Ujian
  const sUjian = ss.getSheetByName(SHEET_UJIAN);
  let jmlUjian = 0;
  if (sUjian && sUjian.getLastRow() >= 2) {
    jmlUjian = sUjian.getLastRow() - 1;
    const archiveUjianName = "Ujian_" + tagPeriode;
    let sArchUjian = ss.getSheetByName(archiveUjianName);
    if (sArchUjian) ss.deleteSheet(sArchUjian);
    sArchUjian = sUjian.copyTo(ss);
    sArchUjian.setName(archiveUjianName);
    const range = sArchUjian.getDataRange();
    range.setValues(range.getValues());

    // Kosongkan baris 2 ke bawah
    sUjian.deleteRows(2, jmlUjian);
  }

  // 4. Arsipkan Sheet Presensi jika ada
  const sPresensi = ss.getSheetByName("Presensi");
  if (sPresensi && sPresensi.getLastRow() >= 2) {
    const archivePresensiName = "Presensi_" + tagPeriode;
    let sArchPresensi = ss.getSheetByName(archivePresensiName);
    if (sArchPresensi) ss.deleteSheet(sArchPresensi);
    sArchPresensi = sPresensi.copyTo(ss);
    sArchPresensi.setName(archivePresensiName);
    const range = sArchPresensi.getDataRange();
    range.setValues(range.getValues());
  }

  // 5. Reset Kehadiran Presensi di Master_Gaji (Kolom R) menjadi 0 untuk bulan baru
  if (sMasterGaji && sMasterGaji.getLastRow() >= 2) {
    sMasterGaji.getRange(2, 18, sMasterGaji.getLastRow() - 1, 1).setValue(0);
  }

  // 6. Sinkronkan kembali rumus
  sinkronkanMasterGaji();

  ui.alert(
    "✅ Tutup Buku Berhasil!",
    `Laporan periode "${periode}" telah berhasil diarsipkan ke spreadsheet:\n\n` +
    `📁 Gaji_${tagPeriode} (Rekap Gaji Permanen)\n` +
    `📁 KBM_${tagPeriode} (${jmlKBM} Data Jam Mengajar)\n` +
    `📁 Ujian_${tagPeriode} (${jmlUjian} Data Honor Ujian)\n\n` +
    `Sheet KBM dan Ujian telah dikosongkan. Aplikasi SiRekap di HP Guru otomatis bersih kembali dan siap digunakan untuk bulan berikutnya! 🟢`,
    ui.ButtonSet.OK
  );
}

/**
 * Menu otomatis di Google Spreadsheet: memudahkan admin menjalankan fungsi tanpa buka Apps Script
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 SiRekap")
    .addItem("🖨️ Inisialisasi Ulang Desain Slip Gaji (Reset Layout)", "setupSlipGaji")
    .addItem("⚙️ Siapkan Sheet Tarif (Tabel Pengaturan)", "setupSheetTarif")
    .addItem("🔢 Urutkan Ulang Nomor Kode Guru (Otomatis)", "urutkanUlangKodeGuru")
    .addItem("📥 Tarik Kehadiran dari Sheet Presensi", "tarikKehadiranPresensi")
    .addItem("🔄 Perbarui Rumus Slip Gaji & Sinkronkan Ujian", "sinkronkanMasterGaji")
    .addSeparator()
    .addItem("📦 Tutup Buku & Arsipkan Data Bulanan (Bulan Baru)", "tutupBukuBulanan")
    .addToUi();
}
