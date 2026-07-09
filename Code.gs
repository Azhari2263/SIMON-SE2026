/**
 * SIMON-PERFORMA: Google Apps Script Backend
 * Konfigurasi & Otorisasi Google Sheets dan Google Drive
 */

const FOLDER_DRIVE_ID = "1ihZrsWTrJp-wwmpA-CMCw6UREsTNQLaw"; // TODO: Ubah dengan ID Folder Drive Anda
const SHEET_REF_NAME = "REF_DATA";
const SHEET_HISTORY_NAME = "RIWAYAT_UPLOAD";
const SHEET_MONITORING_NAME = "DATA_MONITORING";
const SHEET_REPORT_NAME = "LAPORAN_TERKINI";

// Kolom Wajib untuk Data PPL
const REQUIRED_COLUMNS = [
  "Nama PPL", "PML", "PJ Kuda", "Kab/Kota", "OPEN", "SUBMITTED BY Pencacah",
  "DRAFT", "APPROVED BY Pengawas", "REJECTED BY Pengawas", "SUBMITTED RESPONDENT",
  "REVOKED BY Pengawas", "EDITED BY Pengawas", "Realisasi Tanpa Draft",
  "Realisasi Dengan Draft", "Rata-Rata/Hari Tanpa Draft", "Rata-Rata/Hari Dengan Draft"
];

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Simon SE2026 - Monitoring Kinerja PPL')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Menguji koneksi ke Google Spreadsheet
 */
function testConnection() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      return { success: true, name: ss.getName() };
    }
    return { success: false, error: "Spreadsheet tidak terdeteksi atau tidak aktif." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Memformat objek Date ke format Indonesia (e.g. "9 Juli 2026")
 */
function formatIndonesianDate(date) {
  const months = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  // Pastikan konversi date ke timezone Jakarta (GMT+7)
  const dateStr = Utilities.formatDate(date, "Asia/Jakarta", "yyyy-MM-dd");
  const parts = dateStr.split("-"); // [yyyy, MM, dd]
  const day = parseInt(parts[2], 10);
  const monthIdx = parseInt(parts[1], 10) - 1;
  const year = parts[0];
  
  return day + " " + months[monthIdx] + " " + year;
}

/**
 * Memeriksa ketersediaan data referensi hari sebelumnya
 */
/**
 * Membaca teks tanggal bahasa Indonesia (e.g. "9 Juli 2026") menjadi Date
 */
function parseIndonesianDate(dateStr) {
  const months = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  const parts = String(dateStr).trim().split(" ");
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const monthIdx = months.indexOf(parts[1]);
    const year = parseInt(parts[2], 10);
    if (monthIdx !== -1 && !isNaN(day) && !isNaN(year)) {
      return new Date(year, monthIdx, day);
    }
  }
  return null;
}

/**
 * Mengambil tanggal upload terakhir dari DATA_MONITORING
 */
function getLatestMonitoringDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MONITORING_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return null; // Belum ada data
  }
  
  const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  let latestDate = null;
  for (let i = 0; i < dates.length; i++) {
    const dStr = dates[i][0].trim();
    if (!dStr) continue;
    const parsed = parseIndonesianDate(dStr);
    if (parsed) {
      if (!latestDate || parsed > latestDate) {
        latestDate = parsed;
      }
    }
  }
  return latestDate;
}

/**
 * Memeriksa ketersediaan data referensi hari sebelumnya
 */
/**
 * Mengambil data monitoring dari DATA_MONITORING berdasarkan tanggal tertentu
 */
function getMonitoringDataByDate(dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MONITORING_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  const dateDisplayValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const restValues = sheet.getRange(2, 2, lastRow - 1, REQUIRED_COLUMNS.length).getValues();
  
  const results = [];
  for (let i = 0; i < dateDisplayValues.length; i++) {
    const rowDateStr = dateDisplayValues[i][0].trim();
    if (rowDateStr === dateString) {
      const obj = {};
      REQUIRED_COLUMNS.forEach((col, idx) => {
        obj[col] = restValues[i][idx];
      });
      results.push(obj);
    }
  }
  return results;
}

/**
 * Memindai log aktivitas untuk mengambil nama file terakhir yang diunggah untuk tanggal target tertentu
 */
function getLatestUploadedFilenames(yesterdayStr, todayStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_HISTORY_NAME);
    let yesterdayFile = "";
    let todayFile = "";
    
    if (sheet && sheet.getLastRow() > 1) {
      const lastRow = sheet.getLastRow();
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      // Pindai dari data terbaru ke terlama
      for (let i = data.length - 1; i >= 0; i--) {
        const type = String(data[i][1] || "");
        const file = String(data[i][2] || "");
        
        // Cek file H-1
        if (!yesterdayFile && type.indexOf("H-1") !== -1 && type.indexOf(yesterdayStr) !== -1) {
          yesterdayFile = file;
        }
        // Cek file H
        if (!todayFile && (type.indexOf("Hari Ini") !== -1 || type.indexOf("Kalkulasi") !== -1) && type.indexOf(todayStr) !== -1) {
          todayFile = file;
        }
        if (yesterdayFile && todayFile) break;
      }
    }
    return {
      yesterdayFile: yesterdayFile || "Data Tersimpan",
      todayFile: todayFile || "Data Tersimpan"
    };
  } catch (err) {
    return { yesterdayFile: "Data Tersimpan", todayFile: "Data Tersimpan" };
  }
}

/**
 * Memeriksa ketersediaan data referensi hari sebelumnya dan data terbaru hari ini
 */
function checkReferenceData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ambil tanggal upload terakhir dari DATA_MONITORING
    const latestDate = getLatestMonitoringDate();
    
    // Hitung tanggal target H-1 dan H secara default (berdasarkan server saat ini)
    const now = new Date();
    let todayDate = now;
    let yesterdayDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    
    const serverTodayStr = formatIndonesianDate(now);
    const serverYesterdayStr = formatIndonesianDate(yesterdayDate);
    
    let hasData = false;
    let hasTodayData = false;
    let totalCount = 0;
    let todayCount = 0;
    
    if (latestDate) {
      const latestDateStr = formatIndonesianDate(latestDate);
      
      // Jika latestDateStr adalah hari ini, berarti hari ini sudah pernah upload.
      // Target upload hari ini tetap H = serverTodayStr dan H-1 = serverYesterdayStr.
      if (latestDateStr === serverTodayStr) {
        yesterdayDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
        todayDate = now;
        
        // Periksa apakah baseline kemarin (yesterdayStr) ada di DATA_MONITORING
        const yesterdayStr = formatIndonesianDate(yesterdayDate);
        const yesterdayRows = getMonitoringDataByDate(yesterdayStr);
        if (yesterdayRows.length > 0) {
          hasData = true;
          totalCount = yesterdayRows.length;
        } else {
          hasData = false;
        }
        
        // Periksa apakah data hari ini (H) juga sudah ada
        const todayRows = getMonitoringDataByDate(serverTodayStr);
        if (todayRows.length > 0) {
          hasTodayData = true;
          todayCount = todayRows.length;
        }
      } else if (latestDateStr === serverYesterdayStr) {
        // Normal: data kemarin ada di database. User tinggal upload data hari ini.
        yesterdayDate = latestDate;
        todayDate = now;
        
        const yesterdayRows = getMonitoringDataByDate(latestDateStr);
        hasData = true;
        totalCount = yesterdayRows.length;
        
        hasTodayData = false;
      } else {
        // Terjadi jeda/gap!
        const twoDaysAgo = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
        const twoDaysAgoStr = formatIndonesianDate(twoDaysAgo);
        
        if (latestDateStr === twoDaysAgoStr) {
          // Hanya kemarin (H-1) yang kosong. User harus upload H-1 dan H.
          yesterdayDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
          todayDate = now;
          hasData = false;
          hasTodayData = false;
        } else {
          // Jeda > 1 hari. Geser target upload secara kronologis.
          yesterdayDate = latestDate;
          todayDate = new Date(latestDate.getTime() + 24 * 60 * 60 * 1000);
          hasData = true;
          
          const yesterdayRows = getMonitoringDataByDate(latestDateStr);
          totalCount = yesterdayRows.length;
          
          // Periksa apakah target H juga sudah ada di DB
          const targetTodayStr = formatIndonesianDate(todayDate);
          const todayRows = getMonitoringDataByDate(targetTodayStr);
          if (todayRows.length > 0) {
            hasTodayData = true;
            todayCount = todayRows.length;
          }
        }
      }
    } else {
      // Database kosong
      hasData = false;
      hasTodayData = false;
    }
    
    const todayStr = formatIndonesianDate(todayDate);
    const yesterdayStr = formatIndonesianDate(yesterdayDate);
    const filenames = getLatestUploadedFilenames(yesterdayStr, todayStr);
    
    return {
      hasData: hasData,
      hasTodayData: hasTodayData,
      todayStr: todayStr,
      yesterdayStr: yesterdayStr,
      latestDateStr: latestDate ? formatIndonesianDate(latestDate) : "Belum ada",
      totalCount: totalCount,
      todayCount: todayCount,
      yesterdayFileName: filenames.yesterdayFile,
      todayFileName: filenames.todayFile
    };
  } catch (e) {
    return { hasData: false, hasTodayData: false, error: e.toString() };
  }
}

/**
 * Menyimpan data riwayat upload ke sheet RIWAYAT_UPLOAD
 */
function logHistory(fileName, type, userEmail = "User") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_HISTORY_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HISTORY_NAME);
    sheet.appendRow(["Tanggal & Waktu", "Jenis Aksi", "Nama File", "Pengguna", "Status"]);
  }
  const timestamp = new Date();
  sheet.appendRow([timestamp, type, fileName, userEmail, "SUKSES"]);
}

/**
 * Membersihkan log riwayat aktivitas di Google Sheets
 */
function clearHistoryLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_HISTORY_NAME);
    if (sheet) {
      sheet.clearContents();
      sheet.appendRow(["Tanggal & Waktu", "Jenis Aksi", "Nama File", "Pengguna", "Status"]);
    }
    logHistory("-", "Pembersihan Log", Session.getActiveUser().getEmail() || "admin@bps.go.id");
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Menyimpan/memperbarui data monitoring ke sheet DATA_MONITORING berdasarkan tanggal secara cepat (in-memory)
 */
function saveMonitoringData(dataObj, dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MONITORING_NAME);
  
  const headers = ["Tanggal"].concat(REQUIRED_COLUMNS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MONITORING_NAME);
    sheet.appendRow(headers);
  }
  
  const lastRow = sheet.getLastRow();
  const rowsToWrite = [];
  
  // Filter in-memory data lama yang tanggalnya TIDAK sama dengan dateString
  if (lastRow > 1) {
    const allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < allData.length; i++) {
      // Cek kolom pertama (Tanggal)
      const rowDateStr = String(allData[i][0]).trim();
      if (rowDateStr !== dateString) {
        rowsToWrite.push(allData[i]);
      }
    }
  }
  
  // Tambahkan data baru ke array in-memory
  if (dataObj && dataObj.length > 0) {
    dataObj.forEach(row => {
      const rowData = [dateString];
      REQUIRED_COLUMNS.forEach(col => {
        rowData.push(row[col] !== undefined ? row[col] : "");
      });
      rowsToWrite.push(rowData);
    });
  }
  
  // Kosongkan area data lama (di bawah header)
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  
  // Tulis kembali data gabungan dalam satu kali operasi batch
  if (rowsToWrite.length > 0) {
    sheet.getRange(2, 1, rowsToWrite.length, headers.length).setValues(rowsToWrite);
  }
}

/**
 * Menyimpan data laporan terkini ke sheet LAPORAN_TERKINI
 */
function saveReportData(results) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_REPORT_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REPORT_NAME);
  }
  sheet.clearContents();
  if (results && results.length > 0) {
    const headers = Object.keys(results[0]);
    sheet.appendRow(headers);
    const rows = results.map(row => headers.map(h => row[h]));
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

/**
 * Mengambil semua data inisialisasi awal untuk meminimalisasi waktu render client
 */
function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Cek baseline status
    const refStatus = checkReferenceData();
    
    // 2. Ambil laporan terakhir dari LAPORAN_TERKINI
    let currentReport = [];
    const reportSheet = ss.getSheetByName(SHEET_REPORT_NAME);
    if (reportSheet && reportSheet.getLastRow() > 1) {
      const headers = reportSheet.getRange(1, 1, 1, reportSheet.getLastColumn()).getValues()[0];
      const dataRows = reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, reportSheet.getLastColumn()).getValues();
      currentReport = dataRows.map(row => {
        let obj = {};
        headers.forEach((h, idx) => obj[h] = row[idx]);
        return obj;
      });
    }
    
    // 3. Ambil log riwayat dari RIWAYAT_UPLOAD (maks 100 terakhir)
    let history = [];
    const historySheet = ss.getSheetByName(SHEET_HISTORY_NAME);
    if (historySheet && historySheet.getLastRow() > 1) {
      const headers = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
      const lastRow = historySheet.getLastRow();
      const startRow = Math.max(2, lastRow - 99);
      const numRows = lastRow - startRow + 1;
      const dataRows = historySheet.getRange(startRow, 1, numRows, historySheet.getLastColumn()).getValues();
      
      history = dataRows.map(row => {
        let obj = {};
        headers.forEach((h, idx) => {
          let val = row[idx];
          if (val instanceof Date) {
            obj[h] = formatTimestamp(val);
          } else {
            obj[h] = val;
          }
        });
        return obj;
      });
      history.reverse(); // Urutan terbaru di atas
    }
    
    return {
      success: true,
      refStatus: refStatus,
      currentReport: currentReport,
      history: history
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Helper untuk format timestamp log riwayat
function formatTimestamp(date) {
  const pad = (n) => String(n).padStart(2, '0');
  const d = pad(date.getDate());
  const m = pad(date.getMonth() + 1);
  const y = date.getFullYear();
  const h = pad(date.getHours());
  const min = pad(date.getMinutes());
  const s = pad(date.getSeconds());
  return `${d}-${m}-${y} ${h}:${min}:${s}`;
}

/**
 * Memproses berkas unggahan, membandingkan data, dan menyimpan hasil ke Spreadsheet.
 */
function processUploads(yesterdayDataObj, todayDataObj, filenames, yesterdayStr, todayStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const refSheet = ss.getSheetByName(SHEET_REF_NAME);
    const userEmail = Session.getActiveUser().getEmail() || "Pengguna Aplikasi";

    let baselineData = [];
    if (yesterdayDataObj && yesterdayDataObj.length > 0) {
      // Validasi kolom file kemarin
      validateHeaders(yesterdayDataObj[0], REQUIRED_COLUMNS);
      
      // Simpan berkas kemarin ke DATA_MONITORING dengan tanggal kemarin (e.g. 8 Juli 2026)
      saveMonitoringData(yesterdayDataObj, yesterdayStr);
      
      baselineData = yesterdayDataObj;
      logHistory(filenames.yesterday, "Upload File H-1 (" + yesterdayStr + ")", userEmail);
    } else {
      // Ambil baseline langsung dari DATA_MONITORING berdasarkan tanggal kemarin
      baselineData = getMonitoringDataByDate(yesterdayStr);
      if (baselineData.length === 0) {
        throw new Error("Data baseline tanggal " + yesterdayStr + " tidak ditemukan di database. Silakan upload berkas H-1 terlebih dahulu.");
      }
    }

    // Validasi kolom file hari ini
    if (!todayDataObj || todayDataObj.length === 0) {
      throw new Error("Data hari ini kosong atau tidak valid.");
    }
    validateHeaders(todayDataObj[0], REQUIRED_COLUMNS);

    // Simpan berkas hari ini ke DATA_MONITORING dengan tanggal hari ini (e.g. 9 Juli 2026)
    saveMonitoringData(todayDataObj, todayStr);

    // Proses Kalkulasi Perbandingan
    const results = [];
    const yesterdayMap = new Map();
    baselineData.forEach(row => {
      const key = (row["Nama PPL"] || "").toString().trim() + "|" + (row["Kab/Kota"] || "").toString().trim();
      yesterdayMap.set(key, Number(row["DRAFT"]) || 0);
    });

    todayDataObj.forEach(row => {
      const pplName = (row["Nama PPL"] || "").toString().trim();
      const kab = (row["Kab/Kota"] || "").toString().trim();
      const key = pplName + "|" + kab;
      
      const draftToday = Number(row["DRAFT"]) || 0;
      const draftYesterday = yesterdayMap.has(key) ? yesterdayMap.get(key) : 0;
      
      // Rumus: Perubahan = DRAFT Hari Ini - DRAFT Hari Sebelumnya
      const perubahan = draftToday - draftYesterday;

      results.push({
        "Nama PPL": pplName,
        "PML": row["PML"] || "-",
        "PJ Kuda": row["PJ Kuda"] || "-",
        "Kab/Kota": kab,
        "OPEN": Number(row["OPEN"]) || 0,
        "SUBMITTED BY Pencacah": Number(row["SUBMITTED BY Pencacah"]) || 0,
        "DRAFT": draftToday,
        "Rata-Rata/Hari Tanpa Draft": Number(row["Rata-Rata/Hari Tanpa Draft"]) || 0,
        "Rata-Rata/Hari Dengan Draft": Number(row["Rata-Rata/Hari Dengan Draft"]) || 0,
        "Perubahan": perubahan
      });
    });

    // Urutkan berdasarkan DRAFT hari ini (descending)
    results.sort((a, b) => b["DRAFT"] - a["DRAFT"]);

    // Update Referensi Database (REF_DATA) dengan data baseline (kemarin) yang digunakan
    if (refSheet && baselineData.length > 0) {
      refSheet.clearContents();
      const refHeaders = Object.keys(baselineData[0]);
      refSheet.appendRow(refHeaders);
      const refRows = baselineData.map(row => refHeaders.map(h => row[h]));
      refSheet.getRange(2, 1, refRows.length, refHeaders.length).setValues(refRows);
    }

    // Simpan hasil kalkulasi laporan ke LAPORAN_TERKINI
    saveReportData(results);

    logHistory(filenames.today, "Upload File Hari Ini (" + todayStr + ") & Kalkulasi", userEmail);

    // Simpan backup file asli ke Google Drive (diletakkan di akhir secara opsional agar tidak menghentikan perekaman database utama)
    try {
      if (todayDataObj && todayDataObj.length > 0) {
        saveToDrive(todayStr + "_Data Monitoring Kuda", todayDataObj, FOLDER_DRIVE_ID);
      }
      if (yesterdayDataObj && yesterdayDataObj.length > 0) {
        saveToDrive(yesterdayStr + "_Data Monitoring Kuda", yesterdayDataObj, FOLDER_DRIVE_ID);
      }
    } catch (driveErr) {
      Logger.log("Gagal memproses penyimpanan backup Drive: " + driveErr.toString());
    }

    return {
      success: true,
      data: results,
      totalPPL: results.length
    };

  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Validasi header kolom apakah lengkap sesuai syarat
 */
function validateHeaders(rowObj, requiredCols) {
  const fileCols = Object.keys(rowObj);
  const missing = requiredCols.filter(c => !fileCols.includes(c));
  if (missing.length > 0) {
    throw new Error("Format file salah. Kolom berikut tidak ditemukan: " + missing.join(", "));
  }
}

/**
 * Menyimpan backup file Excel asli ke folder Drive yang ditentukan
 */
function saveToDrive(filename, dataObj, folderId) {
  try {
    // Abaikan jika ID folder kosong, masih placeholder default, atau menggunakan ID dummy bawaan
    if (!folderId || folderId === "MASUKKAN_ID_FOLDER_GOOGLE_DRIVE_DISINI" || folderId === "1ihZrsWTrJp-wwmpA-CMCw6UREsTNQLaw") {
      Logger.log("Penyimpanan Google Drive diabaikan: ID folder dummy atau belum terkonfigurasi.");
      return;
    }
    const folder = DriveApp.getFolderById(folderId);
    const headers = Object.keys(dataObj[0]);
    const csvContent = [
      headers.join(","),
      ...dataObj.map(row => headers.map(h => '"' + (row[h] || '').toString().replace(/"/g, '""') + '"').join(","))
    ].join("\n");
    
    // Simpan file dengan ekstensi .csv
    folder.createFile(filename + ".csv", csvContent, MimeType.CSV);
  } catch (e) {
    Logger.log("Gagal backup Drive: " + e.toString());
  }
}

/**
 * Endpoint API POST untuk menangani pemanggilan dari luar container Google Apps Script (CORS)
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const args = postData.arguments || [];
    
    let result;
    if (action === "testConnection") {
      result = testConnection();
    } else if (action === "getInitialData") {
      result = getInitialData();
    } else if (action === "clearHistoryLog") {
      result = clearHistoryLog();
    } else if (action === "processUploads") {
      // args: [yesterdayDataObj, todayDataObj, filenames, yesterdayStr, todayStr]
      result = processUploads(args[0], args[1], args[2], args[3], args[4]);
    } else {
      result = { success: false, error: "Aksi tidak dikenali: " + action };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
