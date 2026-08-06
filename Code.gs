/**
 * AutoFace.Sekolah - Google Apps Script Webhook Handler
 * Integrasi Real-Time Pemindai Wajah Otomatis dengan Google Sheets
 */

// Nama Lembar Kerja / Sheet
var SHEET_NAME_LOG = "Log Presensi";
var SHEET_NAME_MASTER = "Master Data";

/**
 * Memproses permintaan POST yang dikirim oleh Aplikasi Presensi
 */
function doPost(e) {
  // Gunakan LockService untuk mencegah konflik jika banyak orang presensi bersamaan
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Tunggu maksimal 10 detik
  } catch (e) {
    return createJsonResponse({
      status: "error",
      message: "Server sibuk, silakan coba beberapa saat lagi."
    });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        status: "error",
        message: "Data kiriman kosong atau tidak valid."
      });
    }

    // Parse data JSON yang dikirimkan oleh web application
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var payload = data.payload;

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Aksi pencatatan presensi
    if (action === "record_attendance") {
      var sheet = getOrCreateSheet(ss, SHEET_NAME_LOG, [
        "ID Log",
        "NIP / NISN",
        "Nama Lengkap",
        "Kategori",
        "Jabatan / Kelas",
        "Tanggal",
        "Jam Masuk",
        "Jam Pulang",
        "Status",
        "Waktu Sync Backend"
      ]);

      var resultMessage = processAttendanceRecord(sheet, payload);
      return createJsonResponse({
        status: "success",
        message: resultMessage,
        data: payload
      });
    }

    return createJsonResponse({
      status: "error",
      message: "Aksi tidak dikenal: " + action
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: "Terjadi kesalahan server: " + error.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memproses permintaan GET untuk cek status apakah Webhook aktif
 */
function doGet(e) {
  return ContentService.createTextOutput("🟢 Webhook AutoFace.Sekolah Google Apps Script Aktif & Siap Digunakan!")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Logika pencatatan & pembaruan data presensi pada Sheet
 */
function processAttendanceRecord(sheet, payload) {
  var data = sheet.getDataRange().getValues();
  var nipTarget = String(payload.nip || "").trim();
  var tanggalTarget = String(payload.tanggal || "").trim();
  
  var existingRowIndex = -1;

  // Cari apakah sudah ada catatan NIP/NISN pada Tanggal yang sama
  for (var i = 1; i < data.length; i++) {
    var rowNip = String(data[i][1]).trim();
    var rowTanggal = data[i][5];

    // Format tanggal sheet jika bertipe Date
    if (rowTanggal instanceof Date) {
      rowTanggal = Utilities.formatDate(rowTanggal, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      rowTanggal = String(rowTanggal).trim();
    }

    if (rowNip === nipTarget && rowTanggal === tanggalTarget) {
      existingRowIndex = i + 1; // Baris di Apps Script berbasis 1
      break;
    }
  }

  var currentTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  // Jika baris sudah ada (Pembaruan jam pulang atau jam masuk)
  if (existingRowIndex > 0) {
    if (payload.jamKeluar && payload.jamKeluar !== "-") {
      sheet.getRange(existingRowIndex, 8).setValue(payload.jamKeluar); // Kolom Jam Pulang (H)
    }
    if (payload.jamMasuk && payload.jamMasuk !== "-") {
      sheet.getRange(existingRowIndex, 7).setValue(payload.jamMasuk); // Kolom Jam Masuk (G)
    }
    if (payload.status) {
      sheet.getRange(existingRowIndex, 9).setValue(payload.status); // Kolom Status (I)
    }
    sheet.getRange(existingRowIndex, 10).setValue(currentTimeStr); // Kolom Waktu Sync (J)

    return "Data presensi untuk " + payload.nama + " berhasil diperbarui di baris " + existingRowIndex;
  } 
  // Jika baris belum ada (Tambah data presensi baru)
  else {
    sheet.appendRow([
      payload.id || Date.now().toString(),
      payload.nip,
      payload.nama,
      payload.kategori,
      payload.jabatan,
      payload.tanggal,
      payload.jamMasuk || "-",
      payload.jamKeluar || "-",
      payload.status || "Tepat Waktu",
      currentTimeStr
    ]);

    return "Data presensi baru untuk " + payload.nama + " berhasil dicatat.";
  }
}

/**
 * Membuat sheet baru dan format header otomatis jika belum ada
 */
function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    // Tambahkan baris header
    sheet.appendRow(headers);
    
    // Format visual header
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#22c55e"); // Warna Hijau Emerald
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);

    // Atur lebar kolom otomatis
    for (var col = 1; col <= headers.length; col++) {
      sheet.setColumnWidth(col, 150);
    }
  }
  return sheet;
}

/**
 * Helper untuk mengembalikan respon berformat JSON
 */
function createJsonResponse(dataObject) {
  return ContentService.createTextOutput(JSON.stringify(dataObject))
    .setMimeType(ContentService.MimeType.JSON);
}