/**
 * ====================================================================
 * APLIKASI PRESENSI BARCODE & GENERATOR KARTU IDENTITAS
 * Backend Google Apps Script (Code.gs)
 * ====================================================================
 */

// Menampilkan antarmuka HTML utama saat Web App diakses via browser
function doGet(e) {
  // Jika diakses tanpa parameter khusus, tampilkan antarmuka web
  if (!e || !e.parameter || !e.parameter.id) {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('PresensiKita Pro - Presensi Barcode & Kartu Pelajar')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Jika diakses dengan query parameter ID (Fallback mode HTTP GET)
  return simpanDataPresensi(e.parameter);
}

// Menangani sinkronisasi data presensi secara otomatis melalui HTTP POST
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    return simpanDataPresensi(data);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "error": err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Merekam baris data presensi ke dalam Spreadsheet aktif
 */
function simpanDataPresensi(data) {
  // Menggunakan LockService untuk menangani antrean saat banyak siswa scan secara bersamaan
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Batas tunggu maksimal 10 detik
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "error": "Sistem sibuk, silakan coba lagi beberapa saat."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Spreadsheet tidak ditemukan atau belum ditautkan!");
    }
    
    // Gunakan sheet pertama / aktif
    var sheet = ss.getActiveSheet();
    
    // Jika sheet masih kosong, buatkan Header Otomatis
    if (sheet.getLastRow() === 0) {
      var headerRow = [
        "Tanggal & Waktu", 
        "Jenis Presensi", 
        "Kode ID / NISN", 
        "Nama Lengkap", 
        "Peran", 
        "Kelas / Jabatan", 
        "Status Kehadiran"
      ];
      sheet.appendRow(headerRow);
      sheet.getRange(1, 1, 1, headerRow.length)
        .setFontWeight("bold")
        .setBackground("#e2e8f0")
        .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // Ambil variabel parameter dari request
    var timestamp = data.timestamp || new Date().toLocaleString("id-ID");
    var presensiType = data.presensiType || "Masuk";
    var id = data.id || "-";
    var name = data.name || "-";
    var role = data.role || "-";
    var department = data.department || "-";
    var status = data.status || "-";

    // Sisipkan baris data presensi di urutan terbawah
    sheet.appendRow([timestamp, presensiType, id, name, role, department, status]);

    return ContentService.createTextOutput(JSON.stringify({
      "result": "success",
      "message": "Data presensi berhasil dicatat",
      "data": data
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "error": err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper function untuk menyertakan berkas terpisah jika diperlukan di kemudian hari
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
