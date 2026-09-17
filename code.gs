/************************************************************
 * RSVP & Ucapan - Heni & Rama
 * Google Apps Script - ditempel di Extensions > Apps Script
 * Spreadsheet yang SAMA dengan FORM CATIN tidak masalah,
 * script ini memakai sheet khusus bernama "RSVP".
 *
 * Kolom sheet "RSVP" (dibuat otomatis oleh setupSheet):
 *   A: Timestamp | B: Nama | C: Kehadiran | D: Jumlah Tamu | E: Ucapan
 *
 * CARA PAKAI:
 * 1. Ganti SHEET_ID dengan ID spreadsheet (bagian di URL /d/.../edit).
 * 2. Jalankan fungsi setupSheet() sekali (Run > setupSheet).
 * 3. Deploy > New deployment > Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy Web App URL ke index.html sebagai RSVP_API_URL.
 ************************************************************/

var SHEET_ID = "1DywdSjYJP9evu4m77rsuzwzIi9-GPfl3UNlvTCjVam8";
var SHEET_NAME = "Sheet1";

/** Buat sheet + header sekali saja. */
function setupSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Nama", "Kehadiran", "Jumlah Tamu", "Ucapan"]);
  }
}

/** Hapus SEMUA data ucapan (baris 2 ke bawah), header tetap. Jalankan sekali untuk membersihkan. */
function clearWishes() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

/** Ambil sheet, buat jika belum ada. */
function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Nama", "Kehadiran", "Jumlah Tamu", "Ucapan"]);
  }
  return sheet;
}

/** Balas JSON (Apps Script otomatis mengizinkan GET lintas origin). */
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET ?action=list&limit=50
 * Dipakai web untuk MEMUNCULKAN daftar ucapan + kolom lainnya.
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || "list";
    if (action === "list") {
      var limit = parseInt((e.parameter && e.parameter.limit) || "50", 10) || 50;
      var sheet = getSheet_();
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return jsonOut_({ status: "ok", data: [] });
      }
      var startRow = Math.max(2, lastRow - limit + 1);
      var values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 5).getValues();
      var data = [];
      for (var i = values.length - 1; i >= 0; i--) {
        data.push({
          waktu: values[i][0] ? Utilities.formatDate(new Date(values[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "",
          nama: String(values[i][1] || ""),
          kehadiran: String(values[i][2] || ""),
          jumlah: String(values[i][3] || ""),
          ucapan: String(values[i][4] || "")
        });
      }
      return jsonOut_({ status: "ok", data: data });
    }
    return jsonOut_({ status: "error", message: "action tidak dikenal" });
  } catch (err) {
    return jsonOut_({ status: "error", message: String(err) });
  }
}

/**
 * POST untuk menyimpan RSVP dari tombol "Kirim Ucapan".
 * Kirim dari web sebagai text/plain agar lolos tanpa preflight CORS:
 *   fetch(URL, { method: "POST", body: JSON.stringify({...}) })
 * Isi: { nama, kehadiran, jumlah, ucapan }
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (err2) { payload = {}; }
    }
    if (e && e.parameter) {
      if (!payload.nama && e.parameter.nama) payload = e.parameter;
    }
    var nama = String(payload.nama || "").trim();
    var kehadiran = String(payload.kehadiran || "Hadir").trim();
    var jumlah = String(payload.jumlah || "1 Orang").trim();
    var ucapan = String(payload.ucapan || "").trim();

    if (!nama) {
      return jsonOut_({ status: "error", message: "Nama wajib diisi." });
    }

    getSheet_().appendRow([new Date(), nama, kehadiran, jumlah, ucapan]);
    return jsonOut_({ status: "ok", message: "Terima kasih, RSVP tersimpan." });
  } catch (err) {
    return jsonOut_({ status: "error", message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (err3) {}
  }
}
