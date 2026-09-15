/**
 * Rekap LKH MD - eSAF
 *
 * SETUP:
 * 1. Ganti LKHMD_SOURCE_ID dengan ID spreadsheet form responses LKH MD
 * 2. Di Apps Script editor: Services (+) > Drive API > Add
 * 3. Jalankan setupRekapLKHMD() sekali untuk membuat spreadsheet rekap & trigger
 *
 * Kolom source (Form Responses 1):
 *   A: Timestamp, B: Main Dealer, C: Nama Penanggung Jawab,
 *   D: Tema LKH MD, E: No. Registrasi LKH MD, F: Klasifikasi,
 *   G: Tanggal AH, H: Lampiran File LKH MD
 *   → Sesuaikan LKHMD_COL jika urutan kolom berbeda
 */

var LKHMD_SOURCE_ID = "1Z58B52MILTlzs4N2_UeG5Yf6BTmizzPmOcIXG6fh25I";
var LKHMD_START_DATE = new Date(2026, 6, 10); // 10 Juli 2026

var LKHMD_COL = {
  TIMESTAMP: 1,
  MAIN_DEALER: 2,
  NO_REGISTRASI: 5,
  LAMPIRAN: 8
};

function rekapLKHMD() {
  var sourceSS = SpreadsheetApp.openById(LKHMD_SOURCE_ID);
  var sourceSheet = sourceSS.getSheetByName("Form Responses 1");
  if (!sourceSheet) {
    Logger.log("Sheet 'Form Responses 1' tidak ditemukan");
    return;
  }

  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return;

  var lastCol = sourceSheet.getLastColumn();
  var data = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var rekapSS = getOrCreateRekapLKHMD();
  var rekapSheet = rekapSS.getSheets()[0];

  var existing = {};
  var rekapLastRow = rekapSheet.getLastRow();
  if (rekapLastRow > 1) {
    var existingData = rekapSheet.getRange(2, 3, rekapLastRow - 1, 1).getValues();
    for (var e = 0; e < existingData.length; e++) {
      existing[String(existingData[e][0]).trim()] = true;
    }
  }

  var newRows = [];

  for (var i = 0; i < data.length; i++) {
    var timestamp = new Date(data[i][LKHMD_COL.TIMESTAMP - 1]);
    if (isNaN(timestamp.getTime()) || timestamp < LKHMD_START_DATE) continue;

    var noRegistrasi = String(data[i][LKHMD_COL.NO_REGISTRASI - 1]).trim();
    if (!noRegistrasi || existing[noRegistrasi]) continue;

    var mainDealer = String(data[i][LKHMD_COL.MAIN_DEALER - 1]).trim();
    var lampiranLink = String(data[i][LKHMD_COL.LAMPIRAN - 1]).trim();

    var excelData = extractExcelDataLKHMD(lampiranLink);

    newRows.push([
      timestamp,
      mainDealer,
      noRegistrasi,
      excelData.noAhass,
      excelData.namaAhass,
      excelData.tipeMotor,
      excelData.noRangka,
      excelData.noMesin
    ]);

    existing[noRegistrasi] = true;
  }

  if (newRows.length > 0) {
    var startRow = rekapSheet.getLastRow() + 1;
    rekapSheet.getRange(startRow, 1, newRows.length, 8).setValues(newRows);
    Logger.log(newRows.length + " baris baru ditambahkan ke rekap.");
  } else {
    Logger.log("Tidak ada data baru untuk direkap.");
  }
}

function extractExcelDataLKHMD(link) {
  var result = {
    tipeMotor: "",
    noRangka: "",
    noMesin: "",
    noAhass: "",
    namaAhass: ""
  };

  if (!link) return result;

  var fileId = extractDriveFileId(link);
  if (!fileId) {
    Logger.log("Gagal extract file ID dari: " + link);
    return result;
  }

  var tempFileId = null;

  try {
    var file = DriveApp.getFileById(fileId);
    var mimeType = file.getMimeType();
    var ssId;

    if (mimeType === "application/vnd.google-apps.spreadsheet") {
      ssId = fileId;
    } else {
      var blob = file.getBlob();
      var tempFile = Drive.Files.insert(
        { title: "temp_lkhmd_" + fileId, mimeType: "application/vnd.google-apps.spreadsheet" },
        blob,
        { convert: true }
      );
      tempFileId = tempFile.id;
      ssId = tempFileId;
    }

    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName("LKH MD") || ss.getSheets()[0];

    result.tipeMotor = String(sheet.getRange("G7").getValue()).trim();
    result.noRangka = String(sheet.getRange("G8").getValue()).trim();
    result.noMesin = String(sheet.getRange("G9").getValue()).trim();
    result.noAhass = String(sheet.getRange("Z7").getValue()).trim();
    result.namaAhass = String(sheet.getRange("Z8").getValue()).trim();
  } catch (e) {
    Logger.log("Error baca file " + link + ": " + e.message);
  } finally {
    if (tempFileId) {
      try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (err) {}
    }
  }

  return result;
}

function extractDriveFileId(url) {
  if (!url) return null;
  var patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/open\?id=([a-zA-Z0-9_-]+)/
  ];
  for (var p = 0; p < patterns.length; p++) {
    var match = url.match(patterns[p]);
    if (match) return match[1];
  }
  return null;
}

function getOrCreateRekapLKHMD() {
  var props = PropertiesService.getScriptProperties();
  var rekapId = props.getProperty("LKHMD_REKAP_ID");

  if (rekapId) {
    try {
      return SpreadsheetApp.openById(rekapId);
    } catch (e) {
      Logger.log("Rekap spreadsheet lama tidak ditemukan, buat baru...");
    }
  }

  var newSS = SpreadsheetApp.create("Rekap LKH MD - eSAF");
  var sheet = newSS.getSheets()[0];
  sheet.setName("Rekap LKH MD");

  var headers = [
    "Timestamp", "Main Dealer", "No. Registrasi LKH MD",
    "No. AHASS", "Nama AHASS", "Tipe Motor",
    "No. Rangka", "No. Mesin"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);

  props.setProperty("LKHMD_REKAP_ID", newSS.getId());
  Logger.log("Spreadsheet rekap dibuat: " + newSS.getUrl());

  return newSS;
}

function setupRekapLKHMD() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === "onFormSubmitLKHMD") {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }

  var sourceSS = SpreadsheetApp.openById(LKHMD_SOURCE_ID);
  ScriptApp.newTrigger("onFormSubmitLKHMD")
    .forSpreadsheet(sourceSS)
    .onFormSubmit()
    .create();

  rekapLKHMD();

  var props = PropertiesService.getScriptProperties();
  var rekapId = props.getProperty("LKHMD_REKAP_ID");
  Logger.log("Setup selesai!");
  Logger.log("Trigger onFormSubmit aktif.");
  Logger.log("Rekap: https://docs.google.com/spreadsheets/d/" + rekapId);
}

function onFormSubmitLKHMD(e) {
  rekapLKHMD();
}

/**
 * Jalankan fungsi ini untuk debug 1 file Excel lampiran.
 * Cek Execution Log untuk hasilnya.
 * Salin salah satu link dari kolom Lampiran ke variabel testLink di bawah.
 */
function debugExcelFileLKHMD() {
  var testLink = "PASTE_LINK_LAMPIRAN_DISINI";

  var fileId = extractDriveFileId(testLink);
  if (!fileId) {
    Logger.log("Gagal extract file ID dari link");
    return;
  }

  var file = DriveApp.getFileById(fileId);
  Logger.log("=== FILE INFO ===");
  Logger.log("Nama: " + file.getName());
  Logger.log("MIME type: " + file.getMimeType());

  var tempFileId = null;
  var ssId;

  if (file.getMimeType() === "application/vnd.google-apps.spreadsheet") {
    ssId = fileId;
  } else {
    var blob = file.getBlob();
    var tempFile = Drive.Files.insert(
      { title: "debug_lkhmd", mimeType: "application/vnd.google-apps.spreadsheet" },
      blob,
      { convert: true }
    );
    tempFileId = tempFile.id;
    ssId = tempFileId;
  }

  var ss = SpreadsheetApp.openById(ssId);
  var sheets = ss.getSheets();
  Logger.log("\n=== JUMLAH SHEET: " + sheets.length + " ===");

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();
    var lastR = sheet.getLastRow();
    var lastC = sheet.getLastColumn();
    Logger.log("\n--- Sheet " + (s + 1) + ": '" + name + "' (baris: " + lastR + ", kolom: " + lastC + ") ---");

    Logger.log("G7 = [" + sheet.getRange("G7").getValue() + "]");
    Logger.log("G8 = [" + sheet.getRange("G8").getValue() + "]");
    Logger.log("G9 = [" + sheet.getRange("G9").getValue() + "]");
    Logger.log("Z7 = [" + sheet.getRange("Z7").getValue() + "]");
    Logger.log("Z8 = [" + sheet.getRange("Z8").getValue() + "]");

    if (lastR > 0 && lastC > 0) {
      var previewRows = Math.min(lastR, 15);
      var previewCols = Math.min(lastC, 30);
      var preview = sheet.getRange(1, 1, previewRows, previewCols).getValues();
      for (var r = 0; r < preview.length; r++) {
        var rowStr = "Baris " + (r + 1) + ": ";
        for (var c = 0; c < preview[r].length; c++) {
          var val = String(preview[r][c]).trim();
          if (val) {
            var colLetter = "";
            var colNum = c + 1;
            while (colNum > 0) {
              colLetter = String.fromCharCode(((colNum - 1) % 26) + 65) + colLetter;
              colNum = Math.floor((colNum - 1) / 26);
            }
            rowStr += colLetter + "=[" + val.substring(0, 30) + "] ";
          }
        }
        Logger.log(rowStr);
      }
    }
  }

  if (tempFileId) {
    DriveApp.getFileById(tempFileId).setTrashed(true);
  }

  Logger.log("\n=== DEBUG SELESAI ===");
}
