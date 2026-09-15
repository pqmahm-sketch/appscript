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
var LKHMD_START_DATE = new Date(2026, 6, 10);
var LKHMD_TIME_LIMIT = 5 * 60 * 1000; // 5 menit, sisakan 1 menit sebelum batas 6 menit

var LKHMD_COL = {
  TIMESTAMP: 1,
  MAIN_DEALER: 2,
  NO_REGISTRASI: 5,
  LAMPIRAN: 8
};

function rekapLKHMD() {
  var startTime = new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  var lastProcessed = Number(props.getProperty("LKHMD_LAST_INDEX") || 0);

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
  var timedOut = false;
  var processedUntil = data.length;

  for (var i = lastProcessed; i < data.length; i++) {
    if (new Date().getTime() - startTime > LKHMD_TIME_LIMIT) {
      timedOut = true;
      processedUntil = i;
      Logger.log("Batas waktu hampir tercapai, berhenti di baris " + (i + 2) + ". Akan lanjut otomatis.");
      break;
    }

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
  }

  if (timedOut) {
    props.setProperty("LKHMD_LAST_INDEX", String(processedUntil));
    scheduleResumeLKHMD();
  } else {
    props.deleteProperty("LKHMD_LAST_INDEX");
    clearResumeTriggerLKHMD();
    Logger.log("Rekap selesai. Total data: " + data.length + " baris.");
  }
}

function scheduleResumeLKHMD() {
  clearResumeTriggerLKHMD();
  ScriptApp.newTrigger("rekapLKHMD")
    .timeBased()
    .after(30 * 1000)
    .create();
  Logger.log("Trigger resume dijadwalkan 30 detik lagi.");
}

function clearResumeTriggerLKHMD() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    var trigger = triggers[t];
    if (trigger.getHandlerFunction() === "rekapLKHMD" && trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
    }
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

    var gRange = sheet.getRange("G7:G9").getValues();
    result.tipeMotor = String(gRange[0][0]).trim();
    result.noRangka = String(gRange[1][0]).trim();
    result.noMesin = String(gRange[2][0]).trim();

    var zRange = sheet.getRange("Z7:Z8").getValues();
    result.noAhass = String(zRange[0][0]).trim();
    result.namaAhass = String(zRange[1][0]).trim();
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

  PropertiesService.getScriptProperties().deleteProperty("LKHMD_LAST_INDEX");
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
