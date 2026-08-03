var APPEAL_SPREADSHEET_ID = "1abTssFzPCCQrLFcqRKOdB36eTDij8UF4EhjR26wEW8M";
var APPEAL_SHEET_NAME = "Form Responses 1";

var DASHBOARD_SPREADSHEET_ID = "1uU3mG1tu-d9aW1Z-eL-ocOQHhMzg2xPWAN4DvSYmo1o";
var DASHBOARD_SHEET_NAME = "data mentah";

// Mapping nama kolom Form Responses 1 -> nama kolom data mentah
var COLUMN_MAP = {
  "Proses Aplikasi Flintkote": "Proses Aplikasi Flintkote",
  "Proses Setting Regulator":  "Proses Setting Regulator",
  "Proses Timer":              "Proses Timer",
  "Marking garis pada selang": "Marking garis pada selang",
  "Proses Marking Frame Body": "Proses Marking Frame Body",
  "Catatan":                   "Catatan Tambahan",
  "Status Tinjau Ulang":       "Status AHASS"
};

var KEY_APPEAL    = "No. Rangka";
var KEY_DASHBOARD = "No. Rangka";
var TIMESTAMP_APPEAL = "Timestamp";

function syncAppealLCR() {
  var appealSS = SpreadsheetApp.openById(APPEAL_SPREADSHEET_ID);
  var appealSheet = appealSS.getSheetByName(APPEAL_SHEET_NAME);
  if (!appealSheet) throw new Error("Sheet '" + APPEAL_SHEET_NAME + "' tidak ditemukan.");

  var dashSS = SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
  var dashSheet = dashSS.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashSheet) throw new Error("Sheet '" + DASHBOARD_SHEET_NAME + "' tidak ditemukan.");

  var appealValues = appealSheet.getDataRange().getValues();
  if (appealValues.length < 2) {
    Logger.log("Form Responses 1 kosong.");
    return;
  }
  var appealHeaderIdx = findHeaderRow(appealValues, KEY_APPEAL);
  if (appealHeaderIdx < 0) throw new Error("Header '" + KEY_APPEAL + "' tidak ditemukan di Form Responses 1.");
  var appealHeader = appealValues[appealHeaderIdx];
  var appealCol = buildColumnIndex(appealHeader);

  var dashValues = dashSheet.getDataRange().getValues();
  var dashHeaderIdx = findHeaderRow(dashValues, KEY_DASHBOARD);
  if (dashHeaderIdx < 0) throw new Error("Header '" + KEY_DASHBOARD + "' tidak ditemukan di data mentah.");
  var dashHeader = dashValues[dashHeaderIdx];
  var dashCol = buildColumnIndex(dashHeader);

  // Validasi kolom yang dibutuhkan
  var appealNeeded = [KEY_APPEAL, TIMESTAMP_APPEAL].concat(Object.keys(COLUMN_MAP));
  for (var i = 0; i < appealNeeded.length; i++) {
    if (appealCol[appealNeeded[i]] === undefined) {
      throw new Error("Kolom '" + appealNeeded[i] + "' tidak ada di Form Responses 1.");
    }
  }
  var dashNeeded = [KEY_DASHBOARD].concat(
    Object.keys(COLUMN_MAP).map(function (k) { return COLUMN_MAP[k]; })
  );
  for (var j = 0; j < dashNeeded.length; j++) {
    if (dashCol[dashNeeded[j]] === undefined) {
      throw new Error("Kolom '" + dashNeeded[j] + "' tidak ada di data mentah.");
    }
  }

  // Ambil entry banding TERBARU per No. Rangka (berdasarkan Timestamp)
  var latestByRangka = {};
  for (var r = appealHeaderIdx + 1; r < appealValues.length; r++) {
    var row = appealValues[r];
    var rangka = normalizeKey(row[appealCol[KEY_APPEAL]]);
    if (!rangka) continue;
    var ts = toMillis(row[appealCol[TIMESTAMP_APPEAL]]);
    var existing = latestByRangka[rangka];
    if (!existing || ts >= existing.ts) {
      latestByRangka[rangka] = { ts: ts, row: row };
    }
  }

  // Peta No. Rangka -> baris (1-based) di data mentah
  var dashRowIdx = {};
  var dashKeyCol = dashCol[KEY_DASHBOARD];
  for (var d = dashHeaderIdx + 1; d < dashValues.length; d++) {
    var key = normalizeKey(dashValues[d][dashKeyCol]);
    if (!key) continue;
    dashRowIdx[key] = d + 1; // 1-based row untuk getRange
  }

  // Susun update per kolom (batch per column) supaya efisien
  var mappedKeys = Object.keys(COLUMN_MAP);
  var updatesByCol = {};   // colIndex1based -> array of {row, value}
  var updatedCount = 0;
  var notFound = [];

  var rangkaKeys = Object.keys(latestByRangka);
  for (var k = 0; k < rangkaKeys.length; k++) {
    var rk = rangkaKeys[k];
    var targetRow = dashRowIdx[rk];
    if (!targetRow) {
      notFound.push(rk);
      continue;
    }
    var appealRow = latestByRangka[rk].row;
    var changed = false;
    for (var m = 0; m < mappedKeys.length; m++) {
      var srcName = mappedKeys[m];
      var dstName = COLUMN_MAP[srcName];
      var srcVal = appealRow[appealCol[srcName]];
      var dstCol1 = dashCol[dstName] + 1;
      var currentVal = dashValues[targetRow - 1][dashCol[dstName]];
      if (!valuesEqual(currentVal, srcVal)) {
        if (!updatesByCol[dstCol1]) updatesByCol[dstCol1] = [];
        updatesByCol[dstCol1].push({ row: targetRow, value: srcVal });
        changed = true;
      }
    }
    if (changed) updatedCount++;
  }

  // Terapkan update
  var totalCells = 0;
  var colKeys = Object.keys(updatesByCol);
  for (var c = 0; c < colKeys.length; c++) {
    var col1 = Number(colKeys[c]);
    var items = updatesByCol[col1];
    for (var it = 0; it < items.length; it++) {
      dashSheet.getRange(items[it].row, col1).setValue(items[it].value);
      totalCells++;
    }
  }

  Logger.log("Sync banding LCR selesai. AHASS terupdate: " + updatedCount
    + " | Sel diubah: " + totalCells
    + " | No. Rangka tidak ditemukan di data mentah: " + notFound.length);
  if (notFound.length) {
    Logger.log("Contoh No. Rangka tak ditemukan: " + notFound.slice(0, 10).join(", "));
  }
}

function findHeaderRow(values, keyName) {
  var target = String(keyName).trim().toLowerCase();
  var maxScan = Math.min(values.length, 20);
  for (var r = 0; r < maxScan; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim().toLowerCase() === target) return r;
    }
  }
  return -1;
}

function buildColumnIndex(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var name = String(headerRow[i]).trim();
    if (name) map[name] = i;
  }
  return map;
}

function normalizeKey(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().toUpperCase();
}

function toMillis(v) {
  if (v instanceof Date) return v.getTime();
  if (v === null || v === undefined || v === "") return 0;
  var d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function valuesEqual(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  var sa = (a === null || a === undefined) ? "" : String(a).trim();
  var sb = (b === null || b === undefined) ? "" : String(b).trim();
  return sa === sb;
}

// ---- Installer trigger (jalankan sekali secara manual di Apps Script editor) ----
function createSyncAppealLCRTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === "syncAppealLCR") {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
  ScriptApp.newTrigger("syncAppealLCR")
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log("Trigger sync banding LCR (setiap 5 menit) berhasil dibuat.");
}

// Opsional: pasang onFormSubmit agar sync langsung berjalan setelah banding disubmit.
// Jalankan sekali dari Apps Script editor. Script ini HARUS di-bound ke spreadsheet
// Tinjau Ulang Video LCR (Responses), atau ubah openById -> forSpreadsheet(...).
function createSyncAppealLCROnFormSubmit() {
  var appealSS = SpreadsheetApp.openById(APPEAL_SPREADSHEET_ID);
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === "syncAppealLCR"
        && triggers[t].getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
  ScriptApp.newTrigger("syncAppealLCR")
    .forSpreadsheet(appealSS)
    .onFormSubmit()
    .create();
  Logger.log("Trigger onFormSubmit untuk syncAppealLCR berhasil dibuat.");
}
