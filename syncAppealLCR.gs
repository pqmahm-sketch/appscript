// Sync hasil banding LCR ke SUMBER HULU (FMA & GKA).
// Data mentah di dashboard dihasilkan lewat QUERY dari kedua sumber ini,
// jadi update dilakukan di hulu supaya rantai query otomatis membawa
// hasil banding turun ke:
//   Master Verifikasi Video LCR -> [LIVE] Dashboard AMORE For AHM -> For MD.
// Kolom "Status AHASS" di dashboard adalah kolom derived (dihitung dari
// kolom D-H), sehingga TIDAK perlu di-sync manual.

var APPEAL_SPREADSHEET_ID = "1abTssFzPCCQrLFcqRKOdB36eTDij8UF4EhjR26wEW8M"; // Tinjau Ulang Video LCR (Responses)
var APPEAL_SHEET_NAME     = "Form Responses 1";

// Sumber hulu yang isinya akan di-update
var SOURCE_SPREADSHEETS = [
  { id: "1t7ywiElPcdfLsPJ07TI4H-oLJNauT_CyVg2CwKFMN6M", label: "FMA", sheetName: "Form Responses 1" },
  { id: "1Epr6jzg8XyxyqYU_pYkQ7-Xvz_elRaCOl8StHVhVv3g", label: "GKA", sheetName: "Form Responses 1" }
];

// Mapping nama kolom di form banding -> nama kolom di sheet sumber (FMA/GKA)
var COLUMN_MAP = {
  "Proses Aplikasi Flintkote": "Proses Aplikasi Flintkote",
  "Proses Setting Regulator":  "Proses Setting Regulator",
  "Proses Timer":              "Proses Timer",
  "Marking garis pada selang": "Marking garis pada selang",
  "Proses Marking Frame Body": "Proses Marking Frame Body",
  "Catatan":                   "Catatan Tambahan"
};

var KEY_APPEAL       = "No. Rangka";
var KEY_SOURCE       = "No. Rangka";
var TIMESTAMP_APPEAL = "Timestamp";

// Nama kolom audit di sheet sumber. Jika belum ada, akan dibuat otomatis
// di kolom paling kanan. Diisi dengan timestamp entry banding yang diterapkan.
var AUDIT_COLUMN_NAME = "Terakhir Diupdate Banding";

function syncAppealLCR() {
  var appealSS = SpreadsheetApp.openById(APPEAL_SPREADSHEET_ID);
  var appealSheet = appealSS.getSheetByName(APPEAL_SHEET_NAME);
  if (!appealSheet) throw new Error("Sheet '" + APPEAL_SHEET_NAME + "' tidak ditemukan di spreadsheet banding.");

  var appealValues = appealSheet.getDataRange().getValues();
  if (appealValues.length < 2) {
    Logger.log("Sheet banding kosong.");
    return;
  }

  var appealHeaderIdx = findHeaderRow(appealValues, KEY_APPEAL);
  if (appealHeaderIdx < 0) throw new Error("Header '" + KEY_APPEAL + "' tidak ditemukan di sheet banding.");
  var appealHeader = appealValues[appealHeaderIdx];
  var appealCol = buildColumnIndex(appealHeader);

  var appealNeeded = [KEY_APPEAL, TIMESTAMP_APPEAL].concat(Object.keys(COLUMN_MAP));
  for (var i = 0; i < appealNeeded.length; i++) {
    if (appealCol[appealNeeded[i]] === undefined) {
      throw new Error("Kolom '" + appealNeeded[i] + "' tidak ada di sheet banding.");
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

  var totalAhassUpdated = 0;
  var totalCellsWritten = 0;
  var unmatched = {};
  for (var k in latestByRangka) unmatched[k] = true;

  for (var s = 0; s < SOURCE_SPREADSHEETS.length; s++) {
    var src = SOURCE_SPREADSHEETS[s];
    var srcResult = syncOneSource(src, appealCol, latestByRangka);
    Logger.log("[" + src.label + "] AHASS terupdate: " + srcResult.updatedCount
      + " | Sel diubah: " + srcResult.cellsWritten
      + " | Match: " + srcResult.matched);
    totalAhassUpdated += srcResult.updatedCount;
    totalCellsWritten += srcResult.cellsWritten;
    for (var m = 0; m < srcResult.matchedKeys.length; m++) {
      delete unmatched[srcResult.matchedKeys[m]];
    }
  }

  var unmatchedKeys = Object.keys(unmatched);
  Logger.log("=== Sync banding LCR selesai ===");
  Logger.log("Total AHASS terupdate: " + totalAhassUpdated
    + " | Total sel diubah: " + totalCellsWritten
    + " | No. Rangka tidak ditemukan di FMA maupun GKA: " + unmatchedKeys.length);
  if (unmatchedKeys.length) {
    Logger.log("Contoh No. Rangka tak ditemukan: " + unmatchedKeys.slice(0, 10).join(", "));
  }
}

function syncOneSource(src, appealCol, latestByRangka) {
  var ss = SpreadsheetApp.openById(src.id);
  var sheet = ss.getSheetByName(src.sheetName);
  if (!sheet) throw new Error("Sheet '" + src.sheetName + "' tidak ditemukan di sumber '" + src.label + "'.");

  var values = sheet.getDataRange().getValues();
  var headerIdx = findHeaderRow(values, KEY_SOURCE);
  if (headerIdx < 0) throw new Error("Header '" + KEY_SOURCE + "' tidak ditemukan di sumber '" + src.label + "'.");
  var col = buildColumnIndex(values[headerIdx]);

  var needed = [KEY_SOURCE].concat(Object.keys(COLUMN_MAP).map(function (k) { return COLUMN_MAP[k]; }));
  for (var i = 0; i < needed.length; i++) {
    if (col[needed[i]] === undefined) {
      throw new Error("Kolom '" + needed[i] + "' tidak ada di sumber '" + src.label + "'.");
    }
  }

  // Pastikan kolom audit ada; bikin di paling kanan kalau belum ada.
  var auditColIdx0 = col[AUDIT_COLUMN_NAME];
  if (auditColIdx0 === undefined) {
    auditColIdx0 = values[headerIdx].length; // 0-based -> append di kanan
    sheet.getRange(headerIdx + 1, auditColIdx0 + 1).setValue(AUDIT_COLUMN_NAME);
    Logger.log("[" + src.label + "] Kolom audit '" + AUDIT_COLUMN_NAME + "' dibuat di kolom "
      + columnLetter(auditColIdx0 + 1) + ".");
    // Re-read values supaya lebar tiap row konsisten
    values = sheet.getDataRange().getValues();
    col = buildColumnIndex(values[headerIdx]);
  }
  var auditCol1 = auditColIdx0 + 1;

  // Kumpulkan SEMUA row (1-based) untuk setiap No. Rangka di sumber
  // (satu rangka bisa muncul >1x kalau ada penilaian ulang oleh PIC).
  var rowsByRangka = {};
  for (var r = headerIdx + 1; r < values.length; r++) {
    var rk = normalizeKey(values[r][col[KEY_SOURCE]]);
    if (!rk) continue;
    if (!rowsByRangka[rk]) rowsByRangka[rk] = [];
    rowsByRangka[rk].push(r + 1); // 1-based
  }

  var mappedKeys = Object.keys(COLUMN_MAP);
  var updatesByCol = {};
  var updatedCount = 0;
  var matchedKeys = [];

  for (var key in latestByRangka) {
    var targetRows = rowsByRangka[key];
    if (!targetRows || !targetRows.length) continue;
    matchedKeys.push(key);
    var appealEntry = latestByRangka[key];
    var appealRow = appealEntry.row;
    var appealTsRaw = appealRow[appealCol[TIMESTAMP_APPEAL]];
    var appealTsValue = (appealTsRaw instanceof Date)
      ? appealTsRaw
      : (appealTsRaw ? new Date(appealTsRaw) : "");

    // Kalau ada beberapa row untuk 1 rangka, timpa SEMUA supaya
    // query hilir tidak mengambil versi lama.
    var didWriteAny = false;
    for (var t = 0; t < targetRows.length; t++) {
      var targetRow = targetRows[t];
      for (var m = 0; m < mappedKeys.length; m++) {
        var srcName = mappedKeys[m];
        var dstName = COLUMN_MAP[srcName];
        var srcVal  = appealRow[appealCol[srcName]];
        var dstCol1 = col[dstName] + 1;
        var currentVal = values[targetRow - 1][col[dstName]];
        if (!valuesEqual(currentVal, srcVal)) {
          if (!updatesByCol[dstCol1]) updatesByCol[dstCol1] = [];
          updatesByCol[dstCol1].push({ row: targetRow, value: srcVal });
          didWriteAny = true;
        }
      }
      if (didWriteAny) {
        // Tulis/refresh kolom audit hanya kalau row memang berubah, atau
        // kalau timestamp banding lebih baru dari yang sudah tercatat.
        var currentAudit = values[targetRow - 1][auditColIdx0];
        var currentAuditMs = toMillis(currentAudit);
        var newAuditMs = toMillis(appealTsValue);
        if (!valuesEqual(currentAudit, appealTsValue) && newAuditMs >= currentAuditMs) {
          if (!updatesByCol[auditCol1]) updatesByCol[auditCol1] = [];
          updatesByCol[auditCol1].push({ row: targetRow, value: appealTsValue });
        }
      }
    }
    if (didWriteAny) updatedCount++;
  }

  var cellsWritten = 0;
  for (var c in updatesByCol) {
    var items = updatesByCol[c];
    for (var it = 0; it < items.length; it++) {
      sheet.getRange(items[it].row, Number(c)).setValue(items[it].value);
      cellsWritten++;
    }
  }

  return {
    updatedCount: updatedCount,
    cellsWritten: cellsWritten,
    matched: matchedKeys.length,
    matchedKeys: matchedKeys
  };
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

function columnLetter(col1) {
  var s = "";
  var n = col1;
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ---- Installer trigger (jalankan sekali secara manual di Apps Script editor) ----

// Time-based: jalankan setiap 5 menit. Aman untuk semua deployment
// (bound maupun standalone).
function createSyncAppealLCRTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === "syncAppealLCR"
        && triggers[t].getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
  ScriptApp.newTrigger("syncAppealLCR")
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log("Trigger sync banding LCR (setiap 5 menit) berhasil dibuat.");
}

// onFormSubmit: sync langsung berjalan tepat setelah form banding disubmit.
// Script ini HARUS di-bound ke spreadsheet Tinjau Ulang Video LCR (Responses).
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
