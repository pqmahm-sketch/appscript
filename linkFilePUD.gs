// ============================================================
// Database FU PUD — pipeline sinkronisasi & Status Akhir kolom L
// ------------------------------------------------------------
// Alur (sesuai permintaan):
//   1. Look up di sheet PUD (spreadsheet Database FU PUD)
//      → isi kolom L untuk kategori yang cukup dilihat dari
//        kolom E–J saja (Belum PUD, Akan PUD, Deklarasi PUD*)
//   2. Transfer data antar sheet di Database FU PUD
//      → PUD ← breakdown (B10, W01, dst); breakdown baru
//        dipindah ke PUD setelah 2 menit dari waktu pengisian
//   3. Look up antar spreadsheet FU PUD ↔ Database Pengerjaan PUD
//      → isi kolom L untuk Deklarasi PUD, Deklarasi PUD belum
//        dikirim, Selesai PUD, Selesai PUD [pindah unit …]
//   4. Hapus baris ber-status "Selesai PUD" di kolom L
//   5. Transfer balik PUD → breakdown sheets
// ============================================================

var FU_PUD_SPREADSHEET_ID        = 'ISI_ID_DATABASE_FU_PUD';
var PENGERJAAN_PUD_SPREADSHEET_ID = 'ISI_ID_DATABASE_PENGERJAAN_PUD';

var MASTER_SHEET_NAME       = 'PUD';   // di Database FU PUD
var PENGERJAAN_SHEET_NAME   = 'PUD';   // di Database Pengerjaan PUD

var HEADER_ROW = 1;

// Index kolom (1-based) di sheet PUD Database FU PUD
var COL_B_MD     = 2;   // kode MD → penentu nama breakdown sheet
var COL_C_RANGKA = 3;   // no rangka → key lookup
var COL_E        = 5;
var COL_F        = 6;
var COL_G        = 7;
var COL_H        = 8;
var COL_I        = 9;
var COL_J        = 10;
var COL_L        = 12;
var LAST_DATA_COL = 12; // sampai kolom L

// Index kolom (1-based) di sheet PUD Database Pengerjaan PUD
var PGJ_COL_C_RANGKA = 3;
var PGJ_COL_I        = 9;
var PGJ_COL_J        = 10;
var PGJ_COL_K        = 11;
var PGJ_COL_L        = 12;

// Untuk delay 2 menit setelah edit di sheet breakdown
var PENDING_PROP_KEY = 'FU_PUD_PENDING_EDITS';
var DELAY_MS         = 2 * 60 * 1000;

// ============================================================
// ORCHESTRATOR
// ============================================================
function runFullPipeline() {
  flushPendingBreakdownEdits();   // Step 2 (bagian delay 2 menit)
  syncPUDFromBreakdowns();        // Step 2 (transfer breakdown → PUD)
  updateStatusAkhirLocal();       // Step 1
  updateStatusAkhirWithPengerjaan(); // Step 3
  deleteSelesaiPUDRows();         // Step 4
  syncBreakdownsFromPUD();        // Step 5
  Logger.log('Pipeline Database FU PUD selesai.');
}

// ============================================================
// STEP 1 — Status Akhir yang cukup dari data sheet PUD sendiri
//   • Belum PUD
//   • Akan PUD
//   • Deklarasi PUD (tanda awal: J = "Deklarasi PUD")
// ============================================================
function updateStatusAkhirLocal() {
  var sheet = getFuPudSheet_(MASTER_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  var values = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, LAST_DATA_COL).getValues();
  var newL = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var currentL = String(row[COL_L - 1] || '').trim();
    var status = classifyLocal_(row);

    // Jangan timpa status yang sudah diisi oleh step 3
    if (!status && currentL) {
      newL.push([currentL]);
    } else {
      newL.push([status || currentL || '']);
    }
  }

  sheet.getRange(HEADER_ROW + 1, COL_L, newL.length, 1).setValues(newL);
}

function classifyLocal_(row) {
  var e = String(row[COL_E - 1] || '').trim();
  var f = String(row[COL_F - 1] || '').trim();
  var g = String(row[COL_G - 1] || '').trim();
  var h = String(row[COL_H - 1] || '').trim();
  var iVal = row[COL_I - 1];
  var j = String(row[COL_J - 1] || '').trim();

  var isSudah      = /sudah/i.test(e);
  var isTerhubung  = /terhubung/i.test(f);
  var isBersedia   = /bersedia\s*pud/i.test(g);
  var isDeklarasiJ = /deklarasi\s*pud/i.test(j);
  var iIsDate      = isDate_(iVal);

  // Deklarasi PUD (tanda awal di kolom J) — status final ditentukan step 3
  if (isSudah && isTerhubung && g && isDeklarasiJ) {
    return 'Deklarasi PUD, belum dikirim';
  }
  // Akan PUD
  if (isSudah && isTerhubung && isBersedia && h && iIsDate) {
    return 'Akan PUD';
  }
  // Belum PUD (E–H terisi, I belum tanggal)
  if (e && f && g && h && !iIsDate) {
    return 'Belum PUD';
  }
  return '';
}

// ============================================================
// STEP 3 — Lookup ke Database Pengerjaan PUD
//   • Deklarasi PUD               (K mengandung "deklarasi")
//   • Deklarasi PUD, belum dikirim (K tidak mengandung "deklarasi")
//   • Selesai PUD                 (I & J terisi)
//   • Selesai PUD, [pindah unit …] (L mengandung "unit pindah target")
// ============================================================
function updateStatusAkhirWithPengerjaan() {
  var sheet = getFuPudSheet_(MASTER_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  var values = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, LAST_DATA_COL).getValues();
  var pengerjaanMap = buildPengerjaanMap_();
  var newL = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var currentL = String(row[COL_L - 1] || '').trim();
    var rangka   = String(row[COL_C_RANGKA - 1] || '').trim();
    var pgj      = rangka ? pengerjaanMap[rangka] : null;

    var e = String(row[COL_E - 1] || '').trim();
    var f = String(row[COL_F - 1] || '').trim();
    var g = String(row[COL_G - 1] || '').trim();
    var j = String(row[COL_J - 1] || '').trim();

    var isSudah      = /sudah/i.test(e);
    var isTerhubung  = /terhubung/i.test(f);
    var isDeklarasiJ = /deklarasi\s*pud/i.test(j);

    var next = currentL;

    // Prioritas: Selesai PUD [pindah unit] > Selesai PUD > Deklarasi PUD*
    if (pgj) {
      var pgjI = pgj[PGJ_COL_I - 1];
      var pgjJ = pgj[PGJ_COL_J - 1];
      var pgjK = String(pgj[PGJ_COL_K - 1] || '');
      var pgjL = String(pgj[PGJ_COL_L - 1] || '');

      var ijFilled = notEmpty_(pgjI) && notEmpty_(pgjJ);

      if (ijFilled && /unit\s*pindah\s*target/i.test(pgjL)) {
        // Ambil keterangan tambahan dari kolom L Pengerjaan PUD, taruh
        // dalam bracket sesuai instruksi ("nama kategori telah diberikan bracket []")
        var extra = pgjL.replace(/unit\s*pindah\s*target/i, '').replace(/^[\s,;:.\-]+|[\s,;:.\-]+$/g, '');
        next = extra
          ? 'Selesai PUD, [pindah unit ' + extra + ']'
          : 'Selesai PUD, [pindah unit target]';
      } else if (ijFilled) {
        next = 'Selesai PUD';
      } else if (isSudah && isTerhubung && g && isDeklarasiJ) {
        next = /deklarasi/i.test(pgjK) ? 'Deklarasi PUD' : 'Deklarasi PUD, belum dikirim';
      }
    } else if (isSudah && isTerhubung && g && isDeklarasiJ) {
      // Tidak ada match di Pengerjaan PUD → tetap "belum dikirim"
      next = 'Deklarasi PUD, belum dikirim';
    }

    newL.push([next]);
  }

  sheet.getRange(HEADER_ROW + 1, COL_L, newL.length, 1).setValues(newL);
}

function buildPengerjaanMap_() {
  var ss = SpreadsheetApp.openById(PENGERJAAN_PUD_SPREADSHEET_ID);
  var sh = ss.getSheetByName(PENGERJAAN_SHEET_NAME);
  if (!sh) throw new Error('Sheet ' + PENGERJAAN_SHEET_NAME + ' tidak ditemukan di Database Pengerjaan PUD');
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), PGJ_COL_L);
  var map = {};
  if (lastRow <= HEADER_ROW) return map;
  var data = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    var rangka = String(data[i][PGJ_COL_C_RANGKA - 1] || '').trim();
    if (rangka) map[rangka] = data[i];
  }
  return map;
}

// ============================================================
// STEP 4 — Hapus baris ber-status "Selesai PUD" (persis) di kolom L
//   Baris "Selesai PUD, [pindah unit …]" TIDAK dihapus.
// ============================================================
function deleteSelesaiPUDRows() {
  var sheet = getFuPudSheet_(MASTER_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  var lValues = sheet.getRange(HEADER_ROW + 1, COL_L, lastRow - HEADER_ROW, 1).getValues();
  var toDelete = [];
  for (var i = 0; i < lValues.length; i++) {
    if (String(lValues[i][0] || '').trim().toLowerCase() === 'selesai pud') {
      toDelete.push(HEADER_ROW + 1 + i);
    }
  }
  toDelete.sort(function (a, b) { return b - a; });
  toDelete.forEach(function (r) { sheet.deleteRow(r); });
  Logger.log('Baris "Selesai PUD" dihapus: ' + toDelete.length);
}

// ============================================================
// STEP 2 — Transfer breakdown → PUD (setelah delay 2 menit)
// ============================================================

// Installable onEdit trigger — pasang lewat createBreakdownEditTrigger()
function onEditFuPud(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (!isBreakdownSheet_(sheet)) return;
  if (e.range.getRow() <= HEADER_ROW) return;

  var pending = readPending_();
  var key = sheet.getName() + '|' + e.range.getRow();
  pending[key] = Date.now();
  writePending_(pending);
}

// Jalankan tiap menit; sekaligus ambil baris yang usia editnya ≥ 2 menit.
function flushPendingBreakdownEdits() {
  var pending = readPending_();
  var now = Date.now();
  var rowsBySheet = {};
  var keptPending = {};

  Object.keys(pending).forEach(function (key) {
    if (now - pending[key] >= DELAY_MS) {
      var parts = key.split('|');
      var name = parts[0];
      var row = Number(parts[1]);
      (rowsBySheet[name] = rowsBySheet[name] || []).push(row);
    } else {
      keptPending[key] = pending[key];
    }
  });

  if (Object.keys(rowsBySheet).length === 0) {
    writePending_(keptPending);
    return;
  }

  var master = getFuPudSheet_(MASTER_SHEET_NAME);
  var masterIndex = buildMasterIndex_(master);

  Object.keys(rowsBySheet).forEach(function (sheetName) {
    var breakdown = getFuPudSheet_(sheetName);
    if (!breakdown) return;
    var rows = rowsBySheet[sheetName];
    rows.forEach(function (rowNum) {
      if (rowNum > breakdown.getLastRow()) return;
      var data = breakdown.getRange(rowNum, 1, 1, LAST_DATA_COL).getValues()[0];
      pushRowToMaster_(master, masterIndex, data);
    });
  });

  writePending_(keptPending);
}

// Full sync semua breakdown → PUD (dipakai runFullPipeline)
function syncPUDFromBreakdowns() {
  var ss = SpreadsheetApp.openById(FU_PUD_SPREADSHEET_ID);
  var master = ss.getSheetByName(MASTER_SHEET_NAME);
  var masterIndex = buildMasterIndex_(master);

  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    if (!isBreakdownSheet_(sh)) continue;
    var lastRow = sh.getLastRow();
    if (lastRow <= HEADER_ROW) continue;
    var data = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, LAST_DATA_COL).getValues();
    for (var i = 0; i < data.length; i++) {
      var rangka = String(data[i][COL_C_RANGKA - 1] || '').trim();
      if (!rangka) continue;
      pushRowToMaster_(master, masterIndex, data[i]);
    }
  }
}

// ============================================================
// STEP 5 — Transfer balik PUD → breakdown sheets
// ============================================================
function syncBreakdownsFromPUD() {
  var ss = SpreadsheetApp.openById(FU_PUD_SPREADSHEET_ID);
  var master = ss.getSheetByName(MASTER_SHEET_NAME);
  var lastRow = master.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  var masterData = master.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, LAST_DATA_COL).getValues();

  // Kelompokkan baris master berdasarkan kolom B (kode MD)
  var byMD = {};
  for (var i = 0; i < masterData.length; i++) {
    var md = String(masterData[i][COL_B_MD - 1] || '').trim();
    if (!md) continue;
    (byMD[md] = byMD[md] || []).push(masterData[i]);
  }

  Object.keys(byMD).forEach(function (md) {
    var sh = ss.getSheetByName(md);
    if (!sh) return; // breakdown sheet untuk MD ini belum ada — lewati
    var rows = byMD[md];
    var index = buildRowIndexByRangka_(sh);
    for (var r = 0; r < rows.length; r++) {
      var rangka = String(rows[r][COL_C_RANGKA - 1] || '').trim();
      if (!rangka) continue;
      var targetRow = index[rangka];
      if (targetRow) {
        sh.getRange(targetRow, 1, 1, LAST_DATA_COL).setValues([rows[r]]);
      } else {
        sh.appendRow(rows[r]);
        index[rangka] = sh.getLastRow();
      }
    }
  });
}

// ============================================================
// Helpers
// ============================================================
function getFuPudSheet_(name) {
  var ss = SpreadsheetApp.openById(FU_PUD_SPREADSHEET_ID);
  return ss.getSheetByName(name);
}

function isBreakdownSheet_(sheet) {
  var name = sheet.getName();
  if (name === MASTER_SHEET_NAME) return false;
  if (name.charAt(0) === '_') return false; // sheet sistem/hidden util
  // Anggap semua sheet lain di spreadsheet ini adalah breakdown per kode MD
  return sheet.getParent().getId() === FU_PUD_SPREADSHEET_ID;
}

function buildMasterIndex_(master) {
  var index = {};
  var lastRow = master.getLastRow();
  if (lastRow <= HEADER_ROW) return index;
  var keys = master.getRange(HEADER_ROW + 1, COL_C_RANGKA, lastRow - HEADER_ROW, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i][0] || '').trim();
    if (k) index[k] = HEADER_ROW + 1 + i;
  }
  return index;
}

function buildRowIndexByRangka_(sheet) {
  var index = {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return index;
  var keys = sheet.getRange(HEADER_ROW + 1, COL_C_RANGKA, lastRow - HEADER_ROW, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i][0] || '').trim();
    if (k) index[k] = HEADER_ROW + 1 + i;
  }
  return index;
}

function pushRowToMaster_(master, masterIndex, rowData) {
  var rangka = String(rowData[COL_C_RANGKA - 1] || '').trim();
  if (!rangka) return;
  var targetRow = masterIndex[rangka];
  if (targetRow) {
    master.getRange(targetRow, 1, 1, LAST_DATA_COL).setValues([rowData]);
  } else {
    master.appendRow(rowData);
    masterIndex[rangka] = master.getLastRow();
  }
}

function isDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return true;
  if (typeof v === 'string' && v.trim()) {
    var d = new Date(v);
    return !isNaN(d.getTime());
  }
  return false;
}

function notEmpty_(v) {
  if (v instanceof Date) return !isNaN(v.getTime());
  return String(v == null ? '' : v).trim() !== '';
}

function readPending_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PENDING_PROP_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (err) { return {}; }
}

function writePending_(obj) {
  PropertiesService.getScriptProperties().setProperty(PENDING_PROP_KEY, JSON.stringify(obj));
}

// ============================================================
// Trigger installers — jalankan sekali dari editor Apps Script
// ============================================================
function createBreakdownEditTrigger() {
  var ss = SpreadsheetApp.openById(FU_PUD_SPREADSHEET_ID);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditFuPud') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditFuPud').forSpreadsheet(ss).onEdit().create();
  Logger.log('Installable onEdit trigger dipasang.');
}

function createFlushTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'flushPendingBreakdownEdits') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('flushPendingBreakdownEdits').timeBased().everyMinutes(1).create();
  Logger.log('Trigger flush pending (setiap 1 menit) dipasang.');
}

function createFullPipelineTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runFullPipeline') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runFullPipeline').timeBased().everyHours(1).create();
  Logger.log('Trigger pipeline (setiap 1 jam) dipasang.');
}

function installAllTriggers() {
  createBreakdownEditTrigger();
  createFlushTrigger();
  createFullPipelineTrigger();
}
