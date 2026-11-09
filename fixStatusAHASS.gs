// Fix untuk bug "Status AHASS tergeser 1 baris" di kolom J.
//
// Masalah: formula Status AHASS mereferensikan baris N+1 padahal
//          seharusnya mereferensikan baris N (baris yang sama).
//          Akibatnya status OK/NG muncul di baris yang salah.
//
// Solusi:  Tulis ulang formula kolom J dengan referensi baris yang benar.
//
// Kriteria OK (semua harus terpenuhi di kolom D-H baris yang SAMA):
//   D (Flintkote)      : dimulai "1 - "
//   E (Regulator)      : dimulai "23 - "  (> 5 BAR)
//   F (Timer)          : dimulai "1 - "
//   G (Marking Selang) : dimulai "1 - "
//   H (Marking FB)     : dimulai "1 - "
// Selain itu → NG.

var AHM_SPREADSHEET_ID    = "1JzPicTYV2iBnVpS-Lal1ZZ4IfaJARj9_fb6nIlEm42I";
var MASTER_SPREADSHEET_ID = "1xVkuqRoLhsP4BpfifurvV4zxeIDXIHCsZ3wdSmxDuig";

// =====================================================================
// LANGKAH 1 — Jalankan fungsi ini DULU untuk melihat pola offset.
// =====================================================================
function diagnoseStatusAHASS() {
  Logger.log("=== DIAGNOSIS STATUS AHASS ===\n");

  diagnoseSheet_(AHM_SPREADSHEET_ID, "Data Mentah", "For AHM", 1);
  diagnoseSheet_(MASTER_SPREADSHEET_ID, "Gabungan", "Master Gabungan", 0);
}

function diagnoseSheet_(ssId, sheetName, label, headerRows) {
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) { Logger.log("[" + label + "] Sheet '" + sheetName + "' tidak ditemukan."); return; }

  Logger.log("--- " + label + " ---");

  // Log formulas J1-J5
  var startCheck = 1;
  var endCheck   = Math.min(5, sheet.getLastRow());
  for (var r = startCheck; r <= endCheck; r++) {
    var cell = sheet.getRange(r, 10); // kolom J
    var f = cell.getFormula();
    var v = cell.getValue();
    Logger.log("J" + r + ": " + (f ? "FORMULA → " + f : "VALUE  → " + v));
  }

  // Cek apakah A1 adalah QUERY/IMPORTRANGE
  var a1f = sheet.getRange("A1").getFormula();
  if (a1f) Logger.log("A1 formula: " + a1f.substring(0, 300));

  // Verifikasi offset pada baris data
  var dataStart = headerRows + 1; // baris pertama data (1-indexed)
  var rowCount  = Math.min(8, sheet.getLastRow() - dataStart + 1);
  if (rowCount < 1) return;

  var data = sheet.getRange(dataStart, 4, rowCount, 7).getValues(); // D:J
  Logger.log("\nVerifikasi offset (" + label + " rows " + dataStart + "-" + (dataStart + rowCount - 1) + "):");
  for (var i = 0; i < data.length; i++) {
    var expected = allColumnsOK_(data[i]) ? "OK" : "NG";
    var actual   = String(data[i][6]).trim();
    Logger.log("  Row " + (dataStart + i) + ": expected=" + expected + " actual=" + actual
      + (expected !== actual ? "  ← MISMATCH" : "  ✓"));
  }
  Logger.log("");
}

function allColumnsOK_(row) {
  var d = String(row[0]); // Flintkote
  var e = String(row[1]); // Regulator
  var f = String(row[2]); // Timer
  var g = String(row[3]); // Marking Selang
  var h = String(row[4]); // Marking FB
  return d.indexOf("1 - ") === 0
      && e.indexOf("23 - ") === 0
      && f.indexOf("1 - ") === 0
      && g.indexOf("1 - ") === 0
      && h.indexOf("1 - ") === 0;
}

// =====================================================================
// LANGKAH 2 — Jalankan fungsi ini untuk MEMPERBAIKI formula.
//             Otomatis mendeteksi lokasi formula dan menulis ulang.
// =====================================================================
function fixStatusAHASS() {
  Logger.log("=== FIX STATUS AHASS ===\n");

  // --- Cek For AHM Data Mentah ---
  var ahmSS    = SpreadsheetApp.openById(AHM_SPREADSHEET_ID);
  var ahmSheet = ahmSS.getSheetByName("Data Mentah");
  if (!ahmSheet) throw new Error("Sheet 'Data Mentah' tidak ditemukan di For AHM.");

  var ahmJ2f = ahmSheet.getRange("J2").getFormula();
  var ahmJ1f = ahmSheet.getRange("J1").getFormula();
  var ahmA1f = ahmSheet.getRange("A1").getFormula();

  // Kasus 1: Formula individu di setiap sel (J2, J3, …)
  if (ahmJ2f) {
    Logger.log("Formula individu ditemukan di For AHM J2: " + ahmJ2f);
    rewriteIndividualFormulas_(ahmSheet, 2, "For AHM Data Mentah");
    return;
  }

  // Kasus 2: ARRAYFORMULA di J1
  if (ahmJ1f && ahmJ1f.toUpperCase().indexOf("ARRAYFORMULA") >= 0) {
    Logger.log("ARRAYFORMULA ditemukan di For AHM J1: " + ahmJ1f);
    rewriteArrayFormula_(ahmSheet, true, "For AHM Data Mentah");
    return;
  }

  // Kasus 3: Kolom J merupakan bagian dari QUERY/IMPORTRANGE di A1
  // → Sumber data harus diperbaiki di Master Gabungan
  if (ahmA1f) {
    Logger.log("Kolom J di For AHM berasal dari QUERY/IMPORTRANGE.");
    Logger.log("Formula A1: " + ahmA1f.substring(0, 300));
    Logger.log("Beralih ke Master Gabungan untuk perbaikan...\n");
  }

  // --- Cek Master Gabungan ---
  var masterSS  = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var gabSheet  = masterSS.getSheetByName("Gabungan");
  if (!gabSheet) throw new Error("Sheet 'Gabungan' tidak ditemukan di Master.");

  var gabJ1f = gabSheet.getRange("J1").getFormula();
  var gabJ2f = gabSheet.getRange("J2").getFormula();

  // Gabungan TIDAK punya header row — data mulai dari baris 1
  if (gabJ1f && gabJ1f.toUpperCase().indexOf("ARRAYFORMULA") >= 0) {
    Logger.log("ARRAYFORMULA ditemukan di Gabungan J1: " + gabJ1f);
    rewriteArrayFormula_(gabSheet, false, "Master Gabungan");
    return;
  }

  if (gabJ1f) {
    Logger.log("Formula individu ditemukan di Gabungan J1: " + gabJ1f);
    rewriteIndividualFormulas_(gabSheet, 1, "Master Gabungan");
    return;
  }

  if (gabJ2f) {
    Logger.log("Formula individu ditemukan di Gabungan J2: " + gabJ2f);
    rewriteIndividualFormulas_(gabSheet, 2, "Master Gabungan");
    return;
  }

  // Kasus 4: Tidak ada formula di J (data di-generate oleh Apps Script atau macro lain)
  Logger.log("TIDAK ditemukan formula di kolom J di kedua sheet.");
  Logger.log("Menulis formula baru di For AHM Data Mentah...");
  rewriteIndividualFormulas_(ahmSheet, 2, "For AHM Data Mentah");
}

// ----- helper: tulis formula individu per baris -----
function rewriteIndividualFormulas_(sheet, startRow, label) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) { Logger.log("Tidak ada data di " + label); return; }

  var count = lastRow - startRow + 1;
  var formulas = [];
  for (var r = startRow; r <= lastRow; r++) {
    formulas.push([
      '=IF(A' + r + '="",' +
        '"",' +
        'IF(AND(' +
          'LEFT(D' + r + ',4)="1 - ",' +
          'LEFT(E' + r + ',5)="23 - ",' +
          'LEFT(F' + r + ',4)="1 - ",' +
          'LEFT(G' + r + ',4)="1 - ",' +
          'LEFT(H' + r + ',4)="1 - "),' +
        '"OK","NG"))'
    ]);
  }

  sheet.getRange(startRow, 10, count, 1).setFormulas(formulas);
  Logger.log("[" + label + "] Formula ditulis ulang untuk " + count + " baris (J" + startRow + ":J" + lastRow + ").");
  Logger.log("Contoh J" + startRow + ": " + formulas[0][0]);

  verifyAfterFix_(sheet, startRow, label);
}

// ----- helper: tulis ARRAYFORMULA di J1 -----
function rewriteArrayFormula_(sheet, hasHeader, label) {
  if (hasHeader) {
    // Header di baris 1, data mulai baris 2
    var formula =
      '={"Status AHASS ";' +
      'ARRAYFORMULA(IF(A2:A="",' +
        '"",' +
        'IF(' +
          '(LEFT(D2:D,4)="1 - ")*' +
          '(LEFT(E2:E,5)="23 - ")*' +
          '(LEFT(F2:F,4)="1 - ")*' +
          '(LEFT(G2:G,4)="1 - ")*' +
          '(LEFT(H2:H,4)="1 - "),' +
        '"OK","NG")))}';
    sheet.getRange("J1").setFormula(formula);
    Logger.log("[" + label + "] ARRAYFORMULA diperbaiki di J1.");
    Logger.log("Formula: " + formula);
    verifyAfterFix_(sheet, 2, label);
  } else {
    // Tanpa header, data mulai baris 1
    var formula =
      'ARRAYFORMULA(IF(A1:A="",' +
        '"",' +
        'IF(' +
          '(LEFT(D1:D,4)="1 - ")*' +
          '(LEFT(E1:E,5)="23 - ")*' +
          '(LEFT(F1:F,4)="1 - ")*' +
          '(LEFT(G1:G,4)="1 - ")*' +
          '(LEFT(H1:H,4)="1 - "),' +
        '"OK","NG")))';
    sheet.getRange("J1").setFormula("=" + formula);
    Logger.log("[" + label + "] ARRAYFORMULA diperbaiki di J1.");
    Logger.log("Formula: =" + formula);
    verifyAfterFix_(sheet, 1, label);
  }
}

// ----- helper: verifikasi setelah fix -----
function verifyAfterFix_(sheet, startRow, label) {
  SpreadsheetApp.flush();
  var rowCount = Math.min(10, sheet.getLastRow() - startRow + 1);
  if (rowCount < 1) return;

  var data = sheet.getRange(startRow, 4, rowCount, 7).getValues(); // D:J
  Logger.log("\nVerifikasi setelah perbaikan (" + label + "):");
  var allCorrect = true;
  for (var i = 0; i < data.length; i++) {
    var expected = allColumnsOK_(data[i]) ? "OK" : "NG";
    var actual   = String(data[i][6]).trim();
    var match    = expected === actual || (actual === "" && String(data[i][0]) === "");
    if (!match) allCorrect = false;
    Logger.log("  Row " + (startRow + i) + ": expected=" + expected + " actual=" + actual
      + (match ? "  ✓" : "  ← MASIH SALAH"));
  }
  Logger.log(allCorrect
    ? "\n✓ SEMUA baris yang dicek sudah BENAR."
    : "\n✗ Masih ada ketidakcocokan. Periksa log di atas.");
}
