// ============================================================
// FUNCTION 1: Cek duplikat AHASS & hapus yang NG
// ============================================================
function checkDuplicateAHASS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form Responses 1");

  const lastRow = formSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("Tidak ada cukup data untuk diperiksa.");
    return;
  }

  const formData = formSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  const ahassMap = {};
  for (let i = 0; i < formData.length; i++) {
    const noAHASS = formData[i][1];
    if (!noAHASS) continue;

    const key = noAHASS.toString().trim();
    if (!ahassMap[key]) ahassMap[key] = [];
    ahassMap[key].push({
      rowNum: i + 2,
      hasil: formData[i][3].toString().trim().toUpperCase(),
      timestamp: new Date(formData[i][0])
    });
  }

  const rowsToDelete = [];

  for (const ahass in ahassMap) {
    const rows = ahassMap[ahass];
    if (rows.length < 2) continue;

    const hasOK = rows.some(r => r.hasil === "OK");

    if (hasOK) {
      rows.forEach(row => {
        if (row.hasil === "NG") {
          rowsToDelete.push(row.rowNum);
          Logger.log(`AHASS ${ahass} → Baris ${row.rowNum} (NG) dihapus karena ada yang OK.`);
        }
      });
    } else {
      rows.sort((a, b) => b.timestamp - a.timestamp);
      for (let j = 1; j < rows.length; j++) {
        rowsToDelete.push(rows[j].rowNum);
        Logger.log(`AHASS ${ahass} → Semua NG. Baris ${rows[j].rowNum} (lama: ${rows[j].timestamp}) dihapus, pertahankan baris ${rows[0].rowNum} (terbaru: ${rows[0].timestamp}).`);
      }
    }
  }

  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(rowNum => {
    formSheet.deleteRow(rowNum);
    Logger.log(`Baris ${rowNum} berhasil dihapus.`);
  });

  Logger.log(`Selesai. Total ${rowsToDelete.length} baris dihapus.`);
}

// ============================================================
// FUNCTION 2: Pasang trigger otomatis setiap hari jam 08:00
// Jalankan function ini SEKALI SAJA untuk aktivasi
// ============================================================
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkDuplicateAHASS") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("checkDuplicateAHASS")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log("Trigger harian berhasil dipasang! Akan berjalan setiap hari jam 08:00.");
}
