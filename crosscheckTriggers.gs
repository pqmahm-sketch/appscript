/**
 * crosscheckTriggers.gs
 * Installer trigger untuk sistem crosscheck AMORE.
 *
 * Jalankan installTriggersCrosscheck() SEKALI dari editor Apps Script:
 *   - Kamis 09:00 → runCrosscheckCollector
 *   - Jumat 15:30 → runCrosscheckComparison
 *
 * Trigger lama dengan nama fungsi yang sama akan dihapus terlebih dulu agar tidak dobel.
 */

function installTriggersCrosscheck() {
  const targets = ['runCrosscheckCollector', 'runCrosscheckComparison'];
  // Hapus trigger lama untuk kedua fungsi tsb
  ScriptApp.getProjectTriggers().forEach(t => {
    if (targets.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Kamis 09:00 → collector
  ScriptApp.newTrigger('runCrosscheckCollector')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(9)
    .nearMinute(0)
    .create();

  // Jumat 15:30 → comparison
  // Apps Script tidak punya nearMinute untuk resolusi 15 menit selain 0/15/30/45,
  // dan kita ingin 15:30. atHour(15).nearMinute(30) OK.
  ScriptApp.newTrigger('runCrosscheckComparison')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(15)
    .nearMinute(30)
    .create();

  Logger.log('Trigger terpasang: Kamis 09:00 (collector), Jumat 15:30 (comparison).');
  try {
    SpreadsheetApp.getUi().alert('Trigger terpasang!\n\n- Kamis 09:00: runCrosscheckCollector\n- Jumat 15:30: runCrosscheckComparison');
  } catch (e) {
    // Non-UI context (mis. dijalankan dari script editor tanpa spreadsheet aktif)
  }
}

/**
 * Lihat daftar trigger yang aktif untuk kedua fungsi crosscheck.
 */
function listCrosscheckTriggers() {
  const targets = ['runCrosscheckCollector', 'runCrosscheckComparison'];
  const arr = ScriptApp.getProjectTriggers()
    .filter(t => targets.indexOf(t.getHandlerFunction()) >= 0)
    .map(t => ({
      fn: t.getHandlerFunction(),
      type: t.getTriggerSource(),
      eventType: String(t.getEventType())
    }));
  Logger.log(JSON.stringify(arr, null, 2));
  return arr;
}

/**
 * Hapus semua trigger untuk kedua fungsi crosscheck.
 */
function uninstallTriggersCrosscheck() {
  const targets = ['runCrosscheckCollector', 'runCrosscheckComparison'];
  const before = ScriptApp.getProjectTriggers().length;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (targets.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  const after = ScriptApp.getProjectTriggers().length;
  Logger.log('Trigger dihapus: ' + (before - after));
}
