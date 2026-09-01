// ====== KONFIGURASI ======
const BACKUP_SPREADSHEET_ID = '1bC7tJPVoCz6x_lqq1kHbFEO15FRoyjdV6iQ9zMzAFrA';
const BACKUP_FOLDER_NAME = 'Backup Status Follow Up PUD';
const BACKUP_NAME_PREFIX = 'Status Follow Up PUD';
const BACKUP_HOUR = 2;

// ====== BACKUP BULANAN ======

/**
 * Membuat trigger yang menjalankan monthlyBackup() setiap tanggal 1
 * jam BACKUP_HOUR. Menghapus trigger lama dengan handler yang sama.
 */
function setupMonthlyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'monthlyBackup') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('monthlyBackup')
    .timeBased()
    .onMonthDay(1)
    .atHour(BACKUP_HOUR)
    .create();
}

/**
 * Salin spreadsheet target (BACKUP_SPREADSHEET_ID) sebagai .xlsx
 * ke folder BACKUP_FOLDER_NAME.
 * Nama file: "Status Follow Up PUD_[MMM]_[YYYY].xlsx"
 *   MMM  = bulan lalu (English 3-huruf: Jan, Feb, ..., Dec)
 *   YYYY = tahun dari bulan lalu (menangani rollover Januari)
 */
function monthlyBackup() {
  const ss = SpreadsheetApp.openById(BACKUP_SPREADSHEET_ID);
  const now = new Date();
  // Bulan lalu, tanggal 1 — otomatis handle rollover Januari.
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const mmm = MONTH_ABBR[lastMonth.getMonth()];
  const yyyy = lastMonth.getFullYear();
  const filename = `${BACKUP_NAME_PREFIX}_${mmm}_${yyyy}.xlsx`;

  const folder = getBackupFolder_(BACKUP_FOLDER_NAME);

  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error(`Gagal export xlsx: HTTP ${resp.getResponseCode()}`);
  }
  const blob = resp.getBlob().setName(filename);

  // Hapus file dengan nama sama di folder (biar idempotent).
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const file = folder.createFile(blob);
  Logger.log(`Backup dibuat: ${file.getUrl()}`);
  return file.getId();
}

function getBackupFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (!folders.hasNext()) {
    throw new Error(`Folder "${name}" tidak ditemukan di Drive Anda.`);
  }
  return folders.next();
}

// ====== UTILITAS ======
function colLetterToNum_(letters) {
  let n = 0;
  const s = String(letters).toUpperCase();
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function sanitizeSheetName_(name) {
  let clean = String(name).replace(/[:\\\/\?\*\[\]]/g, '_').trim();
  if (!clean) clean = 'kategori';
  if (clean.length > 100) clean = clean.slice(0, 100);
  return clean;
}
