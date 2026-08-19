/**
 * splitDatabasePengerjaan.gs
 *
 * Untuk spreadsheet "Status Follow Up PUD" (copy dari Database Pengerjaan PUD).
 *
 * Fungsi:
 *   runAll()
 *     1. TIDAK mengubah kolom apapun di sheet original.
 *     2. Membuat satu sheet tambahan per kategori unik di kolom B.
 *     3. Mengisi tiap sheet baru dengan baris yang cocok, memakai header
 *        dan struktur kolom yang persis sama dengan sheet original.
 *     4. Mem-protect semua sheet supaya tidak bisa diedit orang lain.
 *
 *   setupMonthlyBackupTrigger()
 *     Membuat trigger otomatis: tiap tanggal 1 jam 07:00 memanggil
 *     monthlyBackup() yang menyimpan salinan file ini sebagai .xlsx
 *     ke folder "16. Database FU PUD" dengan nama
 *     "Status Follow Up PUD_[MMM]_[YYYY].xlsx" (MMM & YYYY = bulan lalu).
 *
 *   monthlyBackup()
 *     Bisa dijalankan manual kapan saja untuk backup segera.
 *
 * Cara pakai pertama kali:
 *   - Extensions > Apps Script > paste seluruh file ini > Save.
 *   - Jalankan runAll() sekali (izin akan diminta).
 *   - Jalankan setupMonthlyBackupTrigger() sekali (izin trigger diminta).
 */

// ====== KONFIGURASI ======
const SOURCE_SHEET_NAME = null;           // null = pakai sheet pertama.
const CATEGORY_COL = 'B';                 // acuan kategori split.
const PROTECT_SHEETS = true;
const MAX_SHEETS_HARDLIMIT = 60;          // pengaman.
const BACKUP_FOLDER_NAME = '16. Database FU PUD';
const BACKUP_NAME_PREFIX = 'Status Follow Up PUD';
const BACKUP_HOUR = 7;                    // trigger jam berapa (0-23).
// =========================

function runAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = getSourceSheet_(ss);
  const cats = createCategorySheets_(ss, src, CATEGORY_COL);
  if (PROTECT_SHEETS) protectAllSheets_(ss);

  ss.toast(
    `Selesai. ${cats.length} sheet kategori dibuat dari kolom ${CATEGORY_COL}.`,
    'splitDatabasePengerjaan',
    10
  );
}

function getSourceSheet_(ss) {
  const src = SOURCE_SHEET_NAME
    ? ss.getSheetByName(SOURCE_SHEET_NAME)
    : ss.getSheets()[0];
  if (!src) throw new Error('Sheet sumber tidak ditemukan.');
  return src;
}

// --- Buat sheet per kategori (kolom B) ---
function createCategorySheets_(ss, src, catColLetter) {
  const catIdx = colLetterToNum_(catColLetter) - 1;
  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < 2) return [];

  const header = src.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = src.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const buckets = new Map();
  data.forEach(row => {
    const key = String(row[catIdx] == null ? '' : row[catIdx]).trim();
    if (!key) return;
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(row);
  });

  const cats = [...buckets.keys()].sort();

  if (cats.length > MAX_SHEETS_HARDLIMIT) {
    throw new Error(
      `Akan membuat ${cats.length} sheet dari kolom ${catColLetter} ` +
      `("${header[catIdx]}") — melebihi MAX_SHEETS_HARDLIMIT (${MAX_SHEETS_HARDLIMIT}).`
    );
  }

  const srcName = src.getName();
  cats.forEach(cat => {
    const name = sanitizeSheetName_(cat);
    if (name === srcName) return; // jangan tabrakan dengan sheet sumber
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
    sh = ss.insertSheet(name);
    const rows = buckets.get(cat);
    sh.getRange(1, 1, 1, header.length)
      .setValues([header])
      .setFontWeight('bold');
    if (rows.length) {
      sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    }
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, header.length);
  });

  return cats;
}

// --- Proteksi semua sheet ---
function protectAllSheets_(ss) {
  const me = Session.getEffectiveUser();
  ss.getSheets().forEach(sh => {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(p => p.remove());
    const p = sh.protect().setDescription('Locked by splitDatabasePengerjaan');
    p.removeEditors(p.getEditors());
    p.addEditor(me);
    if (p.canDomainEdit()) p.setDomainEdit(false);
    p.setWarningOnly(false);
  });
}

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
 * Salin spreadsheet ini sebagai .xlsx ke folder BACKUP_FOLDER_NAME.
 * Nama file: "Status Follow Up PUD_[MMM]_[YYYY].xlsx"
 *   MMM  = bulan lalu (English 3-huruf: Jan, Feb, ..., Dec)
 *   YYYY = tahun dari bulan lalu (menangani rollover Januari)
 */
function monthlyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
