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
const SOURCE_SHEET_NAME = 'PUD';          // sheet sumber. null = pakai sheet pertama.
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

/**
 * Formatting per sheet kategori (semua sheet selain SOURCE_SHEET_NAME):
 *   - Terapkan warna & lock HANYA pada area tabel (baris 1 s/d lastRow).
 *   - Kolom A-D  : background #ffe2ca (peach) + dikunci
 *   - Kolom E-K  : background #b3d7ef (biru muda), bisa diedit
 *   - Kolom L    : background #ffe2ca (peach) + dikunci
 *   - Tambah border tipis di seluruh area tabel supaya terlihat rapi.
 *   - Background di luar area tabel dibersihkan.
 * Aman dijalankan berulang: proteksi lama dihapus dulu.
 */
function formatCategorySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcName = SOURCE_SHEET_NAME || ss.getSheets()[0].getName();
  const me = Session.getEffectiveUser();
  const PEACH = '#ffe2ca';
  const BLUE = '#b3d7ef';
  let count = 0;

  ss.getSheets().forEach(sh => {
    if (sh.getName() === srcName) return;
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;

    // Bersihkan semua background dulu (menghapus sisa run sebelumnya
    // yang mewarnai seluruh kolom di luar area tabel).
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setBackground(null);

    // A-D (peach) — hanya di area tabel.
    const adEnd = Math.min(4, lastCol);
    sh.getRange(1, 1, lastRow, adEnd).setBackground(PEACH);

    // E-K (biru).
    if (lastCol >= 5) {
      const ekEnd = Math.min(11, lastCol);
      sh.getRange(1, 5, lastRow, ekEnd - 5 + 1).setBackground(BLUE);
    }

    // L (peach).
    if (lastCol >= 12) {
      sh.getRange(1, 12, lastRow, 1).setBackground(PEACH);
    }

    // Border di seluruh area tabel supaya terlihat seperti tabel.
    sh.getRange(1, 1, lastRow, lastCol)
      .setBorder(true, true, true, true, true, true);

    // Hapus proteksi lama biar idempotent.
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    // Lock A-D dan L (hanya baris tabel).
    const lockRanges = [sh.getRange(1, 1, lastRow, adEnd)];
    if (lastCol >= 12) lockRanges.push(sh.getRange(1, 12, lastRow, 1));

    lockRanges.forEach(rng => {
      const p = rng.protect()
        .setDescription('locked (splitDatabasePengerjaan)');
      p.removeEditors(p.getEditors());
      p.addEditor(me);
      if (p.canDomainEdit()) p.setDomainEdit(false);
      p.setWarningOnly(false);
    });

    count++;
  });

  ss.toast(`Selesai. Formatting & lock diterapkan ke ${count} sheet.`,
    'formatCategorySheets', 10);
}

/**
 * Buka kunci (unlock) semua sheet: hapus proteksi sheet & range.
 * Jalankan manual dari editor Apps Script kalau mau semua sheet bisa
 * diedit lagi.
 */
function unlockAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let removed = 0;
  ss.getSheets().forEach(sh => {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(p => { p.remove(); removed++; });
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .forEach(p => { p.remove(); removed++; });
  });
  ss.toast(`Selesai. ${removed} proteksi dihapus.`, 'unlockAllSheets', 10);
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
