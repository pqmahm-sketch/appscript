/**
 * splitDatabasePengerjaan.gs
 *
 * Skrip untuk memproses spreadsheet "Database Pengerjaan PUD - Copy":
 *   1. Menyaring baris pada kolom A (Nama PUD) hanya untuk PUD yang di-allow.
 *   2. Menghapus kolom G-I dan M-O (jika ada).
 *   3. Membuat sheet baru untuk setiap kategori unik di kolom CATEGORY_COL
 *      dan mengisinya dengan baris yang cocok.
 *   4. Mem-protect seluruh sheet supaya tidak bisa diubah orang lain.
 *
 * Cara pakai:
 *   - Buka spreadsheet copy: Extensions > Apps Script.
 *   - Paste seluruh file ini ke editor, save.
 *   - Jalankan fungsi runAll() satu kali (beri izin saat diminta).
 *
 * CATATAN PENTING soal "29 kategori di kolom C":
 *   Pada file sumber, kolom C = "No Rangka" yang berisi 3245 nilai unik,
 *   jadi tidak mungkin menghasilkan 29 sheet. Kolom yang jumlah kategorinya
 *   paling dekat dengan 29 adalah kolom B "Kode MD Unit Distribusi" (31 unik)
 *   atau kolom F "Kode MD Pelapor" (31 unik). Jika Anda memang mau membuat
 *   sheet per Kode MD, ubah CATEGORY_COL di bawah menjadi "B" (atau "F").
 *   Default dibiarkan "C" mengikuti instruksi apa adanya.
 */

// ====== KONFIGURASI ======
const SOURCE_SHEET_NAME = null;   // null = pakai sheet pertama; isi nama jika perlu.
const ALLOWED_PUDS_REGEX = /(CBR250|CUV1|CBR600|CRF1100|AT3|AT4|AT5|GL1800|CBR1000)/i;
const COLS_TO_DELETE = ['G-I', 'M-O']; // range huruf kolom yang akan dihapus.
const CATEGORY_COL = 'C';         // kolom sumber kategori (setelah penghapusan).
const MAX_SHEETS_HARDLIMIT = 60;  // pengaman: batalkan jika akan bikin > N sheet.
const PROTECT_SHEETS = true;      // set false kalau tidak ingin proteksi.

// =========================

function runAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = SOURCE_SHEET_NAME
    ? ss.getSheetByName(SOURCE_SHEET_NAME)
    : ss.getSheets()[0];

  if (!src) throw new Error('Sheet sumber tidak ditemukan.');

  filterAllowedPuds_(src);
  deleteColumns_(src, COLS_TO_DELETE);
  const categories = createCategorySheets_(ss, src, CATEGORY_COL);
  if (PROTECT_SHEETS) protectAllSheets_(ss);

  SpreadsheetApp.getActive().toast(
    `Selesai. ${categories.length} sheet kategori dibuat.`,
    'splitDatabasePengerjaan',
    10
  );
}

// --- Langkah 2: filter PUD ---
function filterAllowedPuds_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowsToDelete = [];
  for (let i = 0; i < values.length; i++) {
    const pud = String(values[i][0] || '');
    if (!ALLOWED_PUDS_REGEX.test(pud)) rowsToDelete.push(i + 2);
  }
  // hapus dari bawah ke atas
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}

// --- Langkah 3: hapus kolom G-I dan M-O ---
function deleteColumns_(sheet, ranges) {
  // Kumpulkan indeks kolom (1-based) yang akan dihapus.
  const toDelete = new Set();
  ranges.forEach(r => {
    const [a, b] = r.split('-');
    const start = colLetterToNum_(a);
    const end = b ? colLetterToNum_(b) : start;
    for (let c = start; c <= end; c++) toDelete.add(c);
  });
  const lastCol = sheet.getLastColumn();
  const sorted = [...toDelete].filter(c => c <= lastCol).sort((x, y) => y - x);
  sorted.forEach(c => sheet.deleteColumn(c));
}

// --- Langkah 4-5: bikin sheet per kategori + isi datanya ---
function createCategorySheets_(ss, src, catColLetter) {
  const catIdx = colLetterToNum_(catColLetter) - 1;
  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow < 2) return [];

  const header = src.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = src.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Kelompokkan baris per nilai kategori.
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
      `Akan membuat ${cats.length} sheet — melebihi MAX_SHEETS_HARDLIMIT (${MAX_SHEETS_HARDLIMIT}). ` +
      `Cek CATEGORY_COL — mungkin salah kolom. Kolom C = "No Rangka" (3245 unik), ` +
      `pertimbangkan kolom B "Kode MD Unit Distribusi" (31 unik).`
    );
  }

  cats.forEach(cat => {
    const name = sanitizeSheetName_(cat);
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
    sh = ss.insertSheet(name);
    const rows = buckets.get(cat);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, header.length);
  });

  return cats;
}

// --- Langkah 6: proteksi semua sheet ---
function protectAllSheets_(ss) {
  const me = Session.getEffectiveUser();
  ss.getSheets().forEach(sh => {
    // hapus proteksi lama biar idempotent
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(p => p.remove());
    const p = sh.protect().setDescription('Locked by splitDatabasePengerjaan');
    p.removeEditors(p.getEditors());
    p.addEditor(me);
    if (p.canDomainEdit()) p.setDomainEdit(false);
    p.setWarningOnly(false);
  });
}

// --- Utilitas ---
function colLetterToNum_(letters) {
  let n = 0;
  const s = String(letters).toUpperCase();
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function sanitizeSheetName_(name) {
  // Sheet name: max 100 char, tidak boleh ada : \ / ? * [ ]
  let clean = String(name).replace(/[:\\\/\?\*\[\]]/g, '_').trim();
  if (!clean) clean = 'kategori';
  if (clean.length > 100) clean = clean.slice(0, 100);
  return clean;
}
