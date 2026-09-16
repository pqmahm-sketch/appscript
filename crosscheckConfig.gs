/**
 * crosscheckConfig.gs
 * Konfigurasi bersama untuk sistem crosscheck point + komparasi NKH vs GKA.
 * Semua konstanta terpusat di sini agar mudah diubah tanpa menyentuh logic.
 */

const CROSSCHECK_CONFIG = {
  // -------- Sumber data --------
  SOURCE_SPREADSHEET_ID: '1uU3mG1tu-d9aW1Z-eL-ocOQHhMzg2xPWAN4DvSYmo1o',
  SOURCE_SHEET_NAME: 'Data Mentah',

  // Kolom pada Data Mentah (dicari via header, fallback ke posisi)
  COL_MAIN_DEALER_HEADER: 'Main Dealer',      // fallback B (index 1)
  COL_STATUS_AHASS_HEADER: 'Status AHASS',    // fallback dicari otomatis
  COL_NO_RANGKA_HEADER: 'No. Rangka',
  COL_NO_CLAIM_HEADER: 'No. Claim',

  // Daftar Main Dealer yang di-scope
  MAIN_DEALERS: [
    'B3Z', 'B10', 'C10', 'C3Z', 'D2Z', 'D3Z',
    'E20', 'G01', 'G02', 'G5Z', 'H2Z',
    'I01', 'I3Z', 'J10', 'J20'
  ],

  // Sampling per MD (2 baris per MD, prefer mix OK+NG bila tersedia)
  ROWS_PER_MD: 2,
  STATUS_OK: 'OK',
  STATUS_NG: 'NG',

  // -------- Output crosscheck point --------
  OUTPUT_FOLDER_ID: '14XGGHwrEeb__xI73OAvZHMCOpNxs5kyM',
  CROSSCHECK_FILE_PREFIX: 'crosscheck point ',
  CROSSCHECK_START_NUMBER: 3,     // file pertama yang akan dibuat = crosscheck point 3

  // -------- Sumber komparasi NKH vs GKA --------
  NKH_SPREADSHEET_ID: '1sKsEPC9TAAEnb8WsAqKRJtH0Y-U9izCqf-Pcpf7jPgw',
  GKA_SPREADSHEET_ID: '1Epr6jzg8XyxyqYU_pYkQ7-Xvz_elRaCOl8StHVhVv3g',
  RESPONSE_SHEET_NAME: 'Form Responses 1',

  // Kolom-kolom yang dibandingkan di NKH & GKA (nama header pada Form Responses 1)
  COMPARE_COLUMNS: [
    { letter: 'D', header: 'Kategori video treatment LCR',    subLabel: 'Pemilihan Kategori Video' },
    { letter: 'E', header: 'Proses Aplikasi Flintkote',       subLabel: 'Proses Aplikasi Flintkote' },
    { letter: 'F', header: 'Proses Setting Regulator',        subLabel: 'Proses Setting Regulator' },
    { letter: 'G', header: 'Proses Timer',                    subLabel: 'Proses Timer' },
    { letter: 'H', header: 'Marking garis pada selang',       subLabel: 'Marking garis pada selang' },
    { letter: 'I', header: 'Proses Marking Frame Body',       subLabel: 'Proses Marking Frame Body' },
    { letter: 'J', header: 'Catatan Tambahan',                subLabel: 'Catatan Tambahan' },
    { letter: 'K', header: 'No. Rangka Proses Flinkote',      subLabel: 'No. Rangka Proses Flinkote' },
    { letter: 'L', header: 'No. Rangka Proses Noxudol',       subLabel: 'No. Rangka Proses Noxudol' },
    { letter: 'M', header: 'No. Rangka Proses Marking',       subLabel: 'No. Rangka Proses Marking' },
    { letter: 'N', header: 'Indikasi Masalah',                subLabel: 'Indikasi Masalah' },
    { letter: 'O', header: 'Anomali video yang ditemukan',    subLabel: 'Anomali video yang ditemukan' },
    { letter: 'P', header: 'Deskripsikan Level Karat',        subLabel: 'Deskripsikan Level Karat' }
  ],

  // -------- Output komparasi --------
  COMPARISON_FILE_PREFIX: 'Komparasi Penilaian AMORE NKH vs GKA - ',

  // -------- Email --------
  // GANTI dengan email penerima yang sesungguhnya sebelum deploy.
  EMAIL_TO: 'pqm.ahm@gmail.com',
  EMAIL_CC: '',
  EMAIL_SENDER_NAME: 'PQM AHM - Auto Crosscheck'
};

/**
 * Helper: cari index kolom di header row (case-insensitive, trim).
 * Return -1 jika tidak ditemukan.
 */
function findColumnIndex_(headerRow, targetHeader) {
  const target = String(targetHeader || '').trim().toLowerCase();
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || '').trim().toLowerCase() === target) return i;
  }
  return -1;
}

/**
 * Helper: normalisasi nilai (buang whitespace, upper untuk rangka).
 */
function normUpper_(v) {
  return String(v == null ? '' : v).trim().toUpperCase();
}

function normTrim_(v) {
  return String(v == null ? '' : v).trim();
}
