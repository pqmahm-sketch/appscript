/**
 * crosscheckCollector.gs
 * PROSES 1: Kumpulkan data & buat file "crosscheck point N".
 *
 * Jalankan: runCrosscheckCollector()   (manual)
 * Trigger : dipasang otomatis oleh installTriggersCrosscheck() → tiap Kamis 09:00.
 *
 * Logika:
 *   1. Buka spreadsheet sumber → sheet "Data Mentah".
 *   2. Filter baris berdasarkan kolom Main Dealer (15 inisial).
 *   3. Filter berdasarkan Status AHASS (OK/NG), ambil 2 baris per MD (prefer mix 1 OK + 1 NG).
 *   4. Buang rangka+claim yang sudah muncul di file crosscheck sebelumnya.
 *   5. Buat spreadsheet baru "crosscheck point N" di folder output.
 */

function runCrosscheckCollector() {
  const cfg = CROSSCHECK_CONFIG;
  const startTs = new Date();
  Logger.log('=== runCrosscheckCollector START @ ' + startTs);

  // --- Step 1: baca sheet sumber ---
  const srcSs = SpreadsheetApp.openById(cfg.SOURCE_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(cfg.SOURCE_SHEET_NAME);
  if (!srcSheet) throw new Error('Sheet "' + cfg.SOURCE_SHEET_NAME + '" tidak ditemukan pada sumber.');

  const values = srcSheet.getDataRange().getValues();
  if (values.length < 2) {
    Logger.log('Data Mentah kosong / hanya header.');
    return null;
  }

  // --- Auto-detect baris header (header bisa berada di row 3 pada sumber ini) ---
  // Scan 10 baris pertama, cari baris yang mengandung "No. Rangka" (case-insensitive).
  const HEADER_SCAN_ROWS = Math.min(10, values.length);
  let headerRowIdx = -1;
  for (let r = 0; r < HEADER_SCAN_ROWS; r++) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim().toLowerCase();
      if (cell === 'no. rangka' || cell === 'no rangka' || cell === 'no.rangka') {
        headerRowIdx = r;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }
  if (headerRowIdx < 0) {
    // Fallback ke row index 2 (baris ke-3) sesuai info user
    headerRowIdx = 2;
    Logger.log('Header row tidak terdeteksi otomatis, pakai fallback row 3 (index 2).');
  } else {
    Logger.log('Header row terdeteksi di baris ke-' + (headerRowIdx + 1) + ' (index ' + headerRowIdx + ').');
  }

  const header = (values[headerRowIdx] || []).map(v => String(v || '').trim());

  // Prioritas urutan resolve: (1) letter override di config, (2) header exact match,
  // (3) variasi header, (4) fallback khusus, (5) scan data.
  // Setiap kolom yang sudah "diambil" ditandai agar tidak dipakai kolom lain.
  const taken = new Set();
  function claim(idx) { if (idx >= 0) taken.add(idx); return idx; }
  function firstFreeMatch(fn) {
    for (let i = 0; i < header.length; i++) {
      if (taken.has(i)) continue;
      if (fn(header[i], i)) return i;
    }
    return -1;
  }

  // MD
  let idxMD = letterToIndex_(cfg.COL_MAIN_DEALER_LETTER);
  if (idxMD < 0) idxMD = findColumnIndex_(header, cfg.COL_MAIN_DEALER_HEADER);
  if (idxMD < 0) idxMD = 1;
  claim(idxMD);

  // Rangka
  let idxRangka = letterToIndex_(cfg.COL_NO_RANGKA_LETTER);
  if (idxRangka < 0) idxRangka = findColumnIndex_(header, cfg.COL_NO_RANGKA_HEADER);
  if (idxRangka < 0) idxRangka = firstFreeMatch(h => /rangka/i.test(h));
  if (idxRangka < 0) idxRangka = 2;
  claim(idxRangka);

  // Status
  let idxStatus = letterToIndex_(cfg.COL_STATUS_AHASS_LETTER);
  if (idxStatus < 0) idxStatus = findColumnIndex_(header, cfg.COL_STATUS_AHASS_HEADER);
  if (idxStatus < 0) idxStatus = firstFreeMatch(h => /^status\b/i.test(h) || /status\s*ahass/i.test(h));
  if (idxStatus < 0) {
    // Scan baris data pertama cari OK/NG di kolom yang belum di-claim
    const firstData = values[headerRowIdx + 1] || [];
    for (let c = 0; c < firstData.length; c++) {
      if (taken.has(c)) continue;
      const v = String(firstData[c] || '').trim().toUpperCase();
      if (v === 'OK' || v === 'NG') { idxStatus = c; break; }
    }
  }
  if (idxStatus < 0) throw new Error('Kolom Status AHASS tidak ditemukan.');
  claim(idxStatus);

  // Claim
  let idxClaim = letterToIndex_(cfg.COL_NO_CLAIM_LETTER);
  if (idxClaim < 0) idxClaim = findColumnIndex_(header, cfg.COL_NO_CLAIM_HEADER);
  if (idxClaim < 0) idxClaim = firstFreeMatch(h =>
    /no\.?\s*claim/i.test(h) || /nomor\s*claim/i.test(h) || /^claim\b/i.test(h)
  );
  if (idxClaim < 0) {
    // Scan baris data pertama cari sesuatu yang terlihat seperti nomor claim
    // (string panjang, alfanumerik), pada kolom yang belum di-claim.
    const firstData = values[headerRowIdx + 1] || [];
    for (let c = 0; c < firstData.length; c++) {
      if (taken.has(c)) continue;
      const v = String(firstData[c] || '').trim();
      if (v.length >= 6 && /[A-Za-z]/.test(v) && /\d/.test(v)) { idxClaim = c; break; }
    }
  }
  if (idxClaim < 0) throw new Error('Kolom No. Claim tidak ditemukan. Isi COL_NO_CLAIM_LETTER di crosscheckConfig.gs.');
  claim(idxClaim);

  // Cek collision defensif
  const picked_ = { MD: idxMD, Status: idxStatus, Rangka: idxRangka, Claim: idxClaim };
  const seen_ = {};
  for (const k in picked_) {
    const v = picked_[k];
    if (seen_[v]) throw new Error('Deteksi kolom collision: ' + k + ' dan ' + seen_[v] + ' sama-sama di kolom ' + (v + 1) + '. Isi override letter di crosscheckConfig.gs.');
    seen_[v] = k;
  }

  // Sanity check pesan
  Logger.log('Header row=' + (headerRowIdx + 1) +
             ', MD col=' + (idxMD + 1) +
             ', Status col=' + (idxStatus + 1) +
             ', Rangka col=' + (idxRangka + 1) +
             ', Claim col=' + (idxClaim + 1));

  // --- Step 2 & 3: filter & bucket per MD ---
  const mdSet = new Set(cfg.MAIN_DEALERS.map(x => String(x).trim().toUpperCase()));
  const buckets = {}; // { MD: { OK: [rows], NG: [rows] } }
  cfg.MAIN_DEALERS.forEach(md => { buckets[md.toUpperCase()] = { OK: [], NG: [] }; });

  // Diagnostics
  const seenMdCodes = {};      // code → count
  let totalData = 0;
  let matchedMd = 0;
  let matchedStatus = 0;

  // Data mulai dari baris SETELAH header
  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    // Extract kode MD dari format "K0Z - Astra Motor Semarang" (ambil bagian sebelum " - " atau token pertama)
    const mdRaw = String(row[idxMD] || '').trim().toUpperCase();
    if (!mdRaw) continue;
    totalData++;
    const mdMatch = mdRaw.match(/^([A-Z0-9]+)/);
    const md = mdMatch ? mdMatch[1] : mdRaw;
    seenMdCodes[md] = (seenMdCodes[md] || 0) + 1;
    if (!mdSet.has(md)) continue;
    matchedMd++;

    const status = String(row[idxStatus] || '').trim().toUpperCase();
    if (status !== cfg.STATUS_OK && status !== cfg.STATUS_NG) continue;
    matchedStatus++;

    const rangka = normUpper_(row[idxRangka]);
    const claim  = normTrim_(row[idxClaim]);
    if (!rangka || !claim) continue;

    buckets[md][status].push({ rangka: rangka, claim: claim, rowIndex: r });
  }

  // Log ringkasan diagnostik
  Logger.log('Total baris data: ' + totalData +
             ' | MD unik terlihat: ' + Object.keys(seenMdCodes).length +
             ' | Match MD config: ' + matchedMd +
             ' | Match Status OK/NG: ' + matchedStatus);
  Logger.log('MD codes ditemukan (top 25): ' +
    JSON.stringify(Object.keys(seenMdCodes)
      .sort((a,b) => seenMdCodes[b] - seenMdCodes[a])
      .slice(0, 25)
      .map(k => k + ':' + seenMdCodes[k])));
  // Bucket size per MD (hanya yang match config)
  const bucketSummary = {};
  Object.keys(buckets).forEach(md => {
    bucketSummary[md] = 'OK=' + buckets[md].OK.length + ' NG=' + buckets[md].NG.length;
  });
  Logger.log('Bucket per MD: ' + JSON.stringify(bucketSummary));

  // --- Step 4: exclude rangka/claim yang sudah pernah muncul di crosscheck sebelumnya ---
  const usedRangka = loadPreviousCrosscheckRangkas_();
  Logger.log('Previous crosscheck rangka count: ' + usedRangka.size);

  // --- Step 5: sampling 2 per MD, prefer 1 OK + 1 NG ---
  const picked = [];
  const seenThisRun = new Set();
  cfg.MAIN_DEALERS.forEach(mdRaw => {
    const md = mdRaw.toUpperCase();
    const bucket = buckets[md] || { OK: [], NG: [] };
    // Attach status label ke tiap item lalu filter yang belum dipakai
    const okList = bucket.OK
      .filter(x => !usedRangka.has(x.rangka) && !seenThisRun.has(x.rangka))
      .map(x => Object.assign({ status: cfg.STATUS_OK }, x));
    const ngList = bucket.NG
      .filter(x => !usedRangka.has(x.rangka) && !seenThisRun.has(x.rangka))
      .map(x => Object.assign({ status: cfg.STATUS_NG }, x));

    shuffleInPlace_(okList);
    shuffleInPlace_(ngList);

    // Prefer 1 OK + 1 NG
    const takeMd = [];
    if (okList.length > 0) takeMd.push(okList.shift());
    if (ngList.length > 0) takeMd.push(ngList.shift());

    // Kalau salah satu belum tersedia, isi dengan sisa dari list yang ada
    while (takeMd.length < cfg.ROWS_PER_MD) {
      if (okList.length > 0)      takeMd.push(okList.shift());
      else if (ngList.length > 0) takeMd.push(ngList.shift());
      else break;
    }

    takeMd.forEach(item => {
      seenThisRun.add(item.rangka);
      picked.push({ md: md, status: item.status, rangka: item.rangka, claim: item.claim });
    });
  });

  if (picked.length === 0) {
    Logger.log('Tidak ada data baru yang bisa diambil. Skip pembuatan file.');
    return null;
  }

  // --- Step 6: buat file crosscheck baru ---
  const nextNumber = getNextCrosscheckNumber_();
  const fileName = cfg.CROSSCHECK_FILE_PREFIX + nextNumber;

  const newSs = SpreadsheetApp.create(fileName);
  const newFile = DriveApp.getFileById(newSs.getId());

  // Pindahkan ke folder output
  const targetFolder = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  newFile.moveTo(targetFolder);

  // Tulis data
  const sheet = newSs.getSheets()[0];
  sheet.setName('Crosscheck List');
  const out = [['No.', 'Main Dealer', 'Status AHASS', 'No. Rangka', 'No. Claim']];
  picked.forEach((p, i) => out.push([i + 1, p.md, p.status, p.rangka, p.claim]));
  sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
  sheet.getRange(1, 1, 1, out[0].length)
       .setFontWeight('bold')
       .setBackground('#1F4E78')
       .setFontColor('#FFFFFF');
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 260);
  sheet.setFrozenRows(1);

  Logger.log('File dibuat: ' + fileName + ' (rows=' + picked.length + ') → ' + newFile.getUrl());
  Logger.log('=== runCrosscheckCollector DONE (elapsed ' + ((new Date() - startTs)/1000) + 's) ===');

  return { fileId: newSs.getId(), fileName: fileName, count: picked.length, url: newFile.getUrl() };
}

/**
 * Cari nomor urut berikutnya untuk file crosscheck point N.
 * Scan folder output, ambil angka terbesar dari nama file yang cocok, +1.
 * Jika belum ada file sama sekali → pakai CROSSCHECK_START_NUMBER.
 */
function getNextCrosscheckNumber_() {
  const cfg = CROSSCHECK_CONFIG;
  const folder = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  const it = folder.getFiles();
  const prefixLower = cfg.CROSSCHECK_FILE_PREFIX.toLowerCase();
  let maxN = cfg.CROSSCHECK_START_NUMBER - 1;
  while (it.hasNext()) {
    const f = it.next();
    const name = String(f.getName() || '').trim();
    const lower = name.toLowerCase();
    if (lower.indexOf(prefixLower) !== 0) continue;
    const rest = name.substring(cfg.CROSSCHECK_FILE_PREFIX.length).trim();
    const m = rest.match(/^(\d+)/);
    if (!m) continue;
    const n = parseInt(m[1], 10);
    if (!isNaN(n) && n > maxN) maxN = n;
  }
  return maxN + 1;
}

/**
 * Baca semua file crosscheck point sebelumnya, kumpulkan set No. Rangka yang sudah dipakai.
 */
function loadPreviousCrosscheckRangkas_() {
  const cfg = CROSSCHECK_CONFIG;
  const set = new Set();
  const folder = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  const it = folder.getFiles();
  const prefixLower = cfg.CROSSCHECK_FILE_PREFIX.toLowerCase();
  while (it.hasNext()) {
    const f = it.next();
    const name = String(f.getName() || '').trim().toLowerCase();
    if (name.indexOf(prefixLower) !== 0) continue;
    try {
      const ss = SpreadsheetApp.openById(f.getId());
      const sheet = ss.getSheets()[0];
      const vals = sheet.getDataRange().getValues();
      if (vals.length < 2) continue;
      const header = vals[0].map(v => String(v || '').trim());
      let idxR = findColumnIndex_(header, 'No. Rangka');
      if (idxR < 0) idxR = 3; // fallback D
      for (let i = 1; i < vals.length; i++) {
        const r = normUpper_(vals[i][idxR]);
        if (r) set.add(r);
      }
    } catch (e) {
      Logger.log('Gagal baca ' + f.getName() + ': ' + e.message);
    }
  }
  return set;
}

/**
 * Fisher-Yates shuffle in-place.
 */
function shuffleInPlace_(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
  }
  return arr;
}
