/**
 * crosscheckComparison.gs
 * PROSES 2: Komparasi NKH vs GKA berdasarkan crosscheck point terbaru.
 *
 * Jalankan: runCrosscheckComparison()   (manual)
 * Trigger : dipasang otomatis oleh installTriggersCrosscheck() → tiap Jumat 15:30.
 *
 * Alur:
 *   1. Cari file "crosscheck point N" terbaru di folder output.
 *   2. Ambil list No. Rangka dari file tsb.
 *   3. Lookup rangka tsb di NKH (Form Responses 1) → dapatkan baris NKH-nya saja.
 *   4. Bandingkan dengan GKA (Form Responses 1) menggunakan No. Rangka.
 *   5. Bangun workbook output (Summary Akurasi + Detail AMORE) di folder output.
 *   6. Kirim email berisi link file hasil komparasi.
 */

function runCrosscheckComparison() {
  const cfg = CROSSCHECK_CONFIG;
  const startTs = new Date();
  Logger.log('=== runCrosscheckComparison START @ ' + startTs);

  // --- Step 1: temukan file crosscheck terbaru ---
  const latest = findLatestCrosscheckFile_();
  if (!latest) throw new Error('Tidak ada file crosscheck point ditemukan di folder output.');
  Logger.log('Latest crosscheck: ' + latest.name + ' (n=' + latest.number + ')');

  // Ambil rangka dari file tsb
  const ccSs = SpreadsheetApp.openById(latest.fileId);
  const ccVals = ccSs.getSheets()[0].getDataRange().getValues();
  if (ccVals.length < 2) throw new Error('File ' + latest.name + ' tidak memiliki data.');
  const ccHeader = ccVals[0].map(v => String(v || '').trim());
  let idxCcRangka = findColumnIndex_(ccHeader, 'No. Rangka');
  let idxCcClaim  = findColumnIndex_(ccHeader, 'No. Claim');
  if (idxCcRangka < 0) idxCcRangka = 3;
  if (idxCcClaim < 0)  idxCcClaim  = 4;

  const rangkaList = [];
  const claimByRangka = {};
  for (let i = 1; i < ccVals.length; i++) {
    const r = normUpper_(ccVals[i][idxCcRangka]);
    const c = normTrim_(ccVals[i][idxCcClaim]);
    if (!r) continue;
    rangkaList.push(r);
    claimByRangka[r] = c;
  }
  Logger.log('Rangka dari crosscheck: ' + rangkaList.length);

  // --- Step 2: baca NKH & GKA response sheets ---
  const nkhRows = readResponsesById_(cfg.NKH_SPREADSHEET_ID);
  const gkaRows = readResponsesById_(cfg.GKA_SPREADSHEET_ID);
  Logger.log('NKH rows=' + nkhRows.rows.length + ', GKA rows=' + gkaRows.rows.length);

  // Index by rangka
  const nkhByRangka = indexByRangka_(nkhRows);
  const gkaByRangka = indexByRangka_(gkaRows);

  // --- Step 3: filter NKH sesuai crosscheck (lookup) ---
  const targetRangkas = rangkaList.filter(r => nkhByRangka[r] != null);
  Logger.log('Rangka yang ada di NKH: ' + targetRangkas.length + ' / ' + rangkaList.length);

  // --- Step 4: bangun detail records untuk komparasi ---
  const detail = [];
  targetRangkas.forEach(rangka => {
    const nkhRow = nkhByRangka[rangka];
    const gkaRow = gkaByRangka[rangka] || null;
    const claim = claimByRangka[rangka] || getFieldFromRow_(nkhRow, nkhRows.headerIndex, 'No. Claim');

    const rec = { rangka: rangka, claim: claim, cells: {} };
    cfg.COMPARE_COLUMNS.forEach(colDef => {
      const nv = getFieldFromRow_(nkhRow, nkhRows.headerIndex, colDef.header);
      const gv = gkaRow ? getFieldFromRow_(gkaRow, gkaRows.headerIndex, colDef.header) : '';
      const s = buildComparisonString_(nv, gv);
      if (s !== null) rec.cells[colDef.letter] = s;
    });
    detail.push(rec);
  });

  if (detail.length === 0) {
    Logger.log('Tidak ada rangka NKH yang bisa dibandingkan. Skip pembuatan output.');
    return null;
  }

  // --- Step 5: hitung ringkasan ---
  const stats = computeStats_(detail);

  // --- Step 6: buat spreadsheet output ---
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const outName = cfg.COMPARISON_FILE_PREFIX + latest.name.replace(cfg.CROSSCHECK_FILE_PREFIX, '') + ' (' + ts + ')';
  const outSs = SpreadsheetApp.create(outName);
  const outFile = DriveApp.getFileById(outSs.getId());
  DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID).addFile(outFile);
  DriveApp.getRootFolder().removeFile(outFile);

  writeSummarySheet_(outSs, stats, detail.length);
  writeDetailSheet_(outSs, detail);

  const url = outFile.getUrl();
  Logger.log('Output komparasi: ' + outName + ' → ' + url);

  // --- Step 7: kirim email ---
  sendComparisonEmail_(outName, url, stats, latest);

  Logger.log('=== runCrosscheckComparison DONE (elapsed ' + ((new Date() - startTs)/1000) + 's) ===');
  return { fileId: outSs.getId(), fileName: outName, url: url, stats: stats };
}

// ---------- Support helpers ----------

function findLatestCrosscheckFile_() {
  const cfg = CROSSCHECK_CONFIG;
  const folder = DriveApp.getFolderById(cfg.OUTPUT_FOLDER_ID);
  const it = folder.getFiles();
  const prefixLower = cfg.CROSSCHECK_FILE_PREFIX.toLowerCase();
  let best = null;
  while (it.hasNext()) {
    const f = it.next();
    const name = String(f.getName() || '').trim();
    const lower = name.toLowerCase();
    if (lower.indexOf(prefixLower) !== 0) continue;
    const rest = name.substring(cfg.CROSSCHECK_FILE_PREFIX.length).trim();
    const m = rest.match(/^(\d+)/);
    if (!m) continue;
    const n = parseInt(m[1], 10);
    if (isNaN(n)) continue;
    if (!best || n > best.number) best = { fileId: f.getId(), name: name, number: n };
  }
  return best;
}

function readResponsesById_(id) {
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(CROSSCHECK_CONFIG.RESPONSE_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + CROSSCHECK_CONFIG.RESPONSE_SHEET_NAME + '" tidak ditemukan di ' + id);
  const vals = sheet.getDataRange().getValues();
  const header = (vals[0] || []).map(v => String(v || '').trim());
  const headerIndex = {};
  header.forEach((h, i) => { if (h) headerIndex[h.toLowerCase()] = i; });
  return { rows: vals.slice(1), header: header, headerIndex: headerIndex };
}

function indexByRangka_(data) {
  const map = {};
  const idxR = data.headerIndex['no. rangka'];
  if (idxR == null) throw new Error('Kolom "No. Rangka" tidak ada di response sheet.');
  data.rows.forEach(row => {
    const r = normUpper_(row[idxR]);
    if (r && !map[r]) map[r] = row;
  });
  return map;
}

function getFieldFromRow_(row, headerIndex, targetHeader) {
  if (!row || !headerIndex) return '';
  const key = String(targetHeader || '').trim().toLowerCase();
  // exact match
  if (headerIndex[key] != null) return row[headerIndex[key]];
  // partial match (header sumber sering ada tambahan " (contoh: ...)")
  for (const k in headerIndex) {
    if (k.indexOf(key) === 0) return row[headerIndex[k]];
  }
  return '';
}

function buildComparisonString_(nv, gv) {
  const a = normalizeCell_(nv);
  const b = normalizeCell_(gv);
  if (a === '' && b === '') return null;
  if (a === b) return 'Sesuai (' + a + ')';
  return 'Tidak Sesuai (NKH : ' + (a || '-') + ' | GKA : ' + (b || '-') + ')';
}

function normalizeCell_(v) {
  if (v == null) return '';
  return String(v).replace(/[\r\n]+/g, ' ').trim();
}

function computeStats_(detail) {
  const counts = { M1:0, NM1:0, NM2:0, M2_ADM:0, M2_ANO:0, M2_LVL:0, NM_ETC:0 };
  const m2Match = { M2_ADM:0, M2_ANO:0, M2_LVL:0 };
  const m2NotMatch = { M2_ADM:0, M2_ANO:0, M2_LVL:0 };
  const m1Breakdown = {
    'M1.1':[0,0,0,0,0], 'NM1.1':[0,0,0,0,0], 'NM2.1':[0,0,0,0,0],
    'M2.1':[0,0,0,0,0], 'NM_ETC':[0,0,0,0,0]
  };
  const m1SubCols = ['E','F','G','H','I'];

  function catOfD(d) {
    if (d == null) return null;
    const s = String(d).trim();
    if (s === 'Sesuai (Penilaian Video Treatment LCR)') return 'M1';
    if (s.indexOf('Tidak Sesuai (NKH : Penilaian Video Treatment LCR') === 0) return 'NM1';
    if (s.indexOf('Tidak Sesuai') === 0 && s.indexOf('GKA : Penilaian Video Treatment LCR') >= 0) return 'NM2';
    if (s.indexOf('Sesuai') === 0 && s.indexOf('Kesalahan Administrasi') >= 0) return 'M2_ADM';
    if (s.indexOf('Sesuai') === 0 && s.indexOf('Anomali Video') >= 0) return 'M2_ANO';
    if (s.indexOf('Sesuai') === 0 && s.indexOf('Salah Level Treatment') >= 0) return 'M2_LVL';
    if (s.indexOf('Tidak Sesuai') === 0) return 'NM_ETC';
    return null;
  }
  function isSesuai(v) {
    if (v == null) return null;
    const s = String(v).trim().toLowerCase();
    if (s === '') return null;
    if (s.indexOf('tidak sesuai') === 0) return false;
    if (s.indexOf('sesuai') === 0) return true;
    return null;
  }

  detail.forEach(rec => {
    const c = catOfD(rec.cells['D']);
    if (!c) return;
    counts[c] = (counts[c] || 0) + 1;

    if (c === 'M2_ADM') {
      const vs = ['K','L','M','N'].map(l => isSesuai(rec.cells[l]));
      const nn = vs.filter(v => v !== null);
      if (nn.length > 0 && nn.every(v => v === true)) m2Match.M2_ADM++;
      else m2NotMatch.M2_ADM++;
    } else if (c === 'M2_ANO') {
      if (isSesuai(rec.cells['O']) === true) m2Match.M2_ANO++;
      else m2NotMatch.M2_ANO++;
    } else if (c === 'M2_LVL') {
      if (isSesuai(rec.cells['P']) === true) m2Match.M2_LVL++;
      else m2NotMatch.M2_LVL++;
    } else if (c === 'M1') {
      m1SubCols.forEach((letter, j) => {
        const v = rec.cells[letter];
        if (v == null) return;
        const s = String(v).trim();
        if (s === 'Sesuai (1 - ADA, JELAS)') m1Breakdown['M1.1'][j]++;
        else if (s.indexOf('Tidak Sesuai (NKH : 1 - ADA, JELAS') === 0) m1Breakdown['NM1.1'][j]++;
        else if (s.indexOf('Tidak Sesuai') === 0 && s.indexOf('GKA : 1 - ADA, JELAS') >= 0) m1Breakdown['NM2.1'][j]++;
        else if (s.indexOf('Sesuai') === 0) m1Breakdown['M2.1'][j]++;
        else if (s.indexOf('Tidak Sesuai') === 0) m1Breakdown['NM_ETC'][j]++;
      });
    }
  });

  const GT = Object.keys(counts).reduce((a,k) => a + counts[k], 0);
  const totalMatch = counts.M1 + m2Match.M2_ADM + m2Match.M2_ANO + m2Match.M2_LVL;
  const totalNotMatch = counts.NM1 + counts.NM2 + counts.NM_ETC +
                        m2NotMatch.M2_ADM + m2NotMatch.M2_ANO + m2NotMatch.M2_LVL;

  return {
    counts: counts,
    m2Match: m2Match,
    m2NotMatch: m2NotMatch,
    m1Breakdown: m1Breakdown,
    GT: GT,
    totalMatch: totalMatch,
    totalNotMatch: totalNotMatch,
    accuracy: GT > 0 ? totalMatch / GT : 0
  };
}

function writeSummarySheet_(ss, stats, dataCount) {
  const sh = ss.getSheets()[0];
  sh.setName('Summary Akurasi');
  sh.clear();
  const GT = stats.GT;
  const rows = [];
  rows.push(['SUMMARY AKURASI PENGECEKAN NKH vs GKA']);
  rows.push(['Sumber: Lookup NKH → crosscheck point, lalu dibandingkan dengan GKA. Total data dianalisis: ' + dataCount]);
  rows.push([]);
  rows.push(['KATEGORI','SUB-KATEGORI / DETAIL','STATUS','JUMLAH','% dari Total']);
  const pct = (x) => GT > 0 ? (x / GT * 100).toFixed(2) + '%' : '0.00%';
  rows.push(['Match 1', 'Sesuai (Penilaian Video Treatment LCR)', 'Match', stats.counts.M1, pct(stats.counts.M1)]);
  rows.push(['Not Match 1', 'Tidak Sesuai (NKH: LCR | GKA: Non-LCR)', 'Not Match', stats.counts.NM1, pct(stats.counts.NM1)]);
  rows.push(['Not Match 2', 'Tidak Sesuai (NKH: Non-LCR | GKA: LCR)', 'Not Match', stats.counts.NM2, pct(stats.counts.NM2)]);
  rows.push(['Match 2 - Kesalahan Administrasi', 'Sesuai (Kesalahan Administrasi)', 'Total', stats.counts.M2_ADM, pct(stats.counts.M2_ADM)]);
  rows.push(['', 'Cek Kolom K-N', 'Match', stats.m2Match.M2_ADM, pct(stats.m2Match.M2_ADM)]);
  rows.push(['', 'Cek Kolom K-N', 'Not Match', stats.m2NotMatch.M2_ADM, pct(stats.m2NotMatch.M2_ADM)]);
  rows.push(['Match 2 - Anomali Video', 'Sesuai (Anomali Video)', 'Total', stats.counts.M2_ANO, pct(stats.counts.M2_ANO)]);
  rows.push(['', 'Cek Kolom O', 'Match', stats.m2Match.M2_ANO, pct(stats.m2Match.M2_ANO)]);
  rows.push(['', 'Cek Kolom O', 'Not Match', stats.m2NotMatch.M2_ANO, pct(stats.m2NotMatch.M2_ANO)]);
  rows.push(['Match 2 - Salah Level Treatment', 'Sesuai (Salah Level Treatment)', 'Total', stats.counts.M2_LVL, pct(stats.counts.M2_LVL)]);
  rows.push(['', 'Cek Kolom P', 'Match', stats.m2Match.M2_LVL, pct(stats.m2Match.M2_LVL)]);
  rows.push(['', 'Cek Kolom P', 'Not Match', stats.m2NotMatch.M2_LVL, pct(stats.m2NotMatch.M2_LVL)]);
  rows.push(['Not Match Lainnya', 'Tidak Sesuai (keduanya Non-LCR / tidak ada pasangan)', 'Not Match', stats.counts.NM_ETC, pct(stats.counts.NM_ETC)]);
  rows.push(['GRAND TOTAL DATA', '', '', GT, '100.00%']);
  rows.push([]);
  rows.push(['RINGKASAN AKURASI (Match vs Not Match)']);
  rows.push(['STATUS','DEFINISI','','JUMLAH','% dari Total']);
  const pctG = (x) => GT > 0 ? (x / GT * 100).toFixed(2) + '%' : '0.00%';
  rows.push(['MATCH (Sesuai)', 'PIC memilih kategori sama DAN detail sama', '', stats.totalMatch, pctG(stats.totalMatch)]);
  rows.push(['NOT MATCH (Tidak Sesuai)', 'PIC berbeda kategori ATAU detail berbeda', '', stats.totalNotMatch, pctG(stats.totalNotMatch)]);
  rows.push(['TOTAL', '', '', GT, '100.00%']);
  rows.push([]);
  rows.push(['AKURASI PENGECEKAN NKH terhadap GKA', '', '', '', (stats.accuracy * 100).toFixed(2) + '%']);
  rows.push([]);
  rows.push(['BREAKDOWN MATCH 1 (per Sub-Kategori Penilaian Video Treatment LCR)']);
  rows.push(['SUB-KATEGORI','Flintkote (E)','Regulator (F)','Timer (G)','Marking Selang (H)','Marking Frame (I)','TOTAL','%']);
  const m1 = stats.m1Breakdown;
  const grandM1 = Object.keys(m1).reduce((a,k) => a + m1[k].reduce((x,y)=>x+y,0), 0);
  const lbl = { 'M1.1':'Match 1.1', 'NM1.1':'Not Match 1.1', 'NM2.1':'Not Match 2.1', 'M2.1':'Match 2.1', 'NM_ETC':'Not Match Lainnya' };
  Object.keys(m1).forEach(k => {
    const arr = m1[k];
    const tot = arr.reduce((x,y)=>x+y, 0);
    const p = grandM1 > 0 ? (tot / grandM1 * 100).toFixed(2) + '%' : '0.00%';
    rows.push([lbl[k], arr[0], arr[1], arr[2], arr[3], arr[4], tot, p]);
  });
  const colTotals = [0,0,0,0,0];
  Object.keys(m1).forEach(k => m1[k].forEach((v,i) => colTotals[i]+=v));
  rows.push(['TOTAL'].concat(colTotals).concat([grandM1, '100.00%']));

  // Normalize row widths
  const maxCols = rows.reduce((m,r) => Math.max(m, r.length), 0);
  const norm = rows.map(r => {
    const c = r.slice(); while (c.length < maxCols) c.push(''); return c;
  });
  sh.getRange(1, 1, norm.length, maxCols).setValues(norm);

  // Basic styling
  sh.getRange(1, 1, 1, maxCols).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  sh.setColumnWidth(1, 32*8);
  sh.setColumnWidth(2, 45*8);
  sh.setFrozenRows(1);
}

function writeDetailSheet_(ss, detail) {
  const cfg = CROSSCHECK_CONFIG;
  const sh = ss.insertSheet('Detail AMORE');
  const header1 = ['No.', 'NO. RANGKA', 'NO. CLAIM'];
  const header2 = ['', '', ''];
  cfg.COMPARE_COLUMNS.forEach(cd => {
    header1.push(cd.header);
    header2.push(cd.subLabel);
  });
  sh.appendRow(header1);

  const rows = [];
  detail.forEach((rec, i) => {
    const arr = [i+1, rec.rangka, rec.claim || ''];
    cfg.COMPARE_COLUMNS.forEach(cd => {
      arr.push(rec.cells[cd.letter] || '');
    });
    rows.push(arr);
  });
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  sh.getRange(1, 1, 1, header1.length)
    .setFontWeight('bold')
    .setBackground('#2E75B6')
    .setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 50);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 220);
  for (let c = 4; c <= header1.length; c++) sh.setColumnWidth(c, 240);
}

function sendComparisonEmail_(fileName, fileUrl, stats, crosscheckMeta) {
  const cfg = CROSSCHECK_CONFIG;
  const to = String(cfg.EMAIL_TO || '').trim();
  if (!to) {
    Logger.log('EMAIL_TO kosong, email tidak dikirim.');
    return;
  }
  const cc = String(cfg.EMAIL_CC || '').trim();
  const acc = (stats.accuracy * 100).toFixed(2) + '%';
  const subject = 'Hasil Komparasi NKH vs GKA - ' + fileName;
  const body =
`Dear PIC,

Berikut hasil komparasi otomatis PIC NKH vs PIC GKA untuk ${crosscheckMeta.name}.

Ringkasan:
- Total data dianalisis : ${stats.GT}
- MATCH (Sesuai)        : ${stats.totalMatch}
- NOT MATCH (Tidak)     : ${stats.totalNotMatch}
- Akurasi NKH vs GKA    : ${acc}

Detail lengkap dapat diakses melalui link berikut:
${fileUrl}

File terdiri dari 2 sheet:
1. Summary Akurasi  - ringkasan akurasi & breakdown Match 1 (kolom E-I)
2. Detail AMORE     - hasil perbandingan per baris (kolom D-P)

Email ini dikirim otomatis oleh sistem crosscheck AMORE.
Terima kasih.`;

  try {
    const opts = { name: cfg.EMAIL_SENDER_NAME };
    if (cc) opts.cc = cc;
    GmailApp.sendEmail(to, subject, body, opts);
    Logger.log('Email terkirim ke ' + to + (cc ? ' (cc: ' + cc + ')' : ''));
  } catch (e) {
    Logger.log('Gagal kirim email: ' + e.message);
  }
}
