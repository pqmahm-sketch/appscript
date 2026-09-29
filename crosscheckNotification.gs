/**
 * crosscheckNotification.gs
 * Notifikasi email saat file crosscheck point atau komparasi sudah tersedia.
 *
 * Dipanggil dari:
 *   - runCrosscheckCollector()  → notifyCrosscheckReady_()
 *   - runCrosscheckComparison() → notifyComparisonReady_()
 *
 * Konfigurasi email (TO, CC, SUBJECT, SENDER_NAME, BODY) ada di crosscheckConfig.gs.
 * Isi dulu sebelum deploy — fungsi akan skip jika TO masih kosong.
 *
 * Placeholder pada BODY:
 *   Crosscheck : {fileName}, {rowCount}, {fileUrl}
 *   Komparasi  : {fileName}, {fileUrl}, {summary}
 */

/**
 * Kirim notifikasi bahwa file crosscheck point sudah tersedia.
 * @param {string} fileName  - Nama file, misal "crosscheck point 4"
 * @param {string} fileUrl   - URL Google Sheets
 * @param {number} rowCount  - Jumlah baris data di file
 */
function notifyCrosscheckReady_(fileName, fileUrl, rowCount) {
  const cfg = CROSSCHECK_CONFIG;
  const to = (cfg.NOTIF_CROSSCHECK_TO || '').trim();
  if (!to) {
    Logger.log('[Notif] NOTIF_CROSSCHECK_TO kosong, skip email crosscheck.');
    return;
  }

  const subject = cfg.NOTIF_CROSSCHECK_SUBJECT || ('File Crosscheck Point Tersedia: ' + fileName);

  const body = (cfg.NOTIF_CROSSCHECK_BODY || '')
    .replace(/\{fileName\}/g, fileName)
    .replace(/\{rowCount\}/g, String(rowCount))
    .replace(/\{fileUrl\}/g, fileUrl);

  const options = { name: cfg.NOTIF_CROSSCHECK_SENDER_NAME || 'PQM AHM' };
  const cc = (cfg.NOTIF_CROSSCHECK_CC || '').trim();
  if (cc) options.cc = cc;

  try {
    GmailApp.sendEmail(to, subject, body, options);
    Logger.log('[Notif] Email crosscheck terkirim → ' + to);
  } catch (e) {
    Logger.log('[Notif] Gagal kirim email crosscheck: ' + e.message);
  }
}

/**
 * Kirim notifikasi bahwa file komparasi penilaian sudah tersedia.
 * @param {string} fileName  - Nama file komparasi
 * @param {string} fileUrl   - URL Google Sheets
 * @param {Object} stats     - Objek statistik dari computeStats_ (opsional)
 */
function notifyComparisonReady_(fileName, fileUrl, stats) {
  const cfg = CROSSCHECK_CONFIG;
  const to = (cfg.NOTIF_COMPARISON_TO || '').trim();
  if (!to) {
    Logger.log('[Notif] NOTIF_COMPARISON_TO kosong, skip email komparasi.');
    return;
  }

  const subject = cfg.NOTIF_COMPARISON_SUBJECT || ('Hasil Komparasi Penilaian AMORE Tersedia: ' + fileName);

  let summary = '';
  if (stats) {
    const total = (stats.m1 || 0) + (stats.nm1 || 0) + (stats.nm2 || 0) +
                  (stats.m2_adm || 0) + (stats.m2_ano || 0) + (stats.m2_lvl || 0) + (stats.nm_etc || 0);
    const match = (stats.m1 || 0) + (stats.m2_adm || 0) + (stats.m2_ano || 0) + (stats.m2_lvl || 0);
    const pct = total > 0 ? ((match / total) * 100).toFixed(1) : '0.0';
    summary =
      '\nRingkasan:\n' +
      '  Total data     : ' + total + '\n' +
      '  Sesuai (Match) : ' + match + ' (' + pct + '%)\n' +
      '  Tidak Sesuai   : ' + (total - match) + ' (' + (100 - parseFloat(pct)).toFixed(1) + '%)\n';
  }

  const body = (cfg.NOTIF_COMPARISON_BODY || '')
    .replace(/\{fileName\}/g, fileName)
    .replace(/\{fileUrl\}/g, fileUrl)
    .replace(/\{summary\}/g, summary);

  const options = { name: cfg.NOTIF_COMPARISON_SENDER_NAME || 'PQM AHM' };
  const cc = (cfg.NOTIF_COMPARISON_CC || '').trim();
  if (cc) options.cc = cc;

  try {
    GmailApp.sendEmail(to, subject, body, options);
    Logger.log('[Notif] Email komparasi terkirim → ' + to);
  } catch (e) {
    Logger.log('[Notif] Gagal kirim email komparasi: ' + e.message);
  }
}
