var RESPONSES_SPREADSHEET_ID = "1zyRDiYulkXlEvfvJ1VTFimdk4NSDBzQ_ZUzb9w4KB9M";
var REKAP_SPREADSHEET_ID = "1Pe1q5NGG1xCMFNvoAesv2V-66KStm0wsf385CilEbi8";

var REGION_MAP = {
  "B10": "R1", "B3Z": "R1", "C10": "R1", "C3Z": "R1",
  "D2Z": "R1", "D3Z": "R1", "E20": "R1", "G01": "R1",
  "G02": "R1", "G5Z": "R1", "H2Z": "R1",
  "I01": "R2", "I3Z": "R2", "J10": "R2", "J20": "R2",
  "K0Z": "R2", "L01": "R2", "M2Z": "R2", "M3Z": "R2",
  "N01": "R2", "N02": "R2",
  "Q01": "R3", "R4Z": "R3", "R5Z": "R3", "T10": "R3",
  "U10": "R3", "V2Z": "R3", "W01": "R3", "Z11": "R3"
};

var REKAP_FOLDER_ID = "0AJuZm4K3MzzeUk9PVA";

function syncToRekapLookerStudio() {
  var responsesSS = SpreadsheetApp.openById(RESPONSES_SPREADSHEET_ID);
  var formSheet = responsesSS.getSheetByName("Form Responses 1");

  var lastRow = formSheet.getLastRow();
  if (lastRow < 2) return;

  var formData = formSheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // Kumpulkan nama lengkap MD dan data terbaru per AHASS
  var mdFullNames = {};
  var latestByAhass = {};
  var allEntriesByMD = {};

  for (var i = 0; i < formData.length; i++) {
    var noAhass = String(formData[i][1]).trim();
    var mainDealerFull = String(formData[i][2]).trim();
    var hasil = String(formData[i][3]).trim();
    var alasanNG = String(formData[i][4]).trim();
    var linkFoto = String(formData[i][8] || "").trim();
    var kodeMD = mainDealerFull.split(" - ")[0].trim();

    if (!noAhass) continue;

    mdFullNames[kodeMD] = mainDealerFull;

    latestByAhass[noAhass] = {
      mainDealer: mainDealerFull,
      hasil: hasil,
      kodeMD: kodeMD
    };

    if (!allEntriesByMD[kodeMD]) allEntriesByMD[kodeMD] = [];
    allEntriesByMD[kodeMD].push({
      noAhass: noAhass,
      mainDealer: mainDealerFull,
      hasil: hasil,
      alasanNG: alasanNG,
      linkFoto: linkFoto
    });
  }

  // Hitung OK dan NG per Kode MD (berdasarkan data terbaru per AHASS)
  var countOK = {};
  var countNG = {};
  var ahassKeys = Object.keys(latestByAhass);
  for (var j = 0; j < ahassKeys.length; j++) {
    var entry = latestByAhass[ahassKeys[j]];
    var kode = entry.kodeMD;
    if (!countOK[kode]) countOK[kode] = 0;
    if (!countNG[kode]) countNG[kode] = 0;
    if (entry.hasil === "OK") {
      countOK[kode]++;
    } else if (entry.hasil === "NG") {
      countNG[kode]++;
    }
  }

  // Update Sheet1 di Rekap
  var rekapSS = SpreadsheetApp.openById(REKAP_SPREADSHEET_ID);
  var rekapSheet = rekapSS.getSheetByName("Sheet1");
  var rekapData = rekapSheet.getDataRange().getValues();

  var regionOK = {};
  var regionNG = {};
  var regionTerdaftar = {};
  var mdWithData = [];

  for (var r = 1; r < rekapData.length; r++) {
    var kodeMD = String(rekapData[r][0]).trim();
    var terdaftar = Number(rekapData[r][2]) || 0;

    if (kodeMD === "R1" || kodeMD === "R2" || kodeMD === "R3" || kodeMD === "N45") continue;

    // Update Nama Main Dealer (kolom B) dengan nama lengkap dari Responses
    if (mdFullNames[kodeMD]) {
      rekapSheet.getRange(r + 1, 2).setValue(mdFullNames[kodeMD]);
    }

    var ok = countOK[kodeMD] || 0;
    var ng = countNG[kodeMD] || 0;
    var belumKirim = terdaftar - ok - ng;
    if (belumKirim < 0) belumKirim = 0;

    rekapSheet.getRange(r + 1, 4).setValue(ok);
    rekapSheet.getRange(r + 1, 5).setValue(ng);
    rekapSheet.getRange(r + 1, 6).setValue(belumKirim);

    if (ok > 0 || ng > 0) {
      mdWithData.push(kodeMD);
    }

    var region = REGION_MAP[kodeMD];
    if (region) {
      if (!regionOK[region]) regionOK[region] = 0;
      if (!regionNG[region]) regionNG[region] = 0;
      if (!regionTerdaftar[region]) regionTerdaftar[region] = 0;
      regionOK[region] += ok;
      regionNG[region] += ng;
      regionTerdaftar[region] += terdaftar;
    }
  }

  // Update baris Region dan Nasional
  var nasionalOK = 0, nasionalNG = 0, nasionalTerdaftar = 0;
  for (var r2 = 1; r2 < rekapData.length; r2++) {
    var kode2 = String(rekapData[r2][0]).trim();
    if (kode2 === "R1" || kode2 === "R2" || kode2 === "R3") {
      var rOK = regionOK[kode2] || 0;
      var rNG = regionNG[kode2] || 0;
      var rTerdaftar = regionTerdaftar[kode2] || 0;
      var rBelum = rTerdaftar - rOK - rNG;
      if (rBelum < 0) rBelum = 0;
      rekapSheet.getRange(r2 + 1, 4).setValue(rOK);
      rekapSheet.getRange(r2 + 1, 5).setValue(rNG);
      rekapSheet.getRange(r2 + 1, 6).setValue(rBelum);
      nasionalOK += rOK;
      nasionalNG += rNG;
      nasionalTerdaftar += rTerdaftar;
    }
    if (kode2 === "N45") {
      var nBelum = nasionalTerdaftar - nasionalOK - nasionalNG;
      if (nBelum < 0) nBelum = 0;
      rekapSheet.getRange(r2 + 1, 4).setValue(nasionalOK);
      rekapSheet.getRange(r2 + 1, 5).setValue(nasionalNG);
      rekapSheet.getRange(r2 + 1, 6).setValue(nBelum);
    }
  }

  // Update Pivot untuk Pie Chart (baca ulang setelah update nama)
  var updatedRekapData = rekapSheet.getDataRange().getValues();
  updatePivotSheet(rekapSS, updatedRekapData);

  // Update Download Links untuk semua MD yang punya data
  updateDownloadLinks(rekapSS, allEntriesByMD, mdWithData, mdFullNames);

  Logger.log("Sync ke Rekap Looker Studio selesai.");
}

function updatePivotSheet(rekapSS, rekapData) {
  var pivotSheet = rekapSS.getSheetByName("Pivot untuk Pie Chart");
  if (!pivotSheet) return;

  var pivotData = [["Nama Main Dealer", "Status", "Jumlah"]];

  for (var i = 1; i < rekapData.length; i++) {
    var namaMD = String(rekapData[i][1]).trim();
    var ok = Number(rekapData[i][3]) || 0;
    var ng = Number(rekapData[i][4]) || 0;
    var belum = Number(rekapData[i][5]) || 0;

    pivotData.push([namaMD, "OK", ok]);
    pivotData.push([namaMD, "NG", ng]);
    pivotData.push([namaMD, "Belum Kirim", belum]);
  }

  pivotSheet.clearContents();
  pivotSheet.getRange(1, 1, pivotData.length, 3).setValues(pivotData);
}

function updateDownloadLinks(rekapSS, allEntriesByMD, mdWithData, mdFullNames) {
  var dlSheet = rekapSS.getSheetByName("Download Links");
  if (!dlSheet) return;

  var existingData = dlSheet.getDataRange().getValues();
  var existingLinks = {};
  for (var e = 1; e < existingData.length; e++) {
    var existKode = String(existingData[e][0]).trim();
    var existLink = String(existingData[e][2]).trim();
    if (existKode && existLink) {
      // Ekstrak spreadsheet ID dari URL
      var match = existLink.match(/\/d\/([a-zA-Z0-9_-]+)\//);
      if (match) {
        existingLinks[existKode] = match[1];
      }
    }
  }

  var dlRows = [["Kode MD", "Nama Main Dealer", "Download Link"]];

  for (var m = 0; m < mdWithData.length; m++) {
    var kodeMD = mdWithData[m];
    var namaMD = mdFullNames[kodeMD] || kodeMD;
    var entries = allEntriesByMD[kodeMD] || [];

    var ssId;
    if (existingLinks[kodeMD]) {
      // Gunakan spreadsheet yang sudah ada, update isinya
      ssId = existingLinks[kodeMD];
      try {
        var existingSS = SpreadsheetApp.openById(ssId);
        var sheet = existingSS.getSheets()[0];
        sheet.clearContents();
        writePerMDData(sheet, entries);
        // Pastikan file tetap bisa diakses semua user
        DriveApp.getFileById(ssId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (err) {
        // Spreadsheet lama tidak bisa diakses, buat baru
        ssId = createPerMDSpreadsheet(kodeMD, namaMD, entries);
      }
    } else {
      ssId = createPerMDSpreadsheet(kodeMD, namaMD, entries);
    }

    var downloadUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
    dlRows.push([kodeMD, namaMD, downloadUrl]);
  }

  dlSheet.clearContents();
  dlSheet.getRange(1, 1, dlRows.length, 3).setValues(dlRows);
}

function createPerMDSpreadsheet(kodeMD, namaMD, entries) {
  var newSS = SpreadsheetApp.create("Data Verifikasi Noxudol - " + kodeMD);
  var sheet = newSS.getSheets()[0];
  writePerMDData(sheet, entries);

  try {
    var file = DriveApp.getFileById(newSS.getId());

    // Share: anyone with the link can view (agar bisa didownload semua user)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Pindahkan ke folder Rekap jika memungkinkan
    var folder = DriveApp.getFolderById(REKAP_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (err) {
    Logger.log("Error setup file: " + err.message);
  }

  return newSS.getId();
}

function writePerMDData(sheet, entries) {
  var data = [["No. AHASS", "Nama Main Dealer", "Hasil Penilaian", "Alasan NG", "Link Foto Bukti"]];
  for (var i = 0; i < entries.length; i++) {
    data.push([
      entries[i].noAhass,
      entries[i].mainDealer,
      entries[i].hasil,
      entries[i].alasanNG,
      entries[i].linkFoto
    ]);
  }
  sheet.getRange(1, 1, data.length, 5).setValues(data);
}

function createSyncTriggerTimeDriven() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === "syncToRekapLookerStudio") {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }

  ScriptApp.newTrigger("syncToRekapLookerStudio")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("Time-driven trigger untuk sync setiap 5 menit berhasil dibuat.");
}
