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

function syncToRekapLookerStudio() {
  var responsesSS = SpreadsheetApp.openById(RESPONSES_SPREADSHEET_ID);
  var formSheet = responsesSS.getSheetByName("Form Responses 1");

  var lastRow = formSheet.getLastRow();
  if (lastRow < 2) return;

  var formData = formSheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // Gunakan data terbaru per AHASS (jika ada duplikat, ambil yang terakhir)
  var latestByAhass = {};
  for (var i = 0; i < formData.length; i++) {
    var noAhass = String(formData[i][1]).trim();
    var hasil = String(formData[i][3]).trim();
    if (!noAhass) continue;
    latestByAhass[noAhass] = {
      mainDealer: String(formData[i][2]).trim(),
      hasil: hasil,
      kodeMD: String(formData[i][2]).split(" - ")[0].trim()
    };
  }

  // Hitung OK dan NG per Kode MD
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

  for (var r = 1; r < rekapData.length; r++) {
    var kodeMD = String(rekapData[r][0]).trim();
    var terdaftar = Number(rekapData[r][2]) || 0;

    // Skip baris region dan nasional (akan dihitung dari agregasi)
    if (kodeMD === "R1" || kodeMD === "R2" || kodeMD === "R3" || kodeMD === "N45") continue;

    var ok = countOK[kodeMD] || 0;
    var ng = countNG[kodeMD] || 0;
    var belumKirim = terdaftar - ok - ng;
    if (belumKirim < 0) belumKirim = 0;

    rekapSheet.getRange(r + 1, 4).setValue(ok);
    rekapSheet.getRange(r + 1, 5).setValue(ng);
    rekapSheet.getRange(r + 1, 6).setValue(belumKirim);

    // Akumulasi untuk region
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

  // Update sheet Pivot untuk Pie Chart
  updatePivotSheet(rekapSS, rekapSheet);

  Logger.log("Sync ke Rekap Looker Studio selesai.");
}

function updatePivotSheet(rekapSS, rekapSheet) {
  var pivotSheet = rekapSS.getSheetByName("Pivot untuk Pie Chart");
  if (!pivotSheet) return;

  var rekapData = rekapSheet.getDataRange().getValues();
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
