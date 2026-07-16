function sendEmailNoxudol() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Form Responses 1");
  const emailSheet = ss.getSheetByName("Alamat Email");

  const emailData = emailSheet.getDataRange().getValues();
  const emailMap = {};
  for (let i = 1; i < emailData.length; i++) {
    const mdName = emailData[i][1];
    if (mdName) {
      emailMap[mdName.trim()] = {
        to: (emailData[i][2] || "").trim(),
        cc: (emailData[i][3] || "").trim()
      };
    }
  }

  const lastRow = formSheet.getLastRow();
  if (lastRow < 2) return;

  formSheet.getRange(1, 8).setValue("Status");

  const formData = formSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  let emailCount = 0;

  for (let i = 0; i < formData.length; i++) {
    const rowNum = i + 2;

    const noAHASS    = formData[i][1];
    const mainDealer = formData[i][2];
    const hasil      = formData[i][3];
    const alasanNG   = formData[i][4];
    let   toEmail    = formData[i][5];
    let   ccEmail    = formData[i][6];
    const status     = formData[i][7];

    if (!noAHASS && !mainDealer) continue;
    if (status === "SENT") continue;

    // Hanya kirim email untuk hasil penilaian NG
    if (hasil !== "NG") {
      formSheet.getRange(rowNum, 8).setValue("OK - NO EMAIL NEEDED");
      continue;
    }

    if (!toEmail || !ccEmail) {
      const match = emailMap[mainDealer ? mainDealer.trim() : ""];
      if (match) {
        toEmail = match.to;
        ccEmail = match.cc;
        formSheet.getRange(rowNum, 6).setValue(toEmail);
        formSheet.getRange(rowNum, 7).setValue(ccEmail);
      }
    }

    if (!toEmail) {
      Logger.log(`Baris ${rowNum}: Email tidak ditemukan untuk "${mainDealer}", skip.`);
      formSheet.getRange(rowNum, 8).setValue("EMAIL NOT FOUND");
      continue;
    }

    const subject = "Konfirmasi Penggaris Noxudol NG Kepada MD";
    const body =
`Dear PIC Region,

Berikut informasi yang perlu dikonfirmasi kepada MD :

No. AHASS       : ${noAHASS}
Main Dealer     : ${mainDealer}
Hasil penilaian : ${hasil}
Alasan NG       : ${alasanNG}

Mohon konfirmasi ke MD apakah ini sudah dibuat dengan benar atau belum, lalu minta MD atau AHASS untuk upload ulang foto yang benar pada link yang sudah disebarkan sebelumnya.

Demikian
Terima kasih`;

    try {
      GmailApp.sendEmail(toEmail, subject, body, {
        cc: ccEmail,
        name: "PQM AHM"
      });
      formSheet.getRange(rowNum, 8).setValue("SENT");
      emailCount++;
      Logger.log(`Baris ${rowNum}: SENT → ${toEmail}`);
    } catch (e) {
      formSheet.getRange(rowNum, 8).setValue("ERROR: " + e.message);
      Logger.log(`Baris ${rowNum}: ERROR - ${e.message}`);
    }
  }

  SpreadsheetApp.getUi().alert(`Selesai! ${emailCount} email berhasil dikirim (hanya NG).`);
}

function createTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onFormSubmitHandler")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert("Trigger berhasil dipasang! Email NG & sync Rekap akan otomatis berjalan setiap ada response form baru.");
}

function onFormSubmitHandler(e) {
  sendEmailNoxudol();
  syncToRekapLookerStudio();
}
