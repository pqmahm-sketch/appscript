function createDownloadLinksTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "generateDownloadLinks") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("generateDownloadLinks")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log("Trigger download links berhasil dipasang! Akan berjalan setiap hari jam 08:00.");
}
