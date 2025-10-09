/** Triggers setup
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */

function creaTriggerBenvenuti() {
  ScriptApp.newTrigger("inviaBenvenutiWhatsApp")
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log("✅ Trigger creato: inviaBenvenutiWhatsApp ogni 15 minuti");
}

function createOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();

  // Controlla se il trigger esiste già per evitare duplicati
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onEditInstalled") {
      Logger.log("✅ Trigger 'onEditInstalled' già esistente.");
      return;
    }
  }

  // Se il trigger non esiste, lo crea
  ScriptApp.newTrigger("onEditInstalled")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log("✅ Trigger 'onEditInstalled' creato con successo!");
}

function createOnEditTrigger() {
  ScriptApp.newTrigger("onEditInstalled")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  Logger.log("✅ Trigger onEditInstalled creato correttamente!");
}

function createTriggerCheckForNewRequests() {
  // Cancella eventuali duplicati
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "checkForNewRequests") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Crea un trigger che gira ogni 5 minuti
  ScriptApp.newTrigger("checkForNewRequests")
    .timeBased()
    .everyMinutes(5)
    .create();
}

function setupDailyReminderTrigger() {
  ScriptApp.newTrigger("sendPersistentReminders")
    .timeBased()
    .everyDays(1)
    .atHour(9) // Invia l'email ogni giorno alle 9:00
    .create();
  Logger.log("✅ Trigger per il promemoria giornaliero creato.");
}

function setupDashboardFridayTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(
    (t) => t.getHandlerFunction() === "updateDashboardFromMain"
  );

  if (exists) {
    Logger.log("✅ Il trigger per 'updateDashboardFromMain' esiste già.");
    return;
  }

  ScriptApp.newTrigger("updateDashboardFromMain")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(15)
    .create();

  Logger.log(
    "✅ Trigger creato: la dashboard sarà aggiornata ogni venerdì alle 15:00."
  );
}

function setupEmailProcessingTrigger() {
  ScriptApp.newTrigger("processEmailQueue")
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log("✅ Trigger per svuotare la coda email creato.");
}

function setupProgramTrigger() {
  ScriptApp.newTrigger("avviaProgramma")
    .timeBased()
    .everyMinutes(10) // Esegue ogni 10 minuti (puoi personalizzarlo)
    .create();
}

function setupReminderTrigger() {
  ScriptApp.newTrigger("sendReminderForUncontactedClients")
    .timeBased()
    .everyDays(1)
    .atHour(9) // Invia l'email ogni giorno alle 9:00
    .create();
  Logger.log("✅ Trigger per il promemoria venditori creato.");
}

function setupWeeklyReportTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendWeeklyReport") {
      Logger.log("✅ Trigger 'sendWeeklyReport' già esistente.");
      return;
    }
  }

  ScriptApp.newTrigger("sendWeeklyReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  Logger.log("✅ Trigger per il riepilogo settimanale creato con successo!");
}

function createAutoCloseTrigger() {
  ScriptApp.newTrigger("autoCloseOldQuotes").timeBased().everyDays(7).create();

  Logger.log("✅ Trigger settimanale creato per autoCloseOldQuotes()");
}

function forceInstallMainOnEditTrigger() {
  const functionName = "onEditInstalled_Main";
  const ss = SpreadsheetApp.getActive();

  // ✅ Ricreo il trigger corretto
  ScriptApp.newTrigger(functionName).forSpreadsheet(ss).onEdit().create();

  Logger.log(
    "✅ Trigger reinstallato per " +
      functionName +
      " sul foglio " +
      ss.getName()
  );
}
