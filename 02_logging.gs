/** Logging
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */

function logInfo(msg) {
  Logger.log(msg);
}
function logWarning(message) {
  Logger.log("⚠️ " + message);
  writeToLogSheet("WARNING", message);
}
function logError(messaggio) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log") || ss.insertSheet("Log");
  logSheet.appendRow([new Date().toLocaleString(), "Errore", messaggio]);
}
