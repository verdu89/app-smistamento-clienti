/** Web Endpoints
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */


function doGet(e) {
  var callback = e.parameter.callback;

  // ✅ Chiamata interna a doPost per salvare i dati
  var response = doPost(e, true);

  // ✅ FORZA LA RISPOSTA JSONP
  return ContentService.createTextOutput(
    callback + "(" + JSON.stringify(response) + ")"
  ).setMimeType(ContentService.MimeType.JAVASCRIPT);
}


function doPost(e, isJsonp = false) {
  var sheetId = "1jjA9ix4LkxAiBOKTIhgg32IBUnR5GAgWtIdIhjznTYI";
  var sheetName = "Main";

  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("❌ ERRORE: Il foglio NON è stato trovato!");
    return { success: false, message: "Foglio non trovato!" };
  }

  var params = e.parameter;
  Logger.log("📩 Dati ricevuti:", params);

  var data = new Date().toLocaleString();
  var nome = params.Nome || "";
  var telefono = params.Telefono || "";
  var email = params.Email || "";
  var provincia = params.Provincia || "";
  var luogoConsegna = params["Luogo di Consegna"] || "";
  var messaggio = params.Messaggio || "";

  Logger.log("📝 Scrivendo i seguenti dati nel foglio:");
  Logger.log("📅 Data e ora: " + data);
  Logger.log("👤 Nome: " + nome);
  Logger.log("📞 Telefono: " + telefono);
  Logger.log("📧 Email: " + email);
  Logger.log("📍 Provincia: " + provincia);
  Logger.log("🏠 Luogo di Consegna: " + luogoConsegna);
  Logger.log("💬 Messaggio: " + messaggio);

  try {
    // ✅ TROVA LA PRIMA RIGA VUOTA DOPO L’INTESTAZIONE
    var lastRow = sheet.getLastRow();
    var nextRow = lastRow + 1;

    // ✅ SCRIVE I DATI NELLA PRIMA RIGA DISPONIBILE
    sheet
      .getRange(nextRow, 1, 1, 7)
      .setValues([
        [data, nome, telefono, email, provincia, luogoConsegna, messaggio],
      ]);

    Logger.log("✅ Riga scritta nella riga " + nextRow);
  } catch (error) {
    Logger.log("❌ ERRORE durante l'inserimento dei dati: " + error.message);
    return {
      success: false,
      message: "Errore nell'inserimento dei dati: " + error.message,
    };
  }

  var response = {
    success: true,
    message: "Dati salvati con successo!",
    dati: params,
  };

  if (isJsonp) {
    return response;
  } else {
    return ContentService.createTextOutput(
      JSON.stringify(response)
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
