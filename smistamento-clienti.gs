/******************************************************************************
 * This tutorial is based on the work of Martin Hawksey twitter.com/mhawksey  *
 * But has been simplified and cleaned up to make it more beginner friendly   *
 * All credit still goes to Martin and any issues/complaints/questions to me. *
 ******************************************************************************/

/**var TO_ADDRESS = "newsaverplast@gmail.com"; // change this ...

function formatMailBody(obj) { // function to spit out all the keys/values from the form in HTML
  var result = "";
  for (var key in obj) { // loop over the object passed to the function
    result += "<h5 style='text-transform: capitalize; margin-bottom: 0'>" + key + "</h5><div>" + obj[key] + "</div>";
    // for every key, concatenate an `<h4 />`/`<div />` pairing of the key name and its value, 
    // and append it to the `result` string created at the start.
  }
  return result; // once the looping is done, `result` will be one long string to put in the email body
}*/

/**
 * record_data inserts the data received from the html form submission
 * e is the data received from the POST
 */
function record_data(e) {
  Logger.log(JSON.stringify(e)); // log the POST data in case we need to debug it
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Foglio3"); // select the responses sheet
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1; // get next row
    var row = [new Date()]; // first element in the row should always be a timestamp
    // loop through the header columns
    for (var i = 1; i < headers.length; i++) {
      // start at 1 to avoid Timestamp column
      if (headers[i].length > 0) {
        row.push(e.parameter[headers[i]]); // add data to row
      }
    }
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
  } catch (error) {
    Logger.log(e);
  } finally {
    return;
  }
}

function isDuplicateEntry(sheet, requestData, colsMain) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (
      data[i][colsMain["Nome"]] === requestData.Nome &&
      data[i][colsMain["Telefono"]] === requestData.Telefono &&
      data[i][colsMain["Email"]] === requestData.Email
    ) {
      return true; // Trovato duplicato
    }
  }
  return false;
}

function getFirstEmptyRow(sheet) {
  if (!sheet) {
    Logger.log("❌ ERRORE: Il foglio non è definito in getFirstEmptyRow.");
    return 2; // Ritorna la riga 2 come fallback
  }

  var lastRow = sheet.getLastRow(); // Ottiene l'ultima riga "usata"

  if (lastRow === 0) return 2; // Se non ci sono dati, ritorna la prima riga

  var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i].every((cell) => cell === "")) {
      return i + 1; // Trova la prima riga vuota effettiva
    }
  }

  return lastRow + 1; // Se non ci sono righe completamente vuote, ritorna la successiva
}

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

function pulisciFogli() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var foglio = ss.getActiveSheet();
  var ultimaRiga = foglio.getLastRow(); // Trova l'ultima riga effettivamente usata
  var ultimaColonna = foglio.getLastColumn(); // Trova l'ultima colonna usata
  var maxRighe = foglio.getMaxRows(); // Numero massimo di righe
  var maxColonne = foglio.getMaxColumns(); // Numero massimo di colonne

  // Assicuriamoci di non eliminare tutte le righe
  if (maxRighe > ultimaRiga && ultimaRiga > 0) {
    foglio.deleteRows(ultimaRiga + 1, maxRighe - ultimaRiga);
  }

  // Assicuriamoci di non eliminare tutte le colonne
  if (maxColonne > ultimaColonna && ultimaColonna > 0) {
    foglio.deleteColumns(ultimaColonna + 1, maxColonne - ultimaColonna);
  }

  // Rimuove la formattazione in eccesso sulle righe vuote rimaste
  foglio
    .getRange(ultimaRiga + 1, 1, maxRighe - ultimaRiga, maxColonne)
    .clearFormat();
}

function syncMainToVendors() {
  const changesLog = []; // tiene traccia di tutte le modifiche

  // 🔒 Evita run sovrapposti
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("⛔ Esecuzione già in corso, esco.");
    return;
  }

  try {
    Logger.log("🚀 Avvio syncMainToVendors()");
    aggiornaNumeroPezziInMain(); // lasciata come nel tuo originale

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Main");
    if (!mainSheet) {
      Logger.log("❌ ERRORE: Il foglio 'Main' non esiste!");
      return;
    }

    var data = mainSheet.getDataRange().getValues();
    var headers = data[0];
    var colsMain = getColumnIndexes(headers);
    var vendors = getVendors();
    var provinceToVendor = getProvinceToVendor();

    // 🔹 Controllo se la colonna "Email" esiste
    if (!("Email" in colsMain)) {
      Logger.log(
        "❌ ERRORE: La colonna 'Email' non è stata trovata in 'Main'!"
      );
      return;
    }

    createBackup(mainSheet);

    var vendorsData = {}; // Memorizza dati per ogni venditore
    var updates = []; // Per aggiornare "Venditore Assegnato" (idempotente)

    var startIndex = -1; // Trova la prima riga con un nominativo senza venditore

    for (var index = 1; index < data.length; index++) {
      // Inizia dalla riga 2 (salta intestazione)
      var row = data[index];

      var nomeCliente = row[colsMain["Nome"]]
        ? row[colsMain["Nome"]].toString().trim()
        : "";
      var telefonoCliente = row[colsMain["Telefono"]]
        ? row[colsMain["Telefono"]].toString().trim()
        : "";
      var venditoreAssegnato = row[colsMain["Venditore Assegnato"]]
        ? row[colsMain["Venditore Assegnato"]].toString().trim()
        : "";
      var emailCliente = row[colsMain["Email"]]
        ? row[colsMain["Email"]].toString().trim()
        : "";
      var luogoConsegna = row[colsMain["Luogo di Consegna"]]
        ? row[colsMain["Luogo di Consegna"]].toString().trim()
        : "";
      var messaggio = row[colsMain["Messaggio"]]
        ? row[colsMain["Messaggio"]].toString().trim()
        : "";

      // Se trova una riga completamente vuota (Nome e Telefono assenti), si ferma
      if (nomeCliente === "" && telefonoCliente === "") {
        Logger.log(
          "🛑 Righe vuote trovate. Interruzione alla riga " + (index + 1) + "."
        );
        break;
      }

      // Trova la prima riga senza venditore assegnato
      if (startIndex === -1 && venditoreAssegnato === "") {
        startIndex = index;
      }

      // Se il venditore è già assegnato, lo ignora
      if (venditoreAssegnato !== "") {
        continue;
      }

      var provincia = row[colsMain["Provincia"]]
        ? row[colsMain["Provincia"]].toString().trim().toLowerCase()
        : "";

      // Liste comuni personalizzate
      var comuniPerCristianPiga = [
        "arzana",
        "bari sardo",
        "baunei",
        "cardedu",
        "elini",
        "gairo",
        "girasole",
        "ilbono",
        "jerzu",
        "lanusei",
        "loceri",
        "lotzorai",
        "osini",
        "perdasdefogu",
        "seui",
        "seulo",
        "talana",
        "tertenia",
        "tortolì",
        "tortoli",
        "triei",
        "ulassai",
        "urzulei",
        "ussassai",
        "villagrande strisaili",
        "villanova strisaili",
        "barisardo",
      ];
      var comuniPerMircko = [
        "capoterra",
        "villasor",
        "serramanna",
        "san sperate",
        "monastir",
        "nuraminis",
        "ussana",
        "dolianova",
        "soleminis",
        "decimoputzu",
        "villaspeciosa",
        "villa san pietro",
      ];

      var venditoreNuovo = "Cristian Piga"; // fallback

      if (provincia === "nu" || provincia === "nuoro") {
        var luogoConsegnaLowerNU = luogoConsegna.toLowerCase();
        var matchCristian = comuniPerCristianPiga.some((comune) =>
          luogoConsegnaLowerNU.includes(comune)
        );
        venditoreNuovo = matchCristian ? "Cristian Piga" : "Marco Guidi";
        Logger.log(
          "📌 Assegnazione NU: '" + luogoConsegna + "' → " + venditoreNuovo
        );
      } else if (provincia === "ca" || provincia === "cagliari") {
        var luogoConsegnaLowerCA = luogoConsegna.toLowerCase();
        var comuniPerCristianInCa = ["pula", "villasimius"];
        var matchCristianCA = comuniPerCristianInCa.some((comune) =>
          luogoConsegnaLowerCA.includes(comune)
        );
        venditoreNuovo = matchCristianCA ? "Cristian Piga" : "Mircko Manconi";
        Logger.log(
          "📌 Assegnazione CA: '" + luogoConsegna + "' → " + venditoreNuovo
        );
      } else if (provincia === "su" || provincia === "sud sardegna") {
        var luogoConsegnaLowerSU = luogoConsegna.toLowerCase();
        var matchMircko = comuniPerMircko.some((comune) =>
          luogoConsegnaLowerSU.includes(comune)
        );
        venditoreNuovo = matchMircko ? "Mircko Manconi" : "Cristian Piga";
        Logger.log(
          "📌 Assegnazione SU: '" + luogoConsegna + "' → " + venditoreNuovo
        );
      } else {
        // === LOGICA PERSONALIZZATA PER SASSARI ===
        var pezzi = row[colsMain["Numero pezzi"]]
          ? parseInt(row[colsMain["Numero pezzi"]], 10)
          : 0;

        if ((provincia === "ss" || provincia === "sassari") && pezzi > 7) {
          venditoreNuovo = "Cristian Piga";
          Logger.log(
            "📌 Assegnazione SS con " + pezzi + " pezzi → Cristian Piga"
          );
        } else {
          venditoreNuovo = provinceToVendor[provincia] || "Cristian Piga";
          Logger.log(
            "📌 Assegnazione standard: Provincia '" +
              provincia +
              "' → " +
              venditoreNuovo
          );
        }
        // === FINE LOGICA PERSONALIZZATA ===
      }

      // 🔹 Pianifica aggiornamento venditore (idempotente)
      updates.push([index + 1, venditoreNuovo]);

      // 🔹 Se "Data e ora" è vuota, scriviamo la data corrente
      if (!row[colsMain["Data e ora"]]) {
        mainSheet
          .getRange(index + 1, colsMain["Data e ora"] + 1)
          .setValue(new Date());
      }

      // 🔹 PRIMA ASSEGNAZIONE: scrivi subito e invia email una sola volta
      if (!row[colsMain["Data Assegnazione"]]) {
        const now = new Date();

        // ✍️ Scrive immediatamente "Data Assegnazione" e "Venditore Assegnato"
        mainSheet
          .getRange(index + 1, colsMain["Data Assegnazione"] + 1)
          .setValue(now);
        changesLog.push(`Riga ${index + 1}: scritta Data Assegnazione`);
        mainSheet
          .getRange(index + 1, colsMain["Venditore Assegnato"] + 1)
          .setValue(venditoreNuovo);
        changesLog.push(
          `Riga ${index + 1}: assegnato Venditore → ${venditoreNuovo}`
        );

        // 🔒 Forza la scrittura prima di inviare l'email (riduce rischio doppio invio)
        SpreadsheetApp.flush();

        // 📩 Notifica SEMPRE venditore e azienda;
        //     al cliente solo se l'email è valida.
        //     Se l'email è assente/non valida, scriviamo una nota (se la colonna "Note" esiste).
        Logger.log(
          "📨 Preparazione notifiche - Cliente email: " +
            (emailCliente || "(vuota)")
        );

        notifyAssignment(
          venditoreNuovo,
          emailCliente || "",
          nomeCliente,
          telefonoCliente,
          provincia,
          luogoConsegna,
          messaggio
        );

        // Se email cliente mancante o non valida, aggiungi nota (se c'è la colonna "Note")
        if (!isValidEmail_(emailCliente)) {
          safeSetIfColumnExists_(
            mainSheet,
            colsMain,
            "Note",
            index + 1,
            "Email cliente assente o non valida"
          );
          changesLog.push(
            `Riga ${
              index + 1
            }: aggiunta Nota 'Email cliente assente o non valida'`
          );
          Logger.log(
            "ℹ️ Nota aggiunta in 'Main': Email cliente assente o non valida (riga " +
              (index + 1) +
              ")"
          );
        }
      }

      // 🔹 Prepara dati per i fogli venditori
      if (!vendorsData[venditoreNuovo]) {
        vendorsData[venditoreNuovo] = [];
      }

      var filteredRow = {};
      Object.keys(colsMain).forEach(function (col) {
        filteredRow[col] = row[colsMain[col]];
      });
      filteredRow["Data Assegnazione"] = new Date().toLocaleString();
      vendorsData[venditoreNuovo].push(filteredRow);
    }

    // 🔹 Scrive gli aggiornamenti nel foglio "Main" (idempotente)
    updates.forEach(function (update) {
      var r = update[0];
      var venditore = update[1];
      mainSheet
        .getRange(r, colsMain["Venditore Assegnato"] + 1)
        .setValue(venditore);
    });

    // 🔁 Sincronizza sui fogli venditori (con deduplica in quella funzione)
    syncVendorsSheets(vendorsData, vendors);

    Logger.log("✅ Fine syncMainToVendors()");
  } finally {
    lock.releaseLock();
  }
  Logger.log("📋 Dettaglio modifiche:");
  changesLog.slice(0, 50).forEach((msg) => Logger.log(msg)); // prime 50 per non intasare il log
  Logger.log(`Totale modifiche loggate: ${changesLog.length}`);
}

function syncVendorsSheets(vendorsData, vendors) {
  Object.keys(vendorsData).forEach((venditore) => {
    var vendorSS = SpreadsheetApp.openById(vendors[venditore]);
    var venditoreSheet =
      vendorSS.getSheetByName("Dati") || vendorSS.insertSheet("Dati");

    // Leggi contenuto esistente in modo robusto
    var dataVendor = venditoreSheet.getDataRange().getValues();
    var hasHeader =
      dataVendor &&
      dataVendor.length > 0 &&
      dataVendor[0] &&
      dataVendor[0].length > 0;

    // Se il foglio è vuoto, inizializza un set base di intestazioni compatibile
    if (!hasHeader) {
      var headersVendorInit = [
        "Nome",
        "Telefono",
        "Email",
        "Provincia",
        "Luogo di Consegna",
        "Messaggio",
        "Data Assegnazione",
        "Stato",
        "Vendita Conclusa?",
      ];
      venditoreSheet
        .getRange(1, 1, 1, headersVendorInit.length)
        .setValues([headersVendorInit]);
      dataVendor = venditoreSheet.getDataRange().getValues();
      hasHeader = true;
    }

    var headersVendor = dataVendor[0];
    var colsVendor = getColumnIndexes(headersVendor);

    // Mappatura (rimane invariata rispetto alla tua logica)
    var columnMapping = {
      Nome: "Nome",
      Telefono: "Telefono",
      Email: "Email",
      Provincia: "Provincia",
      "Luogo di Consegna": "Luogo di Consegna",
      Messaggio: "Messaggio",
      "Data Assegnazione": "Data Assegnazione",
      Stato: "Stato",
      // "Vendita Conclusa?" verrà gestita più giù come default quando presente tra le intestazioni
    };

    // 🔒 Costruisci un set delle chiavi già presenti (nome|telefono) nel foglio venditore
    var existingKeys = new Set();
    for (var i = 1; i < dataVendor.length; i++) {
      var n = (dataVendor[i][colsVendor["Nome"]] || "")
        .toString()
        .trim()
        .toLowerCase();
      var t = (dataVendor[i][colsVendor["Telefono"]] || "").toString().trim();
      if (n || t) existingKeys.add(n + "|" + t);
    }

    // 🔁 Evita duplicati anche nella stessa esecuzione (batch corrente)
    var seenInThisRun = new Set();
    var rowsToAdd = [];

    vendorsData[venditore].forEach((row) => {
      var nome = (row["Nome"] || "").toString().trim().toLowerCase();
      var tel = (row["Telefono"] || "").toString().trim();
      if (!nome && !tel) return; // riga non valida

      var key = nome + "|" + tel;
      if (existingKeys.has(key) || seenInThisRun.has(key)) {
        // Già presente: salta
        return;
      }
      seenInThisRun.add(key);

      // Costruisci la riga nel corretto ordine headersVendor
      var newRow = headersVendor.map((header) => {
        if (header === "Data Assegnazione") return new Date().toLocaleString();
        if (header === "Stato") return "Da contattare"; // default
        if (header === "Vendita Conclusa?") return ""; // default
        var mainColumn = Object.keys(columnMapping).find(
          (k) => columnMapping[k] === header
        );
        return mainColumn && row[mainColumn] !== undefined
          ? row[mainColumn]
          : "";
      });

      rowsToAdd.push(newRow);
    });

    if (rowsToAdd.length > 0) {
      var startRow = venditoreSheet.getLastRow() + 1;
      venditoreSheet
        .getRange(startRow, 1, rowsToAdd.length, headersVendor.length)
        .setValues(rowsToAdd);
    }

    // 🔽 Dropdown invariati
    applyDropdownIfColumnExists(venditoreSheet, "Stato", [
      "Da contattare",
      "Preventivo inviato",
      "Preventivo non eseguibile",
      "In trattativa",
      "Trattativa terminata",
    ]);

    applyDropdownIfColumnExists(
      venditoreSheet,
      "Vendita Conclusa?",
      ["SI", "NO"],
      { SI: "#00FF00", NO: "#FF0000" }
    );
  });
}

/**
 * Funzione per applicare il menu a tendina SOLO se la colonna esiste nel foglio
 */
function applyDropdownIfColumnExists(sheet, columnName, values, colors = null) {
  var headers = sheet.getDataRange().getValues()[0]; // Legge le intestazioni
  var colIndex = headers.indexOf(columnName); // Trova la posizione della colonna

  if (colIndex === -1) {
    Logger.log(`⚠️ La colonna "${columnName}" non esiste nel foglio.`);
    return;
  }

  colIndex += 1; // Converti l'indice da 0-based a 1-based per Google Sheets

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // Nessun dato nel foglio oltre le intestazioni

  var range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);

  if (colors) {
    for (var i = 2; i <= lastRow; i++) {
      var cell = sheet.getRange(i, colIndex);
      var value = cell.getValue().toString().trim();
      if (value in colors) {
        cell.setBackground(colors[value]);
      } else {
        cell.setBackground("#FFFFFF");
      }
    }
  }

  Logger.log(
    `✅ Menu a tendina applicato per "${columnName}" alla colonna ${colIndex}.`
  );
}

function testRowCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var data = mainSheet.getDataRange().getValues();
  Logger.log("🔎 Numero effettivo di righe lette: " + data.length);
}

function debugMainSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var data = mainSheet.getDataRange().getValues();
  Logger.log("📌 Dati dal foglio Main: " + JSON.stringify(data.slice(0, 5))); // Mostra le prime 5 righe
}

function debugLastProcessedRow() {
  var scriptProperties = PropertiesService.getScriptProperties(); // Definizione della variabile
  var lastProcessedRow = scriptProperties.getProperty("lastProcessedRow");

  if (lastProcessedRow === null) {
    Logger.log(
      "🔎 lastProcessedRow non esiste ancora nelle proprietà dello script."
    );
  } else {
    Logger.log("🔎 Valore di lastProcessedRow: " + lastProcessedRow);
  }
}

function updateMainFromVendors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var dataMain = mainSheet.getDataRange().getValues();
  var headersMain = dataMain[0];
  var colsMain = getColumnIndexes(headersMain);

  var vendors = getVendors(); // Recupera l'elenco dei venditori

  var updatableColumns = [
    "Stato",
    "Note",
    "Data Preventivo",
    "Importo Preventivo",
    "Vendita Conclusa?",
    "Intestatario Contratto",
  ]; // 🔹 Colonne aggiornabili

  var updates = []; // Array per raccogliere gli aggiornamenti da eseguire in batch

  for (var venditore in vendors) {
    try {
      var vendorSS = SpreadsheetApp.openById(vendors[venditore]);
      var vendorSheet = vendorSS.getSheetByName("Dati");
      if (!vendorSheet) continue;

      var dataVendor = vendorSheet.getDataRange().getValues();
      var headersVendor = dataVendor[0];
      var colsVendor = getColumnIndexes(headersVendor);

      for (var i = 1; i < dataVendor.length; i++) {
        var vendorRow = dataVendor[i];
        var vendorNome = vendorRow[colsVendor["Nome"]];
        var vendorTelefono = vendorRow[colsVendor["Telefono"]];

        for (var j = 1; j < dataMain.length; j++) {
          var mainRow = dataMain[j];

          // 🔍 Confronta Nome e Telefono per trovare la corrispondenza nel foglio "Main"
          if (
            mainRow[colsMain["Nome"]] === vendorNome &&
            mainRow[colsMain["Telefono"]] === vendorTelefono
          ) {
            var rowIndex = j + 1;
            var rowUpdates = []; // Memorizza aggiornamenti per questa riga

            // 🔹 Ora aggiorna SEMPRE le colonne aggiornabili
            updatableColumns.forEach((col) => {
              if (col in colsVendor && col in colsMain) {
                var vendorValue = vendorRow[colsVendor[col]];
                var mainValue = mainRow[colsMain[col]];

                // 🔹 Se il valore del venditore è diverso da quello in Main, aggiornalo
                if (
                  vendorValue !== "" &&
                  vendorValue !== undefined &&
                  vendorValue !== mainValue
                ) {
                  rowUpdates.push([colsMain[col] + 1, vendorValue]); // [colonna, nuovo valore]
                }
              }
            });

            if (rowUpdates.length > 0) {
              updates.push({ rowIndex, rowUpdates });
            }
            break; // Interrompe il ciclo una volta trovata la riga corrispondente
          }
        }
      }
    } catch (e) {
      Logger.log(`❌ Errore aggiornando da ${venditore}: ${e.message}`);
    }
  }

  // 🔹 Applica gli aggiornamenti al foglio "Main" in batch (più veloce)
  updates.forEach((update) => {
    update.rowUpdates.forEach(([colIndex, value]) => {
      mainSheet.getRange(update.rowIndex, colIndex).setValue(value);
    });
  });

  Logger.log(
    `✅ Aggiornamento completato: ${updates.length} righe modificate in "Main".`
  );
}

/**
 * 🔹 Funzione per aggiungere più colonne nuove in "Main"
 */
function addMultipleColumnsToMain(sheet, columnNames) {
  var headers = sheet.getDataRange().getValues()[0];
  var existingColumns = new Set(headers); // 🔹 Contiene le colonne già presenti
  var lastCol = headers.length + 1;

  columnNames.forEach((colName, index) => {
    if (!existingColumns.has(colName)) {
      // 🔹 Aggiunge solo se la colonna non esiste
      sheet.getRange(1, lastCol).setValue(colName);
      Logger.log(
        `✅ Aggiunta nuova colonna "${colName}" in "Main" alla posizione ${lastCol}`
      );
      lastCol++;
    } else {
      Logger.log(
        `⚠️ La colonna "${colName}" esiste già in "Main", non verrà aggiunta.`
      );
    }
  });
}

function syncToVendorSheet(row, venditore, vendors, colsMain) {
  if (!(venditore in vendors)) {
    logError("❌ Nessun foglio venditore trovato per: " + venditore);
    return;
  }

  try {
    var vendorSS = SpreadsheetApp.openById(vendors[venditore]);
    var vendorSheet = vendorSS.getSheetByName("Dati");
    if (!vendorSheet) {
      logError("❌ Il foglio 'Dati' non esiste nel file di " + venditore);
      return;
    }

    var vendorData = vendorSheet.getDataRange().getValues();
    var colsVendor = getColumnIndexes(vendorData[0]);

    // Verifica se l'entry esiste già
    if (isAlreadyAssigned(row, colsMain, vendorData, colsVendor)) {
      logInfo("🔁 Cliente già presente nel foglio di " + venditore);
      return;
    }

    addToVendorSheet(row, vendorSheet, colsMain, colsVendor);
  } catch (e) {
    logError(
      "❌ Errore durante la sincronizzazione con " +
        venditore +
        ": " +
        e.message
    );
  }
}

function addToVendorSheet(row, sheet, colsMain, colsVendor) {
  logInfo("➡️ Avvio aggiunta dati a " + sheet.getName());

  if (!colsVendor || Object.keys(colsVendor).length === 0) {
    logError("❌ Errore: colsVendor è vuoto o non definito!");
    return;
  }

  var newRow = new Array(Object.keys(colsVendor).length).fill("-");

  if (colsVendor["Data Assegnazione"] !== undefined) {
    newRow[colsVendor["Data Assegnazione"]] = new Date().toISOString();
  }

  if (colsVendor["Stato"] !== undefined) {
    newRow[colsVendor["Stato"]] = "Da contattare"; // Valore predefinito
  }

  if (colsVendor["Vendita Conclusa?"] !== undefined) {
    newRow[colsVendor["Vendita Conclusa?"]] = ""; // Casella vuota
  }

  for (var colName in colsMain) {
    if (colsVendor[colName] !== undefined) {
      var value = row[colsMain[colName]];
      if (
        value !== undefined &&
        value !== null &&
        value.toString().trim() !== ""
      ) {
        newRow[colsVendor[colName]] = value;
      }
    }
  }

  try {
    sheet.appendRow(newRow);
    SpreadsheetApp.flush(); // Forza aggiornamento prima di applicare la convalida
    logInfo("✅ Riga inserita per " + row[colsMain["Nome"]]);

    var lastRow = sheet.getLastRow();
    if (colsVendor["Stato"] !== undefined) {
      applyDropdownValidation(sheet, colsVendor["Stato"], [
        "Da contattare",
        "Preventivo inviato",
        "Preventivo non eseguibile",
        "In trattativa",
        "Trattativa terminata",
      ]);
    }
    if (colsVendor["Vendita Conclusa?"] !== undefined) {
      applyDropdownValidation(
        sheet,
        colsVendor["Vendita Conclusa?"],
        ["SI", "NO"],
        { SI: "#00FF00", NO: "#FF0000" },
        lastRow
      );
    }
  } catch (e) {
    logError("❌ Errore durante l'inserimento della riga: " + e.message);
  }
}

function onEditInstalled(e) {
  if (!e || !e.source || !e.range) {
    Logger.log("⚠️ Errore: evento 'e' non è definito.");
    return;
  }

  var sheet = e.source.getActiveSheet();
  var fogliAbilitati = ["Main", "Recensioni Extra"];
  if (!fogliAbilitati.includes(sheet.getName())) return;

  var editedCell = e.range;
  var data = sheet.getDataRange().getValues();
  var cols = getColumnIndexes(data[0]); // Rende la funzione compatibile con più fogli

  // 🔍 Colonne richieste, ma alcune possono non essere presenti in "Recensioni Extra"
  var requiredCols = [
    "Richiedi Recensione",
    "Email",
    "Data richiesta recensione",
  ];
  var optionalColVendita = "Vendita Conclusa?";

  var missingCols = requiredCols.filter((col) => !(col in cols));
  if (missingCols.length > 0) {
    logWarning(
      "⚠️ Colonne mancanti nel foglio '" +
        sheet.getName() +
        "': " +
        missingCols.join(", ")
    );
    return;
  }

  // 📌 Cambia colore se è presente la colonna "Vendita Conclusa?"
  if (optionalColVendita in cols) {
    var colVendita = cols[optionalColVendita] + 1;
    if (editedCell.getColumn() === colVendita) {
      SpreadsheetApp.flush();
      Utilities.sleep(200);

      var newValue = editedCell.getValue().toString().trim();
      var colors = { SI: "#00FF00", NO: "#FF0000" };
      var color = colors[newValue] || "#FFFFFF";

      var validation = editedCell.getDataValidation();
      editedCell.setDataValidation(null);
      editedCell.setBackground(color);
      editedCell.setDataValidation(validation);

      logInfo(`🎨 Colore aggiornato per 'Vendita Conclusa?' a: ${newValue}`);
    }
  }

  // 📩 Invio email recensione se flag su "Richiedi Recensione"
  var colReview = cols["Richiedi Recensione"] + 1;
  if (editedCell.getColumn() === colReview && editedCell.getValue() === true) {
    var row = editedCell.getRow();
    var email = sheet.getRange(row, cols["Email"] + 1).getValue();
    var dataRecensione = sheet
      .getRange(row, cols["Data richiesta recensione"] + 1)
      .getValue();

    if (!email) {
      logWarning(
        "⚠️ Nessuna email trovata nella riga " +
          row +
          ", impossibile inviare richiesta recensione."
      );
      return;
    }

    if (dataRecensione) {
      logInfo("ℹ️ Recensione già inviata in data: " + dataRecensione);
      return;
    }

    var rawNomeCliente =
      sheet.getRange(row, cols["Nome"] + 1).getValue() || "Cliente";
    var nomeCliente = formatNameProperly(rawNomeCliente.toString().trim());

    var reviewEmailBody =
      "Gentile " +
      nomeCliente +
      ",<br><br>" +
      "Siamo davvero felici di averLa accompagnata nella scelta dei Suoi nuovi infissi.<br><br>" +
      "Per noi ogni progetto è speciale, e nulla ci rende più orgogliosi di sapere che i nostri clienti sono soddisfatti del lavoro svolto.<br><br>" +
      "Se ha apprezzato il nostro servizio, La invitiamo a raccontarlo lasciando una recensione: è un piccolo gesto per Lei, ma per noi è fondamentale.<br><br>" +
      "<b>👉 Clicchi qui per lasciare la Sua opinione:</b><br>" +
      "<a href='https://maps.app.goo.gl/1gM31niwMtSfPCk16'>https://maps.app.goo.gl/1gM31niwMtSfPCk16</a><br><br>" +
      "Le recensioni aiutano altri clienti come Lei a fidarsi di noi e ci stimolano ogni giorno a migliorare.<br><br>" +
      "Grazie per il tempo che vorrà dedicarci!<br><br>" +
      "Con stima,<br>" +
      "<b>Il Team Saverplast</b>";

    sendEmail(email, "Come si è trovato con Saverplast?", reviewEmailBody);
    sheet
      .getRange(row, cols["Data richiesta recensione"] + 1)
      .setValue(new Date().toLocaleDateString());
    logInfo("📧 Email richiesta recensione inviata a: " + email);
  }
}

function formatNameProperly(name) {
  return name
    .toLowerCase()
    .split(" ")
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function applyDropdownValidation(
  sheet,
  colIndex,
  values,
  colors = null,
  lastRow = null
) {
  if (lastRow === null) {
    lastRow = sheet.getLastRow();
  }
  if (lastRow < 2) return;

  var range = sheet.getRange(lastRow, colIndex + 1, 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);

  if (colors) {
    var cellValue = range.getValue().toString().trim();
    if (cellValue in colors) {
      range.setBackground(colors[cellValue]);
    } else {
      range.setBackground("#FFFFFF");
    }
  }
}

function applyUpdates(sheet, updates, colsMain) {
  updates.forEach((update) => {
    Logger.log(
      "✍️ Scrittura venditore: " + update[1] + " sulla riga " + update[0]
    );

    sheet
      .getRange(update[0], colsMain["Venditore Assegnato"] + 1)
      .setValue(update[1]);
    sheet.getRange(update[0], colsMain["Data e ora"] + 1).setValue(update[2]);
  });
}

function isAlreadyAssigned(row, colsMain, vendorData, colsVendor) {
  var nomeCliente = row[colsMain["Nome"]].toString().trim().toLowerCase();
  var telefonoCliente = row[colsMain["Telefono"]].toString().trim();

  return vendorData.some((vRow) => {
    var nomeVenditore = vRow[colsVendor["Nome"]]
      ? vRow[colsVendor["Nome"]].toString().trim().toLowerCase()
      : "";
    var telefonoVenditore = vRow[colsVendor["Telefono"]]
      ? vRow[colsVendor["Telefono"]].toString().trim()
      : "";

    return (
      nomeCliente === nomeVenditore && telefonoCliente === telefonoVenditore
    );
  });
}

function getColumnIndexes(headerRow) {
  if (!headerRow || headerRow.length === 0) {
    Logger.log("❌ ERRORE: Intestazione del foglio vuota!");
    return {};
  }

  var indexes = {};
  headerRow.forEach((colName, index) => {
    var cleanName = colName.toString().trim();
    indexes[cleanName] = index;
  });

  Logger.log("📊 Indici colonne trovati: " + JSON.stringify(indexes));
  return indexes;
}

function getVendors() {
  return {
    "Mircko Manconi": "1mGFlFbCYy9ylVjNA9l6b855c6jlIDr6QOua2qfSjckw",
    "Cristian Piga": "1N0_GKbJvZLQbKKIgfVYW4LQGp97mhQcOz9zsD-FBNcE",
    "Marco Guidi": "1CVQSnFGNX8pGUKUABdtzwQmyCKPtuOsK8XAVbJwmUqE",
  };
}

function getVendorEmail(venditore) {
  var vendorEmails = {
    "Mircko Manconi": "mirckox@yahoo.it",
    "Cristian Piga": "cristianpiga@me.com",
    "Marco Guidi": "guidi.marco0308@libero.it",
  };

  return vendorEmails[venditore] || "newsaverplast@gmail.com"; // Email di default in caso di venditore sconosciuto
}

function getVendorPhone(venditore) {
  var vendorPhones = {
    "Mircko Manconi": "+39 3398123123",
    "Cristian Piga": "+39 3939250786",
    "Marco Guidi": "+39 3349630922",
  };

  return vendorPhones[venditore] || "+39 070/247362"; // Numero di default in caso di venditore sconosciuto
}

function getProvinceToVendor() {
  var provinceToVendor = {
    ca: "Mircko Manconi",
    cagliari: "Mircko Manconi",
    su: "Cristian Piga",
    "sud sardegna": "Cristian Piga",
    or: "Cristian Piga",
    oristano: "Cristian Piga",
    nu: "Marco Guidi",
    nuoro: "Marco Guidi",
    ss: "Marco Guidi",
    sassari: "Marco Guidi",
  };

  Logger.log(
    "📊 Mappatura province-venditori caricata: " +
      JSON.stringify(provinceToVendor)
  );
  return provinceToVendor;
}

function createBackup(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastBackupDateStr = scriptProperties.getProperty("lastBackupDate");
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Pulisce l'orario

  if (lastBackupDateStr) {
    var lastBackupDate = new Date(lastBackupDateStr);
    lastBackupDate.setHours(0, 0, 0, 0);

    var lastMonday = getLastMonday();
    lastMonday.setHours(0, 0, 0, 0);

    if (lastBackupDate >= lastMonday) {
      Logger.log(
        "✅ Il backup è già stato fatto questa settimana (" +
          lastBackupDateStr +
          "). Nessuna azione necessaria."
      );
      return;
    }
  }

  // Se siamo qui, significa che dobbiamo fare un nuovo backup
  var todayFormatted = today.toISOString().split("T")[0]; // Formato YYYY-MM-DD
  var backupSheetName = "Backup_" + todayFormatted;
  var existingSheet = ss.getSheetByName(backupSheetName);
  if (existingSheet) ss.deleteSheet(existingSheet);

  var backupSheet = sheet.copyTo(ss);
  backupSheet.setName(backupSheetName);
  scriptProperties.setProperty("lastBackupDate", today.toISOString()); // Salva anche orario preciso

  Logger.log("📁 Backup creato: " + backupSheetName);
}

/** Log **/

function logInfo(message) {
  Logger.log("✅ " + message);
  writeToLogSheet("INFO", message);
}

function logWarning(message) {
  Logger.log("⚠️ " + message);
  writeToLogSheet("WARNING", message);
}

function logError(message) {
  Logger.log("❌ " + message);
  writeToLogSheet("ERROR", message);
}

function writeToLogSheet(type, message) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log") || ss.insertSheet("Log");

  logSheet.appendRow([new Date().toLocaleString(), type, message]);

  // Mantiene solo le ultime 500 righe per evitare che il log diventi enorme
  var maxRows = 500;
  var numRows = logSheet.getLastRow();
  if (numRows > maxRows) {
    logSheet.deleteRows(2, numRows - maxRows);
  }
}

/**
 * Funzione per inviare un'email
 */
var emailQueue = [];

function sendEmail(to, subject, body) {
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: body,
    });
    logInfo("📧 Email inviata a " + to);
  } catch (e) {
    logError("❌ Errore nell'invio email a " + to + ": " + e.message);
    addToEmailQueue(to, subject, body);
  }
}

/** Helper: email valida? */
function isValidEmail_(email) {
  if (!email || typeof email !== "string") return false;
  const e = email.trim();
  // regex semplice e robusta per casi comuni
  return !!e.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/);
}

/** Helper: scrivi in colonna se esiste */
function safeSetIfColumnExists_(sheet, cols, colName, rowIndex, value) {
  if (cols && colName in cols) {
    sheet.getRange(rowIndex, cols[colName] + 1).setValue(value);
  }
}

/**
 * Invio email al cliente e venditore quando viene assegnato un nominativo
 */
/**
 * Invio notifiche al cliente (se email valida), venditore e azienda
 */
function notifyAssignment(
  venditore,
  clienteEmail,
  clienteNome,
  clienteTelefono,
  provincia,
  luogoConsegna,
  messaggio
) {
  var telefonoVenditore = getVendorPhone(venditore);
  var emailVenditore = getVendorEmail(venditore);

  // 📩 Corpo mail al cliente (solo se email valida)
  var vendorInfo =
    "Gentile Cliente,<br><br>" +
    "La ringraziamo per averci contattato e per l’interesse dimostrato nei nostri prodotti.<br><br>" +
    "Sappiamo quanto sia importante per Lei scegliere infissi di alta qualità che garantiscano <b>comfort, efficienza e sicurezza</b> per la Sua casa. Per questo motivo, ci impegniamo a offrirle le migliori soluzioni su misura, con materiali innovativi e un servizio altamente professionale.<br><br>" +
    "💡 <b>Perché scegliere noi?</b><br>" +
    "✔️ <b>Materiali di alta qualità</b> per il massimo isolamento termico e acustico.<br>" +
    "✔️ <b>Infissi su misura</b>, perfettamente adattabili ai Suoi ambienti.<br>" +
    "✔️ <b>Esperienza e professionalità</b>, con anni di successi nel settore.<br>" +
    "✔️ <b>Offerte esclusive</b>, riservate ai nostri clienti.<br><br>" +
    "Per offrirle un <b>preventivo personalizzato e accurato</b>, il nostro esperto <b>" +
    venditore +
    "</b> sarà presto in contatto con Lei.<br><br>" +
    "Se non lo ha già fatto, per velocizzare il processo, La invitiamo a comunicarci le <b>misure indicative</b> degli infissi di Suo interesse (larghezza x altezza).<br><br>Può inviarle anche via email al referente che Le è stato assegnato: troverà i suoi contatti in fondo a questa email. Questo ci permetterà di elaborare una proposta su misura e illustrarle le soluzioni più vantaggiose per le Sue esigenze.<br><br>" +
    "Nel frattempo, se desidera toccare con mano la qualità dei nostri prodotti e ricevere una consulenza diretta, La aspettiamo nei nostri showroom:<br><br>" +
    "✅ <b><a href='https://maps.app.goo.gl/GCr4L3MBEMCE4Fk76' target='_blank'>Cagliari</a></b>, Via della Pineta<br>" +
    "✅ <b><a href='https://maps.app.goo.gl/1gM31niwMtSfPCk16' target='_blank'>Macchiareddu</a></b>, 5° Strada Zona Ovest<br>" +
    "✅ <b><a href='https://maps.app.goo.gl/saVpoWM62aMoZkpg8' target='_blank'>Nuoro</a></b>, Via Badu e Carros<br><br>" +
    "📞 <b>Il nostro esperto sarà lieto di assisterla e consigliarla nella scelta della soluzione più adatta.</b><br><br>" +
    "<b>Contatti del referente Saverplast:</b><br>" +
    "👤 Nome: " +
    venditore +
    "<br>" +
    "📧 Email: <a href='mailto:" +
    emailVenditore +
    "'>" +
    emailVenditore +
    "</a><br>" +
    "📞 Telefono: <a href='tel:" +
    telefonoVenditore +
    "'>" +
    telefonoVenditore +
    "</a><br><br>" +
    "Restiamo a disposizione per qualsiasi ulteriore informazione e confidiamo di poterla assistere al meglio.<br><br>" +
    "Cordiali saluti,<br>" +
    "<b>Il Team Saverplast</b>";

  // ✅ Cliente: invia SOLO se l'email è valida
  if (isValidEmail_(clienteEmail)) {
    try {
      sendEmail(clienteEmail, "Saverplast - Preventivo richiesto", vendorInfo);
    } catch (e) {
      logError(
        "❌ Errore invio email cliente (" + clienteEmail + "): " + e.message
      );
    }
  } else {
    logWarning(
      "⚠️ Email cliente assente o non valida, salto invio al cliente. Valore: '" +
        (clienteEmail || "") +
        "'"
    );
  }

  // ✅ Venditore: sempre
  var vendorBodyVenditore = `<b>📢 Nuovo nominativo assegnato</b><br>
    <b>Nome:</b> ${clienteNome}<br>
    <b>Telefono:</b> ${clienteTelefono}<br>
    <b>Email:</b> ${clienteEmail || "(non fornita)"}<br>
    <b>Provincia:</b> ${provincia}<br>
    <b>Luogo di Consegna:</b> ${luogoConsegna}<br>
    <b>Messaggio Cliente:</b> ${messaggio}<br><br>
    🔹 <b>Contatta il cliente il prima possibile per finalizzare la vendita!</b>`;
  try {
    sendEmail(
      emailVenditore,
      "📢 Nuovo nominativo assegnato",
      vendorBodyVenditore
    );
  } catch (e) {
    logError(
      "❌ Errore invio email al venditore (" +
        emailVenditore +
        "): " +
        e.message
    );
  }

  // ✅ Azienda: sempre
  var aziendaEmail = "newsaverplast@gmail.com";
  var vendorBodyAzienda = `<b>📢 Nuovo nominativo assegnato a ${venditore} </b><br>
    <b>Nome:</b> ${clienteNome}<br>
    <b>Telefono:</b> ${clienteTelefono}<br>
    <b>Email:</b> ${clienteEmail || "(non fornita)"}<br>
    <b>Provincia:</b> ${provincia}<br>
    <b>Luogo di Consegna:</b> ${luogoConsegna}<br>
    <b>Messaggio Cliente:</b> ${messaggio}<br><br>
    🔹 <b>Messaggio per conoscenza</b>`;

  try {
    sendEmail(
      aziendaEmail,
      "📢[Infissipvcsardegna] Nuovo nominativo assegnato a " + venditore,
      vendorBodyAzienda
    );
  } catch (e) {
    logError("❌ Errore invio email all'azienda: " + e.message);
  }
}

/**
 * Invia riepilogo settimanale ogni lunedì
 */
function getLastMonday() {
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0 = Domenica, 1 = Lunedì, ..., 6 = Sabato
  var daysSinceMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1; // Calcola quanti giorni togliere
  var lastMonday = new Date(today);
  lastMonday.setDate(today.getDate() - daysSinceMonday); // Sottrae una settimana esatta
  lastMonday.setHours(0, 0, 0, 0);

  return lastMonday;
}

function getWeekNumber(date) {
  var tempDate = new Date(
    Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  );
  var dayNum = tempDate.getUTCDay() || 7; // domenica = 7
  tempDate.setUTCDate(tempDate.getUTCDate() + 4 - dayNum); // sposta al giovedì della stessa settimana
  var yearStart = new Date(Date.UTC(tempDate.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil(((tempDate - yearStart) / 86400000 + 1) / 7);
  return weekNo;
}

function sendWeeklyReport() {
  aggiornaNumeroPezziInMain(); // ✅ aggiorna campi mancanti

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var data = mainSheet.getDataRange().getValues();
  var colsMain = getColumnIndexes(data[0]);

  var thisMonday = getLastMonday();
  var startDate = new Date(thisMonday);
  startDate.setDate(startDate.getDate() - 7);
  var endDate = new Date(thisMonday);
  endDate.setDate(endDate.getDate() - 1);

  var weekNumber = getWeekNumber(startDate); // 🔢 settimana dei preventivi

  // Imposta orari precisi
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999);

  var clients = [];
  var totalPezzi = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateCell = row[colsMain["Data e ora"]];
    if (!dateCell) {
      logInfo(`⚠️ Riga ${i + 1}: campo "Data e ora" vuoto`);
      continue;
    }

    var assignedDate = tryParseDate(dateCell);
    Logger.log(
      `🔍 Riga ${i + 1} – Valore raw: "${dateCell}" ➝ Parsed: ${assignedDate}`
    );

    if (!(assignedDate instanceof Date) || isNaN(assignedDate)) {
      logInfo(`⚠️ Riga ${i + 1}: data non valida -> "${dateCell}"`);
      continue;
    }

    // ✅ Clona l'oggetto Date per azzerare l'orario
    var assignedDateMidnight = new Date(assignedDate);
    assignedDateMidnight.setHours(0, 0, 0, 0);

    if (assignedDateMidnight >= startDate && assignedDateMidnight <= endDate) {
      clients.push(row);
    }
  }

  if (clients.length === 0) {
    logInfo("📌 Nessun nuovo cliente registrato la settimana scorsa.");
    return;
  }

  var emailBody = `
  <div style="font-family: Arial; max-width: 800px; margin: auto;">
    <h2 style="text-align:center;">📊 Riepilogo Nuovi Clienti della Settimana</h2>
    <p>🗓️ Settimana <b>#${weekNumber}</b> – dal <b>${startDate.toLocaleDateString()}</b> al <b>${endDate.toLocaleDateString()}</b></p>
    <table border="1" style="border-collapse: collapse; width: 100%; font-size: 12px;">
      <thead style="background-color: #f2f2f2;">
        <tr>
          <th>Data</th>
          <th>Nome</th>
          <th>Telefono</th>
          <th>Email</th>
          <th>Luogo di Consegna</th>
          <th>Venditore Assegnato</th>
          <th>Numero pezzi</th>
          <th>Provenienza contatto</th>
        </tr>
      </thead>
      <tbody>`;

  clients.forEach((c) => {
    const dataOra = tryParseDate(c[colsMain["Data e ora"]]);
    const dataFormattata = dataOra ? dataOra.toLocaleDateString() : "-";
    const pezzi = parseInt(c[colsMain["Numero pezzi"]]) || 0;
    totalPezzi += pezzi;

    emailBody += `
      <tr>
        <td>${dataFormattata}</td>
        <td>${c[colsMain["Nome"]] || "-"}</td>
        <td>${c[colsMain["Telefono"]] || "-"}</td>
        <td>${c[colsMain["Email"]] || "-"}</td>
        <td>${c[colsMain["Luogo di Consegna"]] || "-"}</td>
        <td>${c[colsMain["Venditore Assegnato"]] || "-"}</td>
        <td style="text-align:center;">${pezzi}</td>
        <td>${c[colsMain["Provenienza contatto"]] || "Internet"}</td>
      </tr>`;
  });

  emailBody += `
      </tbody>
    </table>
    <br>
    <h3 style="text-align:right;">Totale pezzi richiesti: ${totalPezzi}</h3>
    <p style="font-size: 10px; text-align: center; margin-top: 30px;">Impaginato per stampa su foglio A4</p>
  </div>`;

  sendEmail(
    "newsaverplast@gmail.com",
    "📊 [Riepilogo settimanale] Nuovi Clienti",
    emailBody
  );

  logInfo(
    `✅ Report inviato: ${clients.length} clienti, ${totalPezzi} pezzi totali.`
  );
}

function tryParseDate(value) {
  if (value instanceof Date && !isNaN(value)) return value;

  if (typeof value === "string") {
    let clean = value
      .trim()
      .replace(",", "") // 🔹 rimuove virgole
      .replace(/\s+/g, " ") // 🔹 normalizza spazi
      .replace(/\./g, "/"); // 🔹 converte punti in slash (es. 9.4.2025 → 9/4/2025)

    // 🔍 dd/MM/yyyy HH:mm:ss
    let match = clean.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/
    );
    if (match) {
      const [_, dd, mm, yyyy, h, m, s] = match;
      return new Date(
        `${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}T${h}:${m}:${s}`
      );
    }

    // 🔍 dd/MM/yyyy
    match = clean.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      const [_, dd, mm, yyyy] = match;
      return new Date(`${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`);
    }

    // 🔍 yyyy-MM-dd HH:mm:ss
    match = clean.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);
    if (match) {
      const [_, yyyy, mm, dd, h, m, s] = match;
      return new Date(`${yyyy}-${mm}-${dd}T${h}:${m}:${s}`);
    }

    // fallback finale (se tutto fallisce)
    const parsed = new Date(clean);
    if (!isNaN(parsed)) return parsed;
  }

  return null;
}

/**
 * Gestione della coda email con retry
 */

function getEmailQueueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("EmailQueue");
  if (!sheet) {
    sheet = ss.insertSheet("EmailQueue");
    sheet.appendRow(["Email", "Oggetto", "Corpo", "Tentativi"]);
  }
  return sheet;
}

function addToEmailQueue(to, subject, body) {
  var sheet = getEmailQueueSheet();
  sheet.appendRow([to, subject, body, 0]);
  logWarning("📌 Email messa in coda per " + to);
}

function processEmailQueue() {
  var sheet = getEmailQueueSheet();
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    Logger.log("✅ Nessuna email in coda da processare.");
    return;
  }

  Logger.log(
    "⏳ Tentativo di svuotare la coda email. Email in coda: " +
      (data.length - 1)
  );

  for (var i = data.length - 1; i > 0; i--) {
    // Partiamo dal basso per rimuovere le righe senza problemi
    var row = data[i];
    var to = row[0];
    var subject = row[1];
    var body = row[2];
    var attempts = parseInt(row[3]) || 0;

    if (attempts >= 3) {
      logError("❌ Email non inviata dopo 3 tentativi: " + to);
      sheet.deleteRow(i + 1);
      continue;
    }

    try {
      MailApp.sendEmail({
        to: to,
        subject: subject,
        htmlBody: body,
      });
      logInfo("📧 Email inviata con successo a " + to);
      sheet.deleteRow(i + 1);
    } catch (e) {
      logWarning(
        "⚠️ Retry email a " + to + " (tentativo " + (attempts + 1) + ")"
      );
      sheet.getRange(i + 1, 4).setValue(attempts + 1);
    }
  }
}

/**Gestione dei Reminder **/

function sendReminderForUncontactedClients() {
  var vendors = getVendors();
  var today = new Date();
  var emailSubject = "🔔 Promemoria: Contattare i clienti assegnati!";

  for (var venditore in vendors) {
    try {
      var vendorSS = SpreadsheetApp.openById(vendors[venditore]);
      var vendorSheet = vendorSS.getSheetByName("Dati");
      if (!vendorSheet) continue;

      var vendorData = vendorSheet.getDataRange().getValues();
      var colsVendor = getColumnIndexes(vendorData[0]);

      if (!("Stato" in colsVendor) || !("Data Assegnazione" in colsVendor)) {
        logWarning(
          "⚠️ Colonne 'Stato' o 'Data Assegnazione' mancanti per " + venditore
        );
        continue;
      }

      var uncontactedClients = [];

      for (var i = 1; i < vendorData.length; i++) {
        var stato = vendorData[i][colsVendor["Stato"]].toString().trim();
        var dataAssegnazione = vendorData[i][colsVendor["Data Assegnazione"]];

        if (stato === "Da contattare" && dataAssegnazione) {
          var assignedDate = new Date(dataAssegnazione);
          var diffDays = Math.floor(
            (today - assignedDate) / (1000 * 60 * 60 * 24)
          );

          if (diffDays > 4) {
            uncontactedClients.push([
              vendorData[i][colsVendor["Nome"]],
              vendorData[i][colsVendor["Telefono"]],
              vendorData[i][colsVendor["Email"]],
              assignedDate.toLocaleDateString(),
              diffDays,
            ]);
          }
        }
      }

      if (uncontactedClients.length > 0) {
        var emailBody =
          "<b>Hai clienti in attesa di contatto da oltre 4 giorni!</b><br><br>";
        emailBody +=
          "<table border='1'><tr><th>Nome</th><th>Telefono</th><th>Email</th><th>Data Assegnazione</th><th>Giorni in attesa</th></tr>";

        uncontactedClients.forEach((client) => {
          emailBody += `<tr><td>${client[0]}</td><td>${client[1]}</td><td>${client[2]}</td><td>${client[3]}</td><td>${client[4]}</td></tr>`;
        });

        emailBody += "</table><br>Si prega di contattarli il prima possibile.";

        var vendorEmail = getVendorEmail(venditore);
        sendEmail(vendorEmail, emailSubject, emailBody);
        logInfo("📧 Reminder inviato a " + venditore);
      }
    } catch (e) {
      logError(
        "❌ Errore durante il controllo dei clienti per " +
          venditore +
          ": " +
          e.message
      );
    }
  }
}

/** Reminder dopo il 4 giorno (email tutti i giorni) **/

function sendPersistentReminders() {
  var vendors = getVendors();
  var today = new Date();
  var emailSubject = "🔔 URGENTE: Clienti in attesa di contatto!";

  for (var venditore in vendors) {
    try {
      var vendorSS = SpreadsheetApp.openById(vendors[venditore]);
      var vendorSheet = vendorSS.getSheetByName("Dati");
      if (!vendorSheet) continue;

      var vendorData = vendorSheet.getDataRange().getValues();
      var colsVendor = getColumnIndexes(vendorData[0]);

      if (!("Stato" in colsVendor) || !("Data Assegnazione" in colsVendor)) {
        logWarning(
          "⚠️ Colonne 'Stato' o 'Data Assegnazione' mancanti per " + venditore
        );
        continue;
      }

      var uncontactedClients = [];
      var urgentClients = [];

      for (var i = 1; i < vendorData.length; i++) {
        var stato = vendorData[i][colsVendor["Stato"]].toString().trim();
        var dataAssegnazione = vendorData[i][colsVendor["Data Assegnazione"]];

        if (stato === "Da contattare" && dataAssegnazione) {
          var assignedDate = new Date(dataAssegnazione);
          var diffDays = Math.floor(
            (today - assignedDate) / (1000 * 60 * 60 * 24)
          );

          if (diffDays > 4) {
            if (diffDays > 7) {
              urgentClients.push([
                vendorData[i][colsVendor["Nome"]],
                vendorData[i][colsVendor["Telefono"]],
                vendorData[i][colsVendor["Email"]],
                assignedDate.toLocaleDateString(),
                diffDays,
              ]);
            } else {
              uncontactedClients.push([
                vendorData[i][colsVendor["Nome"]],
                vendorData[i][colsVendor["Telefono"]],
                vendorData[i][colsVendor["Email"]],
                assignedDate.toLocaleDateString(),
                diffDays,
              ]);
            }
          }
        }
      }

      if (uncontactedClients.length > 0 || urgentClients.length > 0) {
        var emailBody =
          "<b>Hai clienti in attesa di contatto da oltre 4 giorni!</b><br><br>";

        if (urgentClients.length > 0) {
          emailBody +=
            "<b style='color: red;'>⚠️ ATTENZIONE! Questi clienti aspettano da più di 7 giorni:</b><br>";
          emailBody +=
            "<table border='1' style='border-collapse: collapse;'><tr><th>Nome</th><th>Telefono</th><th>Email</th><th>Data Assegnazione</th><th>Giorni in attesa</th></tr>";
          urgentClients.forEach((client) => {
            emailBody += `<tr><td>${client[0]}</td><td>${client[1]}</td><td>${client[2]}</td><td>${client[3]}</td><td style="color: red;"><b>${client[4]}</b></td></tr>`;
          });
          emailBody += "</table><br>";
        }

        if (uncontactedClients.length > 0) {
          emailBody += "<b>📌 Clienti in attesa da più di 4 giorni:</b><br>";
          emailBody +=
            "<table border='1' style='border-collapse: collapse;'><tr><th>Nome</th><th>Telefono</th><th>Email</th><th>Data Assegnazione</th><th>Giorni in attesa</th></tr>";
          uncontactedClients.forEach((client) => {
            emailBody += `<tr><td>${client[0]}</td><td>${client[1]}</td><td>${client[2]}</td><td>${client[3]}</td><td>${client[4]}</td></tr>`;
          });
          emailBody += "</table><br>";
        }

        emailBody += "<br>Si prega di contattarli il prima possibile!";

        var vendorEmail = getVendorEmail(venditore);
        sendEmail(vendorEmail, emailSubject, emailBody);
        logInfo("📧 Reminder inviato a " + venditore);
      }
    } catch (e) {
      logError(
        "❌ Errore durante il controllo dei clienti per " +
          venditore +
          ": " +
          e.message
      );
    }
  }
}

/** Dashboard **/

function updateDashboardFromMain() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  sheet.clear();

  const mainSheet = ss.getSheetByName("Main");
  const data = mainSheet.getDataRange().getValues();
  const headers = data[0];
  const cols = getColumnIndexes(headers);

  const today = new Date();
  const venditori = {};
  const statoDistribuzione = {};
  const provenienze = {};

  let totaleClienti = 0,
    preventivi = 0,
    trattative = 0,
    vendite = 0,
    daContattare = 0;
  let tempoTotaleRisposta = 0,
    risposteValide = 0,
    totalePezzi = 0;
  let settimanali = 0,
    settimanaliPreventivi = 0,
    settimanaliVendite = 0,
    settimanaliPezzi = 0;

  const settimanaCorrente = getWeekNumber(today);
  const lunediScorso = getLastMonday();

  const normalizzaProvenienza = (prov) => {
    prov = prov.toLowerCase().trim();
    if (prov.includes("cagliari")) return "Showroom Cagliari";
    if (prov.includes("macchiareddu")) return "Showroom Macchiareddu";
    if (prov.includes("nuoro")) return "Showroom Nuoro";
    if (prov.includes("google")) return "Google";
    if (prov.includes("facebook")) return "Facebook";
    if (prov.includes("instagram")) return "Instagram";
    if (prov.includes("passaparola")) return "Passaparola";
    return prov.charAt(0).toUpperCase() + prov.slice(1);
  };

  function parseCustomDate(val) {
    if (val instanceof Date) return new Date(val.setHours(0, 0, 0, 0));
    if (typeof val === "string") {
      const match = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (match) {
        const [_, d, m, y] = match;
        return new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
      } else {
        const parsed = new Date(val);
        if (!isNaN(parsed)) return new Date(parsed.setHours(0, 0, 0, 0));
      }
    }
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let stato = (row[cols["Stato"]] || "").toString().trim();
    const venditore = (row[cols["Venditore Assegnato"]] || "")
      .toString()
      .trim();
    const dataAssegnazione = parseCustomDate(row[cols["Data e ora"]]);
    const dataPreventivo = parseCustomDate(row[cols["Data Preventivo"]]);
    const venditaConclusa = (row[cols["Vendita Conclusa?"]] || "")
      .toString()
      .trim()
      .toUpperCase();
    const pezzi = parseInt(row[cols["Numero pezzi"]]) || 0;
    const provenienzaRaw = (row[cols["Provenienza contatto"]] || "Internet")
      .toString()
      .trim();
    const provenienza = normalizzaProvenienza(provenienzaRaw);

    if (!venditore) continue;
    if (!venditori[venditore])
      venditori[venditore] = {
        clienti: 0,
        daContattare: 0,
        preventivi: 0,
        trattative: 0,
        vendite: 0,
        tempi: [],
      };

    venditori[venditore].clienti++;
    if (!stato) stato = "Non specificato";
    statoDistribuzione[stato] = (statoDistribuzione[stato] || 0) + 1;
    provenienze[provenienza] = (provenienze[provenienza] || 0) + 1;
    totaleClienti++;
    totalePezzi += pezzi;

    if (stato === "Da contattare") {
      venditori[venditore].daContattare++;
      daContattare++;
    } else if (stato === "Preventivo inviato") {
      venditori[venditore].preventivi++;
      preventivi++;
    } else if (stato === "In trattativa") {
      venditori[venditore].trattative++;
      trattative++;
    }

    if (
      venditaConclusa === "SI" ||
      (venditaConclusa === "" && stato === "Trattativa terminata")
    ) {
      venditori[venditore].vendite++;
      vendite++;
    }

    if (stato === "Preventivo inviato" && dataAssegnazione && dataPreventivo) {
      const diffMillis = dataPreventivo.getTime() - dataAssegnazione.getTime();
      const giorni = Math.round(diffMillis / (1000 * 60 * 60 * 24));

      // Considera solo risposte tra 0 e 30 giorni per evitare outlier
      if (!isNaN(giorni) && giorni >= 0 && giorni <= 30) {
        venditori[venditore].tempi.push(giorni);
        tempoTotaleRisposta += giorni;
        risposteValide++;
      } else {
        Logger.log(
          `⚠️ Giorni anomali (esclusi): ${giorni} – Riga ${
            i + 1
          } | Venditore: ${venditore}`
        );
      }
    }

    if (dataAssegnazione && dataAssegnazione >= lunediScorso) {
      settimanali++;
      settimanaliPezzi += pezzi;
      if (stato === "Preventivo inviato") settimanaliPreventivi++;
      if (
        venditaConclusa === "SI" ||
        (venditaConclusa === "" && stato === "Trattativa terminata")
      )
        settimanaliVendite++;
    }
  }

  sheet
    .getRange("B1")
    .setValue("📊 DASHBOARD DATI CLIENTI (fonte: Main)")
    .setFontSize(16)
    .setFontWeight("bold");
  sheet
    .getRange("B2")
    .setValue("Aggiornata al: " + today.toLocaleString())
    .setFontStyle("italic");
  sheet.appendRow([""]);

  sheet.appendRow([
    "Totale clienti",
    "Preventivi inviati",
    "Vendite concluse",
    "Tempo medio risposta (gg)",
    "Totale pezzi richiesti",
  ]);
  sheet.appendRow([
    totaleClienti,
    preventivi,
    vendite,
    risposteValide ? Math.round(tempoTotaleRisposta / risposteValide) : "-",
    totalePezzi,
  ]);
  sheet
    .getRange("B4:F4")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#1f4e78")
    .setFontColor("white");
  sheet
    .getRange("B5:F5")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#d9d9d9")
    .setFontColor("black");
  sheet.appendRow([""]);

  sheet.appendRow([
    "Venditore",
    "Totale",
    "Da contattare",
    "Preventivi",
    "Trattative",
    "Vendite",
    "Risposta media (gg)",
    "% Chiusura",
  ]);

  Object.keys(venditori).forEach((nome) => {
    const v = venditori[nome];
    const media = v.tempi.length
      ? Math.round(v.tempi.reduce((a, b) => a + b, 0) / v.tempi.length)
      : "-";
    const chiusura =
      v.preventivi > 0
        ? ((v.vendite / v.preventivi) * 100).toFixed(1) + "%"
        : "-";
    sheet.appendRow([
      nome,
      v.clienti,
      v.daContattare,
      v.preventivi,
      v.trattative,
      v.vendite,
      media,
      chiusura,
    ]);
  });
  const headerStartRow = sheet.getLastRow() - Object.keys(venditori).length - 1;
  sheet
    .getRange(headerStartRow, 2, 1, 8)
    .setFontWeight("bold")
    .setBackground("#f2f2f2")
    .setHorizontalAlignment("center");

  sheet.appendRow([""]);
  sheet.appendRow([`📅 Riepilogo settimana #${getWeekNumber(lunediScorso)}`]);
  sheet.appendRow([
    "Nuovi clienti",
    "Pezzi richiesti",
    "Preventivi",
    "Vendite",
  ]);
  sheet.appendRow([
    settimanali,
    settimanaliPezzi,
    settimanaliPreventivi,
    settimanaliVendite,
  ]);
  sheet
    .getRange(sheet.getLastRow() - 1, 2, 2, 4)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  const startRow = sheet.getLastRow() + 2;

  const statoKeys = Object.keys(statoDistribuzione);
  sheet.getRange(startRow, 2).setValue("Distribuzione per Stato");
  statoKeys.forEach((k, i) => {
    sheet.getRange(startRow + i + 1, 2).setValue(k);
    sheet.getRange(startRow + i + 1, 3).setValue(statoDistribuzione[k]);
  });
  const chart1 = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(startRow + 1, 2, statoKeys.length, 2))
    .setPosition(startRow, 5, 0, 0)
    .setOption("title", "Distribuzione per Stato")
    .build();
  sheet.insertChart(chart1);

  const provKeys = Object.keys(provenienze);
  const provStart = startRow + statoKeys.length + 12;
  sheet.getRange(provStart, 2).setValue("Distribuzione Provenienza Contatto");
  provKeys.forEach((k, i) => {
    sheet.getRange(provStart + i + 1, 2).setValue(k);
    sheet.getRange(provStart + i + 1, 3).setValue(provenienze[k]);
  });
  const chart2 = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(provStart + 1, 2, provKeys.length, 2))
    .setPosition(provStart, 5, 0, 0)
    .setOption("title", "Provenienza Contatti")
    .build();
  sheet.insertChart(chart2);

  logInfo("✅ Dashboard aggiornata da 'Main'");
}

function debugVendorSheets() {
  var vendors = getVendors();
  for (var venditore in vendors) {
    try {
      var vendorSS = SpreadsheetApp.openById(vendors[venditore]);
      var vendorSheet = vendorSS.getSheetByName("Dati");
      if (!vendorSheet) {
        Logger.log("❌ Il foglio 'Dati' non esiste per " + venditore);
        continue;
      }

      var vendorData = vendorSheet.getDataRange().getValues();
      Logger.log(
        "📌 Dati trovati per " + venditore + ": " + vendorData.length + " righe"
      );

      if (vendorData.length < 2) {
        Logger.log("⚠️ Il foglio di " + venditore + " non contiene dati!");
      }
    } catch (e) {
      Logger.log("❌ Errore con il venditore " + venditore + ": " + e.message);
    }
  }
}

/** Avvio programma**/

function avviaProgramma() {
  var emailDestinatario = "newsaverplast@gmail.com"; // Indirizzo email per le notifiche
  var errori = [];

  try {
    syncMainToVendors(); // Sincronizza il foglio "Main" con i venditori
  } catch (e) {
    var errore1 = "Errore in syncMainToVendors(): " + e.message;
    errori.push(errore1);
    logError(errore1);
  }

  try {
    updateMainFromVendors(); // Aggiorna "Main" con i dati dei venditori
  } catch (e) {
    var errore2 = "Errore in updateMainFromVendors(): " + e.message;
    errori.push(errore2);
    logError(errore2);
  }

  // Se ci sono errori, invia un'email di avviso
  if (errori.length > 0) {
    MailApp.sendEmail({
      to: emailDestinatario,
      subject: "⚠️ Errore nell'esecuzione del programma",
      body: "Si sono verificati i seguenti errori:\n\n" + errori.join("\n"),
    });
  }
}

// 📌 Funzione per registrare gli errori nel foglio "Log"
function logError(messaggio) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log") || ss.insertSheet("Log");
  logSheet.appendRow([new Date().toLocaleString(), "Errore", messaggio]);
}

/**Intervento ai per numero pezzi */

function getNumeroPezziConOpenAI(prompt) {
  const apiKey = getOpenAIKey();
  if (!apiKey) {
    Logger.log("❌ API Key non trovata!");
    return 0;
  }

  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content:
          "Sei un assistente che deve SOLO contare il numero totale di oggetti richiesti. Rispondi SEMPRE e SOLO con un numero intero, senza testo aggiuntivo, se non capisci scrivi 0.",
      },
      {
        role: "user",
        content: prompt,
      },
    ],
    max_tokens: 10,
    temperature: 0,
  };

  const options = {
    method: "post",
    headers: {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const text = response.getContentText();
    Logger.log("🔎 Risposta OpenAI: " + text);
    const json = JSON.parse(text);

    if (json.error) {
      Logger.log("❌ Errore OpenAI: " + json.error.message);
      return 0;
    }

    if (json.choices && json.choices.length > 0) {
      const content = json.choices[0].message.content.trim();
      Logger.log("👉 Contenuto: " + content);
      const numero = parseInt(content, 10);
      return isNaN(numero) ? 0 : numero;
    } else {
      return 0;
    }
  } catch (err) {
    Logger.log("❌ Errore fetch: " + err);
    return 0;
  }
}

function testEstrazionePezzi() {
  const messaggio = "Vorrei 3 finestre, una porta finestra e due persiane.";
  const numero = getNumeroPezziConOpenAI(messaggio);
  Logger.log("🧪 Numero rilevato: " + numero);
}

function aggiornaNumeroPezziInMain() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  if (!mainSheet) {
    logError("❌ Foglio 'Main' non trovato.");
    return;
  }

  addMultipleColumnsToMain(mainSheet, ["Numero pezzi", "Provenienza contatto"]);

  var data = mainSheet.getDataRange().getValues();
  var headers = data[0];
  var cols = getColumnIndexes(headers);
  var numAggiornati = 0;
  var provenienzeAggiornate = 0;

  for (var i = 1; i < data.length; i++) {
    var messaggio = data[i][cols["Messaggio"]] || "";
    var numeroPezzi = data[i][cols["Numero pezzi"]];
    var provenienza = data[i][cols["Provenienza contatto"]];
    var numeroValido = parseInt(numeroPezzi);

    // ✅ aggiorna "Numero pezzi" solo se la cella è davvero vuota
    const isBlank =
      numeroPezzi === "" ||
      numeroPezzi === null ||
      typeof numeroPezzi === "undefined";
    if (isBlank) {
      var valoreEstratto = getNumeroPezziConOpenAI(messaggio);
      mainSheet
        .getRange(i + 1, cols["Numero pezzi"] + 1)
        .setValue(valoreEstratto);
      numAggiornati++;
    }

    // ✅ aggiorna "Provenienza contatto" solo se vuota
    if (!provenienza || provenienza.toString().trim() === "") {
      mainSheet
        .getRange(i + 1, cols["Provenienza contatto"] + 1)
        .setValue("Internet");
      provenienzeAggiornate++;
    }
  }

  logInfo(
    `✅ ${numAggiornati} 'Numero pezzi' aggiornati, ${provenienzeAggiornate} 'Provenienza contatto' impostate su 'Internet'.`
  );
}

/**
 * 🔹 Deduplica nel foglio "Main" usando chiave Nome|Telefono|Email
 *    - Tiene la PRIMA occorrenza, rimuove le successive
 *    - Esegue un backup del foglio "Main" prima di intervenire (riusa createBackup)
 */
function dedupMainOnce() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Main");
  if (!sh) {
    Logger.log("❌ 'Main' non trovato.");
    return;
  }

  // Backup del foglio Main (usa la tua funzione esistente)
  try {
    createBackup(sh);
  } catch (e) {
    Logger.log("⚠️ Impossibile creare backup di 'Main': " + e.message);
  }

  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) {
    Logger.log("ℹ️ Nessun dato da deduplicare in 'Main'.");
    return;
  }

  var cols = getColumnIndexes(data[0]);
  if (!("Nome" in cols) || !("Telefono" in cols) || !("Email" in cols)) {
    Logger.log(
      "⚠️ 'Main' non ha tutte le colonne richieste (Nome, Telefono, Email)."
    );
    return;
  }

  var seen = new Set();
  var toDelete = [];

  for (var r = 1; r < data.length; r++) {
    var nome = (data[r][cols["Nome"]] || "").toString().trim().toLowerCase();
    var tel = (data[r][cols["Telefono"]] || "").toString().trim();
    var mail = (data[r][cols["Email"]] || "").toString().trim().toLowerCase();

    // salta righe completamente vuote
    if (!nome && !tel && !mail) continue;

    var key = nome + "|" + tel + "|" + mail;
    if (seen.has(key)) {
      toDelete.push(r + 1); // indice 1-based per deleteRow
    } else {
      seen.add(key);
    }
  }

  // elimina dal basso verso l’alto
  for (var i = toDelete.length - 1; i >= 0; i--) {
    sh.deleteRow(toDelete[i]);
  }

  Logger.log(
    `🧹 Main: rimossi ${toDelete.length} duplicati (criterio Nome|Telefono|Email).`
  );
}

/**
 * 🔹 Deduplica in TUTTI i fogli "Dati" dei venditori (getVendors())
 *    - Chiave Nome|Telefono
 *    - Tiene la PRIMA occorrenza, rimuove le successive
 *    - Crea una copia di backup del foglio "Dati" nel file del venditore
 */
function dedupVendorsOnce() {
  var vendors = getVendors();
  var totalDeleted = 0;

  Object.keys(vendors).forEach((venditore) => {
    try {
      var vSS = SpreadsheetApp.openById(vendors[venditore]);
      var sh = vSS.getSheetByName("Dati");
      if (!sh) {
        Logger.log(`ℹ️ ${venditore}: foglio 'Dati' non trovato, salto.`);
        return;
      }

      // Backup del foglio "Dati" nel file del venditore
      try {
        var backupName =
          "Dati_backup_" +
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "yyyyMMdd_HHmmss"
          );
        var copied = sh.copyTo(vSS);
        copied.setName(backupName);
      } catch (e) {
        Logger.log(
          `⚠️ ${venditore}: impossibile creare backup del foglio Dati → ${e.message}`
        );
      }

      var data = sh.getDataRange().getValues();
      if (!data || data.length < 2) {
        Logger.log(`ℹ️ ${venditore}: nessun dato da deduplicare.`);
        return;
      }

      var cols = getColumnIndexes(data[0]);
      if (!("Nome" in cols) || !("Telefono" in cols)) {
        Logger.log(`⚠️ ${venditore}: mancano colonne 'Nome' o 'Telefono'.`);
        return;
      }

      var seen = new Set();
      var toDelete = [];

      for (var r = 1; r < data.length; r++) {
        var nome = (data[r][cols["Nome"]] || "")
          .toString()
          .trim()
          .toLowerCase();
        var tel = (data[r][cols["Telefono"]] || "").toString().trim();

        if (!nome && !tel) continue; // riga “vuota”
        var key = nome + "|" + tel;

        if (seen.has(key)) {
          toDelete.push(r + 1); // 1-based
        } else {
          seen.add(key);
        }
      }

      // elimina bottom-up
      for (var i = toDelete.length - 1; i >= 0; i--) {
        sh.deleteRow(toDelete[i]);
      }

      Logger.log(
        `🧹 ${venditore}: rimossi ${toDelete.length} duplicati (criterio Nome|Telefono).`
      );
      totalDeleted += toDelete.length;
    } catch (e) {
      Logger.log(`❌ Errore dedup per ${venditore}: ${e.message}`);
    }
  });

  Logger.log(
    `✅ Dedup venditori completata. Totale duplicati rimossi: ${totalDeleted}.`
  );
}

/**
 * 🔸 Esegue entrambe le dedupliche (Main + tutti i venditori)
 */
function dedupAllContacts() {
  dedupMainOnce();
  dedupVendorsOnce();
  Logger.log("✅ Dedup completa su Main e fogli venditori.");
}

function dedupEmailInRichiestaPreventivo_DELETE_DUPLICATES() {
  const LABEL_NAME = "Richiesta Preventivo - infissipvcsardegna";
  const BATCH_SIZE = 100; // quanti thread per batch
  const DRY_RUN = false; // true = non cancella davvero, solo log

  const label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    Logger.log("❌ Etichetta non trovata: " + LABEL_NAME);
    return;
  }

  // Conta thread (può essere pesante su etichette enormi, ma per 400-1000 ok)
  const totalThreads = label.getThreads().length;
  Logger.log(
    "📂 Etichetta: " + LABEL_NAME + " – Thread totali: " + totalThreads
  );

  // Mappa: key -> array di {id, date, msg}
  const groups = new Map();
  let processedMsgs = 0;

  for (let start = 0; start < totalThreads; start += BATCH_SIZE) {
    const threads = label.getThreads(
      start,
      Math.min(BATCH_SIZE, totalThreads - start)
    );

    threads.forEach((thread) => {
      // NB: l’etichetta è sul thread; qui scorriamo tutti i messaggi del thread
      const msgs = thread.getMessages();

      msgs.forEach((msg) => {
        const from = (msg.getFrom() || "").toLowerCase().trim();
        const subject = (msg.getSubject() || "").toLowerCase().trim();

        // Corpo normalizzato per evitare falsi positivi/negativi
        let body = msg.getPlainBody() || msg.getBody() || "";
        body = normalizeBody_(body);

        // Chiave robusta: From + Subject + hash(corpo)
        const bodyHash = md5Hex_(body);
        let key = from + "|" + subject + "|" + bodyHash;

        // Fallback: se il corpo è troppo corto, usa From+Subject+giorno
        if (!body || body.length < 10) {
          const day = Utilities.formatDate(
            msg.getDate(),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          );
          key = from + "|" + subject + "|" + day;
        }

        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push({
          id: msg.getId(),
          date: msg.getDate(),
          msg: msg,
        });

        processedMsgs++;
      });
    });

    Utilities.sleep(200); // piccola pausa anti-quota
  }

  // Per ogni gruppo: tieni il più vecchio, elimina gli altri
  let deleteCount = 0,
    groupCount = 0,
    keepCount = 0;

  groups.forEach((arr) => {
    if (arr.length <= 1) return; // nessun duplicato

    groupCount++;
    // ordina per data asc (più vecchio per primo)
    arr.sort((a, b) => a.date - b.date);

    // tieni il primo
    keepCount++;

    // elimina (sposta nel cestino) i successivi
    for (let i = 1; i < arr.length; i++) {
      const m = arr[i].msg;
      if (!DRY_RUN) {
        try {
          m.moveToTrash(); // 👈 elimina SOLO quel messaggio (non l’intero thread)
          deleteCount++;
        } catch (e) {
          Logger.log(
            "⚠️ Errore nel cestinare msg " + m.getId() + ": " + e.message
          );
        }
      } else {
        Logger.log(
          "🧪 DRY_RUN: doppione da eliminare: " +
            m.getDate() +
            " | " +
            m.getFrom() +
            " | " +
            m.getSubject()
        );
      }
    }
  });

  Logger.log("✅ Analizzati messaggi: " + processedMsgs);
  Logger.log("🧩 Gruppi con duplicati: " + groupCount);
  Logger.log(
    (DRY_RUN ? "🧪 [DRY RUN] " : "") +
      "🗑️ Messaggi duplicati spostati nel Cestino: " +
      deleteCount
  );
}

/** Normalizza corpo: spazi/rientri, righe vuote, firme semplici… */
function normalizeBody_(text) {
  let t = text
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{2,}/g, "\n")
    .trim();

  // (opzionale) rimuovi footer tipo firma dopo "--"
  // t = t.replace(/--\s*\n.*$/s, "");

  return t;
}

/** MD5 hex di una stringa (per fingerprint contenuti) */
function md5Hex_(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    str,
    Utilities.Charset.UTF_8
  );
  let hex = "";
  for (let i = 0; i < bytes.length; i++) {
    let h = (bytes[i] & 0xff).toString(16);
    if (h.length === 1) h = "0" + h;
    hex += h;
  }
  return hex;
}

function reconcileGmailWithMain() {
  const LABEL_NAME = "Richiesta Preventivo - infissipvcsardegna";
  const EXTRA_LABEL = "Non in Main";
  const BATCH_SIZE = 100;

  const DRY_RUN = false; // 👉 PRIMA PROVA COSÌ
  const DELETE_EXTRAS = true; // se false: etichetta gli "extra"; se true: li cestina

  // === 1) Costruisci gli indici da MAIN ===
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Main");
  if (!sh) {
    Logger.log("❌ 'Main' non trovato");
    return;
  }

  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("ℹ️ 'Main' senza righe utili");
    return;
  }

  const cols = getColumnIndexes(data[0]);
  if (!("Email" in cols) || !("Messaggio" in cols) || !("Telefono" in cols)) {
    Logger.log("❌ 'Main' deve avere 'Email', 'Messaggio', 'Telefono'");
    return;
  }

  // Indici:
  // - Set di chiavi email|hash(body)
  const mainKeyExact = new Set();
  // - Mappa telefono -> array entries
  const mainByPhone = new Map();
  // - Mappa email -> array entries (per similarità)
  const mainByEmail = new Map();

  // entry = { email, phone, bodyNorm, bodyHash, keyExact }
  for (let r = 1; r < data.length; r++) {
    const email = (data[r][cols["Email"]] || "")
      .toString()
      .trim()
      .toLowerCase();
    const phone = normalizePhone_((data[r][cols["Telefono"]] || "").toString());
    const body = (data[r][cols["Messaggio"]] || "").toString();
    const bodyNorm = normalizeTextForCompare_(body);
    const bodyHash = md5Hex_(bodyNorm);
    const keyExact = email + "|" + bodyHash;

    if (email || bodyNorm) mainKeyExact.add(keyExact);

    if (phone) {
      if (!mainByPhone.has(phone)) mainByPhone.set(phone, []);
      mainByPhone
        .get(phone)
        .push({ email, phone, bodyNorm, bodyHash, keyExact });
    }

    if (email) {
      if (!mainByEmail.has(email)) mainByEmail.set(email, []);
      mainByEmail
        .get(email)
        .push({ email, phone, bodyNorm, bodyHash, keyExact });
    }
  }

  Logger.log(
    "📊 Main: chiavi esatte: " +
      mainKeyExact.size +
      ", telefoni distinti: " +
      mainByPhone.size +
      ", email distinte: " +
      mainByEmail.size
  );

  // === 2) Scorri GMAIL ===
  const label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    Logger.log("❌ Etichetta Gmail non trovata: " + LABEL_NAME);
    return;
  }

  let extraLabel = GmailApp.getUserLabelByName(EXTRA_LABEL);
  if (!extraLabel) extraLabel = GmailApp.createLabel(EXTRA_LABEL);

  const totalThreads = label.getThreads().length;
  Logger.log("📂 Thread totali: " + totalThreads);

  let processed = 0,
    matched = 0,
    extras = 0,
    removed = 0;
  // Per report Main non trovati (opzionale)
  const foundExactOrPhoneOrSimilar = new Set();

  for (let start = 0; start < totalThreads; start += BATCH_SIZE) {
    const threads = label.getThreads(
      start,
      Math.min(BATCH_SIZE, totalThreads - start)
    );

    threads.forEach((thread) => {
      const msgs = thread.getMessages();
      msgs.forEach((msg) => {
        processed++;

        const bodyPlain = msg.getPlainBody() || msg.getBody() || "";
        const normBody = normalizeTextForCompare_(bodyPlain);
        const bodyHash = md5Hex_(normBody);

        // email cliente: prova body, poi header From
        const emailFromBody = extractEmailFromText_(bodyPlain) || "";
        const fromHeader = msg.getFrom() || "";
        const fromEmail = extractEmailFromText_(fromHeader) || "";
        const email = (emailFromBody || fromEmail).toLowerCase();

        // telefono: prova dal body
        const phoneList = extractPhones_(bodyPlain);
        const phoneNorm = phoneList.length ? normalizePhone_(phoneList[0]) : "";

        // 2.1 Match esatto: email + hash(body)
        let isMatch = false;
        if (email && mainKeyExact.has(email + "|" + bodyHash)) {
          isMatch = true;
          foundExactOrPhoneOrSimilar.add(email + "|" + bodyHash);
        }

        // 2.2 Match per telefono
        if (!isMatch && phoneNorm && mainByPhone.has(phoneNorm)) {
          isMatch = true;
          // Non abbiamo la keyExact sicura, ma segnalare che c'è match
          // (Se vuoi, potresti registrare phone-only)
        }

        // 2.3 Match per similarità con stessa email
        if (!isMatch && email && mainByEmail.has(email)) {
          const candidates = mainByEmail.get(email);
          // Similarità = contenimento reciproco minimo
          const sim = candidates.some((ent) =>
            isSimilarText_(normBody, ent.bodyNorm)
          );
          if (sim) {
            isMatch = true;
          }
        }

        if (isMatch) {
          matched++;
          return; // non toccare il messaggio
        }

        // Se non ha match con nessuna strategia → extra
        extras++;
        if (DRY_RUN) {
          Logger.log(
            "🧪 Extra: " +
              msg.getDate() +
              " | email=" +
              email +
              " | tel=" +
              phoneNorm +
              " | subj=" +
              (msg.getSubject() || "")
          );
        } else {
          if (DELETE_EXTRAS) {
            try {
              msg.moveToTrash();
              removed++;
            } catch (e) {
              Logger.log(
                "⚠️ Errore cancellando msg " + msg.getId() + ": " + e.message
              );
            }
          } else {
            try {
              thread.addLabel(extraLabel);
            } catch (e) {
              Logger.log("⚠️ Errore etichettando thread: " + e.message);
            }
          }
        }
      });
    });

    Utilities.sleep(200);
  }

  Logger.log("✅ Messaggi esaminati: " + processed);
  Logger.log("✔️ Match trovati: " + matched);
  Logger.log(
    (DRY_RUN ? "🧪 [DRY RUN] " : "") +
      (DELETE_EXTRAS
        ? "🗑️ Extra cestinati: " + removed
        : "🏷️ Extra etichettati: " + extras)
  );
  if (DRY_RUN && extras === 0 && matched === 0) {
    Logger.log(
      "ℹ️ Se hai ancora 0 match, rivediamo la strategia di estrazione (formati atipici)."
    );
  }
}

/* =========================
   Helper per confronti
   ========================= */

function normalizeTextForCompare_(text) {
  return (text || "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{2,}/g, "\n")
    .replace(/--\s*\n.*$/s, "") // rimuovi firma semplice dopo “--”
    .trim()
    .toLowerCase();
}

function extractEmailFromText_(t) {
  if (!t) return "";
  const m = t.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0] : "";
}

function extractPhones_(t) {
  if (!t) return [];
  // Estrai numeri italiani plausibili (togli spazi/punti; accetta +39 opzionale)
  const cleaned = t.replace(/[\.\-\(\)]/g, " ");
  const digits = cleaned.match(/(\+?\s*39)?\s*\d{6,}/g); // >= 6 cifre utili
  if (!digits) return [];
  return digits.map((x) => x.replace(/\D+/g, ""));
}

function normalizePhone_(p) {
  if (!p) return "";
  let digits = p.replace(/\D+/g, "");
  // rimuovi prefisso 39 ripetuto
  if (digits.startsWith("39") && digits.length > 10)
    digits = digits.slice(digits.length - 10);
  return digits;
}

function md5Hex_(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    str,
    Utilities.Charset.UTF_8
  );
  let hex = "";
  for (let i = 0; i < bytes.length; i++) {
    let h = (bytes[i] & 0xff).toString(16);
    if (h.length === 1) h = "0" + h;
    hex += h;
  }
  return hex;
}

function isSimilarText_(a, b) {
  // similarità “leggera”: uno contiene l’altro per almeno 30 caratteri
  if (!a || !b) return false;
  if (a === b) return true;
  const minLen = 30;
  return (
    (a.length >= minLen && b.includes(a.slice(0, minLen))) ||
    (b.length >= minLen && a.includes(b.slice(0, minLen)))
  );
}

/**
 * Elenca le righe di "Main" che NON hanno alcuna email corrispondente
 * nella/e label Gmail Richiesta(e) Preventivo - infissipvcsardegna.
 * Corrispondenza per: Email oppure Telefono (normalizzato).
 * Output: foglio "Main_non_in_Gmail".
 */
function reportMainNotInGmail() {
  const LABEL_CANDIDATES = [
    "Richiesta Preventivo - infissipvcsardegna",
    "Richieste Preventivo - infissipvcsardegna",
  ];

  // 👉 opzionale: limita l'indicizzazione Gmail agli ultimi N giorni (commenta per tutto lo storico)
  // const DATE_FROM_DAYS = 90;
  // const dateFrom = new Date(Date.now() - DATE_FROM_DAYS*24*60*60*1000);

  // ========== 1) INDICI DA GMAIL ==========
  const gmailIndex = buildGmailIndex_(LABEL_CANDIDATES /*, dateFrom*/);
  const emailsInLabel = gmailIndex.emails; // Set<string>
  const phonesInLabel = gmailIndex.phones; // Set<string>

  // ========== 2) LEGGI "Main" ==========
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Main");
  if (!sh) {
    Logger.log("❌ 'Main' non trovato.");
    return;
  }
  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("ℹ️ 'Main' vuoto.");
    return;
  }
  const cols = getColumnIndexes(data[0]);
  const required = [
    "Data e ora",
    "Nome",
    "Telefono",
    "Email",
    "Provincia",
    "Luogo di Consegna",
    "Messaggio",
    "Venditore Assegnato",
    "Stato",
  ];
  const missing = required.filter((k) => !(k in cols));
  if (missing.length) {
    Logger.log("⚠️ In 'Main' mancano colonne: " + missing.join(", "));
  }

  // ========== 3) COSTRUISCI RISULTATI ==========
  const results = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];

    const email = (row[cols["Email"]] || "").toString().trim().toLowerCase();
    const tel = normalizePhone_((row[cols["Telefono"]] || "").toString());
    const nome = (row[cols["Nome"]] || "").toString().trim();
    const emptyRow = !email && !tel && !nome;
    if (emptyRow) continue;

    const hasEmail = email && emailsInLabel.has(email);
    const hasPhone = tel && phonesInLabel.has(tel);

    if (!hasEmail && !hasPhone) {
      results.push([
        row[cols["Data e ora"]] || "",
        nome,
        row[cols["Telefono"]] || "",
        row[cols["Email"]] || "",
        row[cols["Provincia"]] || "",
        row[cols["Luogo di Consegna"]] || "",
        row[cols["Messaggio"]] || "",
        cols["Venditore Assegnato"] !== undefined
          ? row[cols["Venditore Assegnato"]]
          : "",
        cols["Stato"] !== undefined ? row[cols["Stato"]] : "",
      ]);
    }
  }

  // ========== 4) SCRIVI FOGLIO OUTPUT ==========
  const outName = "Main_non_in_Gmail";
  const out = ss.getSheetByName(outName) || ss.insertSheet(outName);
  out.clear();
  const header = [
    "Data e ora",
    "Nome",
    "Telefono",
    "Email",
    "Provincia",
    "Luogo di Consegna",
    "Messaggio",
    "Venditore Assegnato",
    "Stato",
  ];
  out.getRange(1, 1, 1, header.length).setValues([header]);
  if (results.length) {
    out.getRange(2, 1, results.length, header.length).setValues(results);
  }
  // un po' di formattazione utile
  out
    .getRange(1, 1, 1, header.length)
    .setFontWeight("bold")
    .setBackground("#f2f2f2");
  out.setFrozenRows(1);
  autoResizeAllColumns_(out);

  Logger.log(`✅ Nominativi in Main ma NON in Gmail: ${results.length}`);
}

/**
 * Indicizza le EMAIL e i TELEFONI presenti nei messaggi sotto una o più label.
 * Ritorna { emails:Set<string>, phones:Set<string> }.
 */
function buildGmailIndex_(labelNames /*, dateFrom*/) {
  const emails = new Set();
  const phones = new Set();

  // trova la prima label esistente tra i candidati (o indicizza tutte quelle presenti)
  const labels = [];
  labelNames.forEach((name) => {
    const l = GmailApp.getUserLabelByName(name);
    if (l) labels.push(l);
  });

  if (!labels.length) {
    Logger.log(
      "⚠️ Nessuna delle label specificate esiste in Gmail: " +
        labelNames.join(" | ")
    );
    return { emails, phones };
  }

  // Scorri TUTTI i thread di ciascuna label
  labels.forEach((label) => {
    const totalThreads = label.getThreads().length;
    const BATCH = 100;

    Logger.log(
      `📂 Indicizzo label "${label.getName()}": thread=${totalThreads}`
    );

    for (let start = 0; start < totalThreads; start += BATCH) {
      const threads = label.getThreads(
        start,
        Math.min(BATCH, totalThreads - start)
      );
      threads.forEach((thread) => {
        // opzionale: filtro per data (se servisse)
        // if (dateFrom && thread.getLastMessageDate() < dateFrom) return;

        const msgs = thread.getMessages();
        msgs.forEach((msg) => {
          // email dal From (o dal corpo come fallback)
          const fromHeader = msg.getFrom() || "";
          const bodyPlain = msg.getPlainBody() || msg.getBody() || "";

          const emailFromHeader = extractEmailFromText_(fromHeader) || "";
          const emailFromBody = extractEmailFromText_(bodyPlain) || "";
          const email = (emailFromHeader || emailFromBody).toLowerCase();
          if (email) emails.add(email);

          // telefoni plausibili dal corpo
          const phonesList = extractPhones_(bodyPlain);
          phonesList.forEach((p) => {
            const norm = normalizePhone_(p);
            if (norm) phones.add(norm);
          });
        });
      });

      Utilities.sleep(200); // anti-quota
    }
  });

  Logger.log(
    `📊 Indice Gmail — emails: ${emails.size}, phones: ${phones.size}`
  );
  return { emails, phones };
}

/** Auto-fit colonne di un foglio */
function autoResizeAllColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  for (let c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
  }
}

/** Trigger**/

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

/** Resetta il conto delle righe **/

function resetLastProcessedRow() {
  PropertiesService.getScriptProperties().deleteProperty("lastProcessedRow");
  Logger.log("🔄 Ultima riga elaborata resettata!");
}

/** Email trigger **/

function setupEmailProcessingTrigger() {
  ScriptApp.newTrigger("processEmailQueue")
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log("✅ Trigger per svuotare la coda email creato.");
}

/** Reminder trigger dopo 4 giorni **/

function setupReminderTrigger() {
  ScriptApp.newTrigger("sendReminderForUncontactedClients")
    .timeBased()
    .everyDays(1)
    .atHour(9) // Invia l'email ogni giorno alle 9:00
    .create();
  Logger.log("✅ Trigger per il promemoria venditori creato.");
}

/** Reminder trigger dopo 4 giorni + 1 giorno (mail ogni giorno) **/

function setupDailyReminderTrigger() {
  ScriptApp.newTrigger("sendPersistentReminders")
    .timeBased()
    .everyDays(1)
    .atHour(9) // Invia l'email ogni giorno alle 9:00
    .create();
  Logger.log("✅ Trigger per il promemoria giornaliero creato.");
}

/** Dashboard Trigger */

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

/** Riepilogo settimanale Trigger */

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

/**Triggher avvio programma */

function setupProgramTrigger() {
  ScriptApp.newTrigger("avviaProgramma")
    .timeBased()
    .everyMinutes(10) // Esegue ogni 10 minuti (puoi personalizzarlo)
    .create();
}
