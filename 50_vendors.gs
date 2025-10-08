/** Vendors sync & helpers
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */


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


function getVendors() {
  return {
    "Mircko Manconi": "1mGFlFbCYy9ylVjNA9l6b855c6jlIDr6QOua2qfSjckw",
    "Cristian Piga": "1N0_GKbJvZLQbKKIgfVYW4LQGp97mhQcOz9zsD-FBNcE",
    "Marco Guidi": "1CVQSnFGNX8pGUKUABdtzwQmyCKPtuOsK8XAVbJwmUqE",
  };
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
