/** Core business logic
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
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


function alignPrevPeriodEnd(
    prevPeriodStart,
    workingDaysToMatch,
    span /* "month"|"year" */
  ) {
    // Calcola la data di fine del periodo precedente in modo da avere lo stesso # di giorni lavorativi
    const start = stripTime(prevPeriodStart);
    // se span è "year", fine base = 31/12 prev year; se mese, ultimo giorno del mese precedente
    let theoreticalEnd;
    if (span === "year") {
      theoreticalEnd = new Date(start.getFullYear(), 11, 31);
    } else {
      theoreticalEnd = new Date(start.getFullYear(), start.getMonth() + 1, 0);
    }
    // Trova la data che garantisce lo stesso numero di giorni lavorativi
    let end = new Date(start);
    let count = 0;
    while (end <= theoreticalEnd && count < workingDaysToMatch) {
      const day = end.getDay();
      if (day >= 1 && day <= 5) count++;
      if (count >= workingDaysToMatch) break;
      end.setDate(end.getDate() + 1);
    }
    return stripTime(end);
  }


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


function checkForNewRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Recensioni Extra");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxNome = headers.indexOf("Nome");
  const idxTelefono = headers.indexOf("Telefono");
  const idxEmail = headers.indexOf("Email");
  const idxRichiesta = headers.indexOf("Richiedi Recensione");
  const idxData = headers.indexOf("Data richiesta recensione");

  for (let i = 1; i < data.length; i++) {
    if (data[i][idxRichiesta] === true && data[i][idxData] === "") {
      const rawNomeCliente = data[i][idxNome] || "Cliente";
      const nomeCliente = formatNameProperly(rawNomeCliente.toString().trim());

      const telefono = data[i][idxTelefono] || "";
      const email = data[i][idxEmail];

      if (!email) continue;

      // Se vuoi, puoi passare anche il telefono all'email
      const { subject, body } = buildReviewEmail(nomeCliente, telefono); // <-- solo se modifichi la funzione
      sendEmail(email, subject, body);

      // Aggiorno la data richiesta recensione
      sheet
        .getRange(i + 1, idxData + 1)
        .setValue(new Date().toLocaleDateString());
    }
  }
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


function debugMainSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var data = mainSheet.getDataRange().getValues();
  Logger.log("📌 Dati dal foglio Main: " + JSON.stringify(data.slice(0, 5))); // Mostra le prime 5 righe
}


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


function findFieldValue(fieldData, options) {
  if (!fieldData || fieldData.length === 0) return "";

  const normalize = (s) =>
    s
      .toString()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "") // rimuove accenti
      .replace(/[^a-z0-9]+/g, "_") // non alfanumerici -> _
      .replace(/^_+|_+$/g, ""); // trim _

  // Mappa "nome_normalizzato -> value"
  const normMap = new Map();
  fieldData.forEach((fd) => {
    const keyN = normalize(fd.name || "");
    const val = (fd.values || []).join(", ");
    normMap.set(keyN, val);
  });

  // 1) tentativo con nomi esatti normalizzati
  for (const ex of options.exactNormalized || []) {
    if (normMap.has(ex)) return normMap.get(ex);
  }

  // 2) fallback: include tutte le parole chiave indicate
  const must = (options.mustInclude || []).map(normalize);
  for (const fd of fieldData) {
    const n = normalize(fd.name || "");
    if (must.every((tok) => n.includes(tok))) {
      return (fd.values || []).join(", ");
    }
  }

  return "";
}


function formatMailBody(obj) { // function to spit out all the keys/values from the form in HTML
  var result = "";
  for (var key in obj) { // loop over the object passed to the function
    result += "<h5 style='text-transform: capitalize; margin-bottom: 0'>" + key + "</h5><div>" + obj[key] + "</div>";
    // for every key, concatenate an `<h4 />`/`<div />` pairing of the key name and its value, 
    // and append it to the `result` string created at the start.
  }
  return result; // once the looping is done, `result` will be one long string to put in the email body
}


function formatNameProperly(name) {
  return name
    .toLowerCase()
    .split(" ")
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}


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


function getWorkingDaysInRange(start, end) {
    let days = 0;
    const d = new Date(start);
    while (d <= end) {
      const day = d.getDay();
      if (day >= 1 && day <= 5) days++;
      d.setDate(d.getDate() + 1);
    }
    return Math.max(days, 1);
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


function normalizzaData(val) {
    // Se mancante/illeggibile, assegna la data di oggi per coerenza dei KPI temporali
    return parseCustomDate(val) || stripTime(today);
  }


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


function parseCustomDate(val) {
    if (val instanceof Date && !isNaN(val)) return stripTime(val);
    if (typeof val === "string" && val.trim() !== "") {
      const parsed = new Date(val);
      if (!isNaN(parsed)) return stripTime(parsed);
    }
    return null;
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


function resetLastProcessedRow() {
  PropertiesService.getScriptProperties().deleteProperty("lastProcessedRow");
  Logger.log("🔄 Ultima riga elaborata resettata!");
}


function safeDiv(a, b) {
    return b > 0 ? a / b : 0;
  }


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


function sortPeriodKeys(a, b) {
    // Ordina "YYYY-XX" correttamente
    const [ya, xa] = a.split("-").map(Number);
    const [yb, xb] = b.split("-").map(Number);
    return ya === yb ? xa - xb : ya - yb;
  }


function testEstrazionePezzi() {
  const messaggio = "Vorrei 3 finestre, una porta finestra e due persiane.";
  const numero = getNumeroPezziConOpenAI(messaggio);
  Logger.log("🧪 Numero rilevato: " + numero);
}


function testRowCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var data = mainSheet.getDataRange().getValues();
  Logger.log("🔎 Numero effettivo di righe lette: " + data.length);
}
