/** WhatsApp functions
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */

function inviaBenvenutiWhatsApp() {
  // ✅ Lock per evitare doppie esecuzioni parallele
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("⛔ inviaBenvenutiWhatsApp già in esecuzione, salto.");
    return;
  }

  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const BOT_SERVER_URL = scriptProps.getProperty("BOT_SERVER_URL");

    if (!BOT_SERVER_URL) {
      Logger.log("❌ BOT_SERVER_URL non trovato nelle proprietà dello script!");
      return;
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("⚠️ Nessuna riga di dati trovata.");
      return;
    }

    const headers = data[0];
    const now = new Date();

    // Trova indici colonne
    const idxTelefono = headers.indexOf("Telefono");
    const idxNome = headers.indexOf("Nome");
    const idxBenvenuto = headers.indexOf("Messaggio Benvenuto");

    Logger.log(
      `📑 Indici colonne → Telefono:${idxTelefono}, Nome:${idxNome}, Benvenuto:${idxBenvenuto}`
    );

    for (let i = 1; i < data.length; i++) {
      const telefono = data[i][idxTelefono];
      const nome = data[i][idxNome];
      const benvenuto = data[i][idxBenvenuto];

      if (!telefono) {
        Logger.log(`⏭️ Riga ${i + 1}: manca il numero di telefono → salto`);
        continue;
      }
      if (benvenuto) {
        Logger.log(`⏭️ Riga ${i + 1}: benvenuto già inviato o segnato → salto`);
        continue;
      }

      const url = BOT_SERVER_URL + "/benvenuto";
      const payload = { numero: String(telefono), nome: String(nome || "") };
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      };

      // 🛡️ Blocca subito in modo che anche se crasha non rimanda
      sheet.getRange(i + 1, idxBenvenuto + 1).setValue("IN_ATTESA");
      Logger.log(
        `📡 Riga ${i + 1}: invio richiesta a bot → ${JSON.stringify(payload)}`
      );

      try {
        const response = safeFetch_(url, options);
        const text = response.getContentText();
        Logger.log(`📩 Riga ${i + 1}: risposta server → ${text}`);

        const dataRes = JSON.parse(text);

        if (dataRes.ok) {
          sheet
            .getRange(i + 1, idxBenvenuto + 1)
            .setValue(
              Utilities.formatDate(
                new Date(),
                Session.getScriptTimeZone(),
                "dd/MM/yyyy HH:mm"
              )
            );
          Logger.log(`✅ Riga ${i + 1}: WA benvenuto accodato con successo`);
        } else {
          sheet
            .getRange(i + 1, idxBenvenuto + 1)
            .setValue("ERRORE: " + (dataRes.error || "Errore invio"));
          Logger.log(
            `⚠️ Riga ${i + 1}: errore server → ${
              dataRes.error || "Errore invio"
            }`
          );
        }
      } catch (err) {
        sheet.getRange(i + 1, idxBenvenuto + 1).setValue("🚧 ERRORE SCRIPT");
        Logger.log(`❌ Riga ${i + 1}: eccezione invio → ${err}`);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function inviaRecensioniWhatsApp() {
  const scriptProps = PropertiesService.getScriptProperties();
  const BOT_SERVER_URL = scriptProps.getProperty("BOT_SERVER_URL");

  if (!BOT_SERVER_URL) {
    Logger.log("❌ BOT_SERVER_URL non trovato nelle proprietà dello script!");
    return;
  }

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recensioni Extra");
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("⚠️ Nessuna riga di dati trovata.");
    return;
  }

  const headers = data[0];
  const now = new Date();

  // Trova indici colonne
  const idxTelefono = headers.indexOf("Telefono");
  const idxRichiedi = headers.indexOf("Richiedi Recensione");
  const idxDataMail = headers.indexOf("Data richiesta recensione");
  const idxDataWA = headers.indexOf("Data richiesta su whatsapp");
  const idxMsgFail = headers.indexOf("Messaggi non inviati");

  Logger.log(
    `📑 Indici colonne → Telefono:${idxTelefono}, Richiedi:${idxRichiedi}, Mail:${idxDataMail}, WA:${idxDataWA}, Fail:${idxMsgFail}`
  );

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const richiedi = row[idxRichiedi];
    const dataEmail = row[idxDataMail];
    const telefono = row[idxTelefono];
    const dataWA = row[idxDataWA];

    // ---- Controlli preliminari ----
    if (!richiedi) {
      Logger.log(`⏭️ Riga ${i + 1}: richiesta recensione non attiva → salto`);
      continue;
    }
    if (!dataEmail) {
      Logger.log(`⏭️ Riga ${i + 1}: manca la data email → salto`);
      continue;
    }
    if (!telefono) {
      Logger.log(`⏭️ Riga ${i + 1}: manca il numero di telefono → salto`);
      continue;
    }
    if (dataWA) {
      Logger.log(`⏭️ Riga ${i + 1}: WA già inviato o segnato → salto`);
      continue;
    }

    // ---- Controllo tempo ----
    const diffOre = (now - new Date(dataEmail)) / (1000 * 60 * 60);
    if (diffOre < 24) {
      Logger.log(
        `⏸️ Riga ${i + 1}: solo ${diffOre.toFixed(
          1
        )}h dalla mail, troppo presto → salto`
      );
      continue;
    }

    // ---- Invio richiesta al bot ----
    const url = BOT_SERVER_URL + "/richiedi-recensione";
    const payload = { numero: String(telefono) };
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    Logger.log(
      `📡 Riga ${i + 1}: invio richiesta a bot → ${JSON.stringify(payload)}`
    );

    // 🔒 Previeni duplicazione in caso di crash
    sheet.getRange(i + 1, idxDataWA + 1).setValue("IN_ATTESA");

    try {
      const response = safeFetch_(url, options);
      const text = response.getContentText();
      Logger.log(`📩 Riga ${i + 1}: risposta server → ${text}`);

      const dataRes = JSON.parse(text);

      if (dataRes.ok) {
        sheet
          .getRange(i + 1, idxDataWA + 1)
          .setValue(
            Utilities.formatDate(
              new Date(),
              Session.getScriptTimeZone(),
              "dd/MM/yyyy HH:mm"
            )
          );
        Logger.log(`✅ Riga ${i + 1}: WA accodato con successo`);
      } else {
        sheet
          .getRange(i + 1, idxDataWA + 1)
          .setValue("ERRORE: " + (dataRes.error || "Errore invio"));
        Logger.log(
          `⚠️ Riga ${i + 1}: errore server → ${dataRes.error || "Errore invio"}`
        );

        if (idxMsgFail >= 0 && dataRes.message) {
          sheet.getRange(i + 1, idxMsgFail + 1).setValue(dataRes.message);
          Logger.log(
            `📝 Riga ${
              i + 1
            }: scritto messaggio errore in colonna 'Messaggi non inviati'`
          );
        }
      }
    } catch (err) {
      sheet.getRange(i + 1, idxDataWA + 1).setValue("🚧 ERRORE SCRIPT");
      if (idxMsgFail >= 0)
        sheet.getRange(i + 1, idxMsgFail + 1).setValue("Script error");
      Logger.log(`❌ Riga ${i + 1}: eccezione invio → ${err}`);
    }
  }
}

function normalizzaProvenienza(prov) {
  if (!prov) return "Altro";
  prov = prov.toString().toLowerCase().trim();
  if (prov.includes("cagliari")) return "Showroom Cagliari";
  if (prov.includes("macchiareddu")) return "Showroom Macchiareddu";
  if (prov.includes("nuoro")) return "Showroom Nuoro";
  if (prov.includes("google")) return "Google";
  if (prov.includes("facebook")) return "Facebook";
  if (prov.includes("instagram")) return "Instagram";
  if (prov.includes("whatsapp")) return "Whatsapp";
  if (prov.includes("mail") || prov.includes("email")) return "Email";
  if (prov.includes("chiamata")) return "Chiamata";
  if (prov.includes("passaparola")) return "Passaparola";
  return prov.charAt(0).toUpperCase() + prov.slice(1);
}
