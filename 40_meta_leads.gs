/** Meta (Facebook) Leads
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */


function importaLeadDaMetaNuovi() {
  const scriptProps = PropertiesService.getScriptProperties();

  const SYSTEM_TOKEN = scriptProps.getProperty("META_ACCESS_TOKEN"); // System User Token
  const FORM_ID = scriptProps.getProperty("META_FORM_ID"); // Form ID
  const SHEET_NAME = scriptProps.getProperty("META_SHEET_NAME") || "Leads";
  const ID_COLUMN_NAME = scriptProps.getProperty("META_ID_COLUMN") || "Lead ID";
  const PAGE_ID = scriptProps.getProperty("META_PAGE_ID"); // Page ID

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Intestazioni del foglio
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndexes = getColumnIndexes(headers);

  const idColIndex = colIndexes[ID_COLUMN_NAME];
  if (idColIndex === undefined) {
    throw new Error(
      `❌ Colonna "${ID_COLUMN_NAME}" non trovata. Intestazioni: ${headers.join(
        ", "
      )}`
    );
  }

  // 🔹 funzione per normalizzare i numeri di telefono
  function normalizePhone(tel) {
    if (!tel) return "";
    let clean = String(tel).replace(/\D/g, ""); // toglie tutto tranne cifre
    if (clean.startsWith("39") && clean.length > 10) {
      clean = clean.substring(2); // rimuove prefisso internazionale
    }
    return clean;
  }

  // Recupera chiavi già presenti
  const lastRow = sheet.getLastRow();
  let existingKeys = [];
  if (lastRow > 1) {
    const idRange = sheet
      .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .getValues();
    idRange.forEach((row) => {
      const leadId = row[idColIndex] ? row[idColIndex].toString().trim() : "";
      const nome =
        colIndexes["Nome"] !== undefined
          ? (row[colIndexes["Nome"]] || "").toString().trim().toLowerCase()
          : "";
      const tel =
        colIndexes["Telefono"] !== undefined
          ? normalizePhone(row[colIndexes["Telefono"]])
          : "";
      const email =
        colIndexes["Email"] !== undefined
          ? (row[colIndexes["Email"]] || "").toString().trim().toLowerCase()
          : "";

      // 🔹 usa SOLO una chiave coerente
      const uniqueKey = leadId || `${nome}|${tel}|${email}`;
      if (uniqueKey) existingKeys.push(uniqueKey);
    });
  }
  const existingSet = new Set(existingKeys);

  // 1) Recupera Page Token
  const accountsUrl = `https://graph.facebook.com/v19.0/me/accounts?access_token=${SYSTEM_TOKEN}`;
  const accountsResp = UrlFetchApp.fetch(accountsUrl);
  const accounts = JSON.parse(accountsResp.getContentText());
  const pageObj = (accounts.data || []).find((p) => p.id === PAGE_ID);
  if (!pageObj)
    throw new Error(`❌ La pagina ${PAGE_ID} non è accessibile col token`);
  const PAGE_TOKEN = pageObj.access_token;

  // 2) Scarica i lead
  const url = `https://graph.facebook.com/v19.0/${FORM_ID}/leads?access_token=${PAGE_TOKEN}`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());

  if (!data.data || data.data.length === 0) {
    Logger.log("ℹ️ Nessun lead ricevuto");
    return;
  }

  // 3) Inserisci solo i lead nuovi
  data.data.forEach((lead) => {
    const leadId = lead.id;
    const createdTime = lead.created_time;

    const record = {};
    (lead.field_data || []).forEach((f) => {
      // 🔴 ignora il campo "lead_status"
      if ((f.name || "").toLowerCase() === "lead_status") {
        return;
      }
      record[f.name] = (f.values || []).join(", ");
    });

    // Pulizia campi
    let telefono = normalizePhone(record.phone_number || "");
    let provincia = (record.inserisci_la_provincia_di_consegna || "")
      .replace(/_/g, " ")
      .trim();
    const luogoConsegna = findFieldValue(lead.field_data, {
      mustInclude: ["luogo", "consegna"],
    });
    const messaggio = record.descrivici_quali_sono_le_tue_esigenze || "";

    // Fallback robusto
    const fallbackKey =
      (record.full_name || "").trim().toLowerCase() +
      "|" +
      telefono +
      "|" +
      (record.email || "").trim().toLowerCase();

    // Chiave finale
    const uniqueKey = leadId || fallbackKey;

    // Evita duplicati
    if (existingSet.has(uniqueKey)) {
      Logger.log("ℹ️ Lead già presente: " + uniqueKey);
      return;
    }

    // Riempi le colonne
    const row = new Array(headers.length).fill("");
    row[colIndexes["Data e ora"]] = createdTime;
    row[colIndexes["Nome"]] = record.full_name || "";
    row[colIndexes["Telefono"]] = telefono;
    row[colIndexes["Email"]] = record.email || "";
    row[colIndexes["Provincia"]] = provincia;
    row[colIndexes["Luogo di Consegna"]] = luogoConsegna;
    row[colIndexes["Messaggio"]] = messaggio;

    // Lead ID o fallback in Note
    if (leadId) {
      row[colIndexes[ID_COLUMN_NAME]] = leadId;
    } else if (colIndexes["Note"] !== undefined) {
      row[colIndexes["Note"]] = "FallbackKey: " + fallbackKey;
    }

    // Aggiunge la riga
    sheet.appendRow(row);

    // Aggiorna il set
    existingSet.add(uniqueKey);
  });

  Logger.log("✅ Importazione completata");
}
