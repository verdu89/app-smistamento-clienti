/** Maintenance Mode — toggle globale e guardie invio (senza toast / popup) */

/** Stato manutenzione */
function isMaintenanceOn_() {
  const v =
    PropertiesService.getScriptProperties().getProperty("MAINTENANCE_MODE");
  return v === "on" || v === "true" || v === "1";
}

function setMaintenanceMode(on) {
  PropertiesService.getScriptProperties().setProperty(
    "MAINTENANCE_MODE",
    on ? "on" : "off"
  );
}

/** Messaggio standard bloccante per trigger / invii */
function maintenanceMessage_() {
  return "Errore script: 🚧 Il sistema è in manutenzione. Gli invii di email e messaggi sono temporaneamente disabilitati. Riprova più tardi.";
}

/** Indicatore persistente: rinomina titolo file + colore tab di Main (nessuna modifica dati) */
function setMaintenanceIndicators_(on) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // 1) Titolo del file (prefisso 🚧, poi ripristino)
  const PROPS = PropertiesService.getScriptProperties();
  const KEY_TITLE = "ORIG_SPREADSHEET_TITLE";

  if (on) {
    if (!PROPS.getProperty(KEY_TITLE)) {
      PROPS.setProperty(KEY_TITLE, ss.getName());
    }
    const origTitle = PROPS.getProperty(KEY_TITLE) || ss.getName();
    const newTitle = "🚧 " + origTitle + " (Manutenzione attiva)";
    if (ss.getName() !== newTitle) {
      ss.rename(newTitle);
    }
  } else {
    const origTitle = PROPS.getProperty(KEY_TITLE);
    if (origTitle) {
      ss.rename(origTitle);
      PROPS.deleteProperty(KEY_TITLE);
    }
  }

  // 2) Colore tab del foglio Main
  const main = ss.getSheetByName("Main");
  if (main) {
    try {
      main.setTabColor(on ? "#EA4335" : null); // rosso Google
    } catch (_) {}
  }
}

/** Wrapper sicuro per email (nessun alert / toast) */
function safeSendEmail_(payloadOrTo, subject, body) {
  const callerStack = (new Error().stack || "").toString();

  // ✅ Lista di funzioni autorizzate a inviare email ANCHE in manutenzione
  const allowedDuringMaintenance = [
    "importaLeadDaMetaNuovi", // <-- puoi aggiungere altre funzioni se necessario empio "nomefunzione"
  ];

  if (isMaintenanceOn_()) {
    const isAllowed = allowedDuringMaintenance.some((fn) =>
      callerStack.includes(fn)
    );
    if (!isAllowed) {
      Logger.log(maintenanceMessage_());
      return { ok: false, maintenance: true, message: maintenanceMessage_() };
    } else {
      Logger.log(
        "✅ Bypass manutenzione email per: " +
          allowedDuringMaintenance.find((fn) => callerStack.includes(fn))
      );
    }
  }

  // ✅ Invio email invariato (con supporto a payload object oppure parametri singoli)
  if (typeof payloadOrTo === "object") {
    MailApp.sendEmail(payloadOrTo);
  } else {
    MailApp.sendEmail({ to: payloadOrTo, subject: subject, htmlBody: body });
  }

  return { ok: true };
}

/** Wrapper sicuro per chiamate HTTP (nessun alert / toast) */
function safeFetch_(url, options) {
  const callerStack = (new Error().stack || "").toString();

  // ✅ Lista di funzioni autorizzate a funzionare anche in manutenzione
  const allowedDuringMaintenance = [
    "importaLeadDaMetaNuovi", // <-- Qui puoi aggiungere altre funzioni future esempio "nomefunzione"
  ];

  if (isMaintenanceOn_()) {
    const isAllowed = allowedDuringMaintenance.some((fn) =>
      callerStack.includes(fn)
    );

    if (!isAllowed) {
      Logger.log(maintenanceMessage_());
      throw new Error(maintenanceMessage_());
    } else {
      Logger.log(
        "✅ Bypass manutenzione per: " +
          allowedDuringMaintenance.find((fn) => callerStack.includes(fn))
      );
    }
  }

  return UrlFetchApp.fetch(url, options);
}

/** Pulizia celle con messaggi di manutenzione, in colonne specifiche */
/** Pulizia celle con messaggi di manutenzione, in colonne specifiche (match dinamico) */
/** Pulizia celle con messaggi di manutenzione (robusta: nomi fogli/colonne flessibili, match dinamico) */
/** Pulizia celle con messaggi di manutenzione / errori script */
function clearMaintenanceMessages_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Frasi da intercettare (minuscolo)
  const MATCH_KEYWORDS = ["manutenzione", "script error", "errore script"];

  // Config: fogli e colonne da pulire
  const targets = [
    { sheets: ["Main"], cols: ["Messaggio Benvenuto"] },
    {
      sheets: ["Recensioni Extra", "Recensioni extra"],
      cols: ["Data richiesta su whatsapp", "Messaggi non inviati"],
    },
  ];

  targets.forEach(({ sheets, cols }) => {
    let sheet = null;
    for (const name of sheets) {
      sheet = ss.getSheetByName(name);
      if (sheet) break;
    }
    if (!sheet) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    const norm = (s) =>
      (s || "").toString().toLowerCase().replace(/\s+/g, " ").trim();
    const headerMap = {};
    header.forEach((h, i) => (headerMap[norm(h)] = i + 1));

    let idxMap = null;
    try {
      idxMap = getColumnIndexes(sheet);
    } catch (_) {}

    cols.forEach((colName) => {
      let col = null;
      if (idxMap && idxMap[colName] != null) col = idxMap[colName];
      if (!col && idxMap) {
        for (const k in idxMap) {
          if (norm(k) === norm(colName)) {
            col = idxMap[k];
            break;
          }
        }
      }
      if (!col) col = headerMap[norm(colName)] || null;
      if (!col) return;

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      const range = sheet.getRange(2, col, lastRow - 1, 1);
      const vals = range.getValues();

      const toClear = [];
      for (let r = 0; r < vals.length; r++) {
        const v = (vals[r][0] || "").toString().toLowerCase();
        if (MATCH_KEYWORDS.some((word) => v.includes(word))) {
          toClear.push(r + 2);
        }
      }

      if (!toClear.length) return;
      const a1s = toClear.map((r) => sheet.getRange(r, col).getA1Notation());
      sheet.getRangeList(a1s).clearContent();
    });
  });
}

/** Menu rapido */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("⚙️ Manutenzione")
      .addItem("Attiva manutenzione", "menuActivateMaintenance_")
      .addItem("Disattiva manutenzione", "menuDeactivateMaintenance_")
      .addSeparator()
      .addItem("Mostra stato", "menuShowMaintenance_")
      .addToUi();
  } catch (_) {}
}

function menuActivateMaintenance_() {
  setMaintenanceMode(true);
  setMaintenanceIndicators_(true); // titolo + tab color
}

function menuDeactivateMaintenance_() {
  clearMaintenanceMessages_(); // pulizia errori dalle colonne richieste
  setMaintenanceMode(false);
  setMaintenanceIndicators_(false); // ripristino titolo + tab color
}

function menuShowMaintenance_() {
  const ui = SpreadsheetApp.getUi();
  const msg = isMaintenanceOn_()
    ? "🚧 Modalità manutenzione ATTIVA: invii esterni bloccati."
    : "✅ Modalità manutenzione DISATTIVATA: invii abilitati.";
  // Alert va bene perché è azione manuale da menu (no trigger)
  try {
    ui.alert("Stato sistema", msg, ui.ButtonSet.OK);
  } catch (_) {}
}
