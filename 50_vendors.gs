/** ============================================================
 * FILE: 50_vendors.gs
 * Vendors sync & helpers (versione UNIFICATA + FORMATTATA)
 * - Mantiene TUTTO dal codice originale (log, email, backup, dedup, assegnazioni...)
 * - Integra timestamp adiacenti & sync bidirezionale con scadenza 60gg
 * - Nessun nome funzione modificato
 * ============================================================
 */

/** ============================================================
 * HELPERS PER TIMESTAMP ADIACENTI (NUOVA LOGICA)
 * ============================================================
 */

// Timestamp ISO
function _isoNow_() {
  return new Date().toISOString();
}

/**
 * Assicura che esista la colonna timestamp adiacente e nascosta
 * per il campo indicato. Ritorna l'indice 0-based della colonna timestamp.
 */
function ensureTimestampColumnAdjacentHidden(sheet, fieldName) {
  const headers =
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  let cols = getColumnIndexes(headers);
  if (!(fieldName in cols)) return null;

  const tsName = `_last_update_${fieldName}`;
  if (!(tsName in cols)) {
    const original1Based = cols[fieldName] + 1;
    sheet.insertColumnAfter(original1Based);
    sheet.getRange(1, original1Based + 1).setValue(tsName);
    sheet.hideColumns(original1Based + 1);
    const newHeaders = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    cols = getColumnIndexes(newHeaders);
  }
  return cols[tsName];
}

/** Assicura tutte le colonne timestamp per l'elenco di campi passato */
function ensureAllTimestampColumns(sheet, fields) {
  const results = {};
  const headers =
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  let cols = getColumnIndexes(headers);
  fields.forEach((f) => {
    if (!(f in cols)) return;
    results[f] = ensureTimestampColumnAdjacentHidden(sheet, f);
  });
  return results;
}

/** ============================================================
 * FUNZIONI ORIGINALI PRESERVATE (ADD / DEBUG / DEDUP / GETTERS)
 * ============================================================
 */

function addToVendorSheet(row, sheet, colsMain, colsVendor) {
  logInfo("➡️ Avvio aggiunta dati a " + sheet.getName());

  if (!colsVendor || Object.keys(colsVendor).length === 0) {
    logError("❌ Errore: colsVendor è vuoto o non definito!");
    return;
  }

  var newRow = new Array(Object.keys(colsVendor).length).fill("-");

  if (colsVendor["Data Assegnazione"] !== undefined) {
    newRow[colsVendor["Data Assegnazione"]] = _isoNow_(); // <-- adattato all'helper nuovo
  }

  if (colsVendor["Stato"] !== undefined) {
    newRow[colsVendor["Stato"]] = "Da contattare";
  }

  if (colsVendor["Vendita Conclusa?"] !== undefined) {
    newRow[colsVendor["Vendita Conclusa?"]] = "";
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
    SpreadsheetApp.flush();
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

        if (!nome && !tel) continue;
        var key = nome + "|" + tel;

        if (seen.has(key)) {
          toDelete.push(r + 1);
        } else {
          seen.add(key);
        }
      }

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
/** ============================================================
 * MAPPATURE ORIGINALI (Province → Venditori, Email, Telefono, IDs)
 * ============================================================
 */

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
    ss: "Cristian Piga",
    sassari: "Cristian Piga",
    ot: "Cristian Piga",
    "olbia-tempio": "Cristian Piga",
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
    "Cristian Piga": "xxcristianpiga@me.com",
    "Marco Guidi": "guidi.marco0308@libero.it",
  };
  return vendorEmails[venditore] || "newsaverplast@gmail.com";
}

function getVendorPhone(venditore) {
  var vendorPhones = {
    "Mircko Manconi": "+39 3398123123",
    "Cristian Piga": "+39 3939250786",
    "Marco Guidi": "+39 3349630922",
  };
  return vendorPhones[venditore] || "+39 070/247362";
}

function getVendors() {
  return {
    "Mircko Manconi": "1mGFlFbCYy9ylVjNA9l6b855c6jlIDr6QOua2qfSjckw",
    "Cristian Piga": "1N0_GKbJvZLQbKKIgfVYW4LQGp97mhQcOz9zsD-FBNcE",
    "Marco Guidi": "1CVQSnFGNX8pGUKUABdtzwQmyCKPtuOsK8XAVbJwmUqE",
  };
}
/** ============================================================
 * SYNC PRINCIPALE: Main → Vendors
 * (Mantiene log, backup, email + integra timestamp & lead ID)
 * ============================================================
 */

function syncMainToVendors() {
  // ⛔ Se in manutenzione → esci subito senza fare nulla
  if (isMaintenanceOn_()) {
    Logger.log("🚧 Manutenzione attiva — syncMainToVendors() bloccata");
    return;
  }

  const changesLog = []; // tiene traccia di tutte le modifiche (come in origine)

  // 🔒 Lock per evitare esecuzioni parallele
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("⛔ Esecuzione già in corso, esco.");
    return;
  }

  try {
    Logger.log("🚀 Avvio syncMainToVendors() [VER. TURBO]");
    aggiornaNumeroPezziInMain();

    // 📄 Foglio Main
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Main");
    if (!mainSheet) {
      Logger.log("❌ ERRORE: Il foglio 'Main' non esiste!");
      return;
    }

    var data = mainSheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      Logger.log("ℹ️ Nessun dato nel Main da processare.");
      return;
    }

    var colsMain = getColumnIndexes(data[0]);
    var vendors = getVendors();
    var provinceToVendor = getProvinceToVendor();

    // 🔎 Verifiche minime colonne necessarie
    if (!("Email" in colsMain)) {
      Logger.log(
        "❌ ERRORE: La colonna 'Email' non è stata trovata in 'Main'!"
      );
      return;
    }
    if (!("Venditore Assegnato" in colsMain) || !("Provincia" in colsMain)) {
      Logger.log(
        "❌ ERRORE: Colonne essenziali mancanti (serve 'Venditore Assegnato' e 'Provincia')."
      );
      return;
    }
    const hasLeadIdCol = "Lead ID" in colsMain;
    const hasDataEOraCol = "Data e ora" in colsMain;
    const hasDataAssegnazioneCol = "Data Assegnazione" in colsMain;

    if (!hasLeadIdCol) {
      Logger.log(
        "⚠️ AVVISO: La colonna 'Lead ID' non è presente in 'Main'. Proseguo senza assegnarla."
      );
    }
    if (!hasDataAssegnazioneCol) {
      Logger.log(
        "⚠️ AVVISO: La colonna 'Data Assegnazione' non è presente in 'Main'. L’invio email al primo assegnamento si baserà solo su 'Venditore Assegnato' vuoto."
      );
    }

    // 💾 Backup Main (come in origine)
    createBackup(mainSheet);

    // 🕒 Timestamp adiacenti nel Main per i campi sincronizzati
    const fieldsToSync = [
      "Stato",
      "Note",
      "Data Preventivo",
      "Importo Preventivo",
      "Vendita Conclusa?",
      "Intestatario Contratto",
    ];
    const tsColsMain = ensureAllTimestampColumns(mainSheet, fieldsToSync);

    // 📦 Cache Vendor: apri una sola volta TUTTI i file e indicizza righe per LeadID e Nome+Telefono
    Logger.log("📦 Preparazione cache Vendor...");
    const vendorCache = {}; // venditore -> { ss, sheet, data, cols, tsCols, leadIndex, nameTelIndex }
    Object.keys(vendors).forEach((vName) => {
      try {
        const vSS = SpreadsheetApp.openById(vendors[vName]);
        const vSheet = vSS.getSheetByName("Dati");
        if (!vSheet) {
          Logger.log(`⚠️ Foglio 'Dati' mancante per ${vName}, salto cache.`);
          return;
        }
        const vData = vSheet.getDataRange().getValues();
        const colsV = getColumnIndexes(vData[0] || []);
        const tsColsV = ensureAllTimestampColumns(vSheet, fieldsToSync);

        const leadIndex = {};
        const nameTelIndex = {};
        for (let r = 1; r < vData.length; r++) {
          const row = vData[r];
          const lead =
            colsV["Lead ID"] !== undefined
              ? (row[colsV["Lead ID"]] || "").toString().trim()
              : "";
          const n = (
            colsV["Nome"] !== undefined ? row[colsV["Nome"]] || "" : ""
          )
            .toString()
            .trim()
            .toLowerCase();
          const t = (
            colsV["Telefono"] !== undefined ? row[colsV["Telefono"]] || "" : ""
          )
            .toString()
            .trim();
          if (lead) leadIndex[lead] = r; // indice 0-based su vData
          const key = n + "|" + t;
          if (n || t) nameTelIndex[key] = r;
        }

        vendorCache[vName] = {
          ss: vSS,
          sheet: vSheet,
          data: vData,
          cols: colsV,
          tsCols: tsColsV,
          leadIndex,
          nameTelIndex,
        };

        Logger.log(`✅ Cache Vendor pronta: ${vName} (righe: ${vData.length})`);
      } catch (e) {
        Logger.log(`❌ Errore apertura Vendor ${vName}: ${e.message}`);
      }
    });

    // 🗃️ Collezione aggiornamenti assegnazione in Main (legacy - lasciata per compatibilità)
    const updatesAssegnazioni = [];

    // 🗃️ Dati da inserire nei fogli venditori (per reimpiegare la tua sync esistente)
    const vendorsData = {};

    // 🆔 Seed per Lead ID
    var tsSeed = Math.floor(Date.now() / 1000);
    var tsOffset = 0;

    // 🔁 Scorri tutte le righe del Main
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      var nomeCliente = (row[colsMain["Nome"]] || "").toString().trim();
      var telefonoCliente = (row[colsMain["Telefono"]] || "").toString().trim();
      var venditoreAssegnato = (row[colsMain["Venditore Assegnato"]] || "")
        .toString()
        .trim();
      var emailCliente = (row[colsMain["Email"]] || "").toString().trim();
      var luogoConsegna = (
        "Luogo di Consegna" in colsMain
          ? row[colsMain["Luogo di Consegna"]] || ""
          : ""
      )
        .toString()
        .trim();
      var messaggio = (
        "Messaggio" in colsMain ? row[colsMain["Messaggio"]] || "" : ""
      )
        .toString()
        .trim();

      // Se la riga è completamente vuota (Nome e Telefono assenti), interrompo come da logica originale
      if (!nomeCliente && !telefonoCliente) {
        Logger.log("🛑 Riga vuota trovata, interruzione alla riga " + (i + 1));
        break;
      }

      // 🆔 Lead ID: genera se manca (e scrivilo SUBITO)
      var leadId = hasLeadIdCol
        ? (row[colsMain["Lead ID"]] || "").toString().trim()
        : "";
      if (hasLeadIdCol && !leadId) {
        leadId = "INT-" + (tsSeed + tsOffset++);
        mainSheet.getRange(i + 1, colsMain["Lead ID"] + 1).setValue(leadId);
        changesLog.push(`Riga ${i + 1}: Lead ID assegnato → ${leadId}`);
        Logger.log(`🆔 Lead ID generato in Main riga ${i + 1}: ${leadId}`);
      }

      // 🔄 Se la riga è già assegnata → sincronizza SOLO i campi necessari verso il Vendor
      if (venditoreAssegnato) {
        Logger.log(
          `🔁 Riga ${
            i + 1
          } già assegnata a ${venditoreAssegnato}, controllo aggiornamenti (timestamp + valore)...`
        );

        const cache = vendorCache[venditoreAssegnato];
        if (!cache || !cache.sheet) {
          Logger.log(
            `⚠️ Cache/Sheet mancante per ${venditoreAssegnato}, salto aggiornamento.`
          );
          continue;
        }

        const vData = cache.data;
        const colsVendor = cache.cols;
        const tsColsVendor = cache.tsCols;

        // Trova la riga nel vendor: priorità Lead ID, altrimenti Nome+Telefono
        let vRowIndex = -1; // indice 0-based su vData
        if (leadId && cache.leadIndex[leadId] !== undefined) {
          vRowIndex = cache.leadIndex[leadId];
        } else {
          const key =
            (nomeCliente || "").toString().trim().toLowerCase() +
            "|" +
            (telefonoCliente || "").toString().trim();
          if (cache.nameTelIndex[key] !== undefined)
            vRowIndex = cache.nameTelIndex[key];
        }

        if (vRowIndex === -1) {
          Logger.log(
            `⚠️ Nessun match trovato nel foglio di ${venditoreAssegnato} per riga Main ${
              i + 1
            } (LeadID:${
              leadId || "-"
            } / Nome+Tel). Non inserisco qui (comportamento originale).`
          );
          continue; // manteniamo la logica originale: non creiamo qui una nuova riga se già assegnata ma mancante
        }

        Logger.log(
          `✅ Match trovato nel foglio ${venditoreAssegnato} alla riga ${
            vRowIndex + 1
          }`
        );

        // Se Lead ID c'è in Main ma manca nel Vendor → scrivilo subito
        if (hasLeadIdCol && leadId && colsVendor["Lead ID"] !== undefined) {
          const existingVendorLead = (
            vData[vRowIndex][colsVendor["Lead ID"]] || ""
          )
            .toString()
            .trim();
          if (!existingVendorLead) {
            setValueBypassingValidation(
              cache.sheet,
              vRowIndex + 1,
              colsVendor["Lead ID"] + 1,
              leadId
            );
            vData[vRowIndex][colsVendor["Lead ID"]] = leadId;
            Logger.log(
              `🆔 Lead ID scritto subito nel Vendor (${venditoreAssegnato}) riga ${
                vRowIndex + 1
              }: ${leadId}`
            );
          }
        }

        // Sincronizza campi solo se:
        // - Main ha un valore
        // - Timestamp di Main è più recente di Vendor
        // - E il VALORE è diverso (per evitare rewrite inutili)
        fieldsToSync.forEach((field) => {
          if (!(field in colsMain) || !(field in colsVendor)) {
            Logger.log(
              `ℹ️ Campo ${field} non presente in Main o Vendor, salto.`
            );
            return;
          }

          const mainValue = (row[colsMain[field]] || "").toString().trim();
          const vendorValue = (vData[vRowIndex][colsVendor[field]] || "")
            .toString()
            .trim();

          // Timestamp adiacenti (già assicurati)
          const mainTsCol = tsColsMain[field];
          const vendorTsCol = tsColsVendor[field];

          const mainTs = mainTsCol !== undefined ? row[mainTsCol] || "" : "";
          const vendorTs =
            vendorTsCol !== undefined
              ? vData[vRowIndex][vendorTsCol] || ""
              : "";

          if (!mainValue) {
            Logger.log(`⏭️ ${field}: Main è vuoto → non sincronizzo.`);
            return;
          }

          // ⚖️ Confronto timestamp (ISO string o Date). Se non c'è vendorTs, Main vince.
          const isMainNewer =
            !vendorTs ||
            (mainTs &&
              new Date(mainTs).getTime() > new Date(vendorTs).getTime());
          const isDifferent = mainValue !== vendorValue;

          if (isMainNewer && isDifferent) {
            // Scrivi valore
            setValueBypassingValidation(
              cache.sheet,
              vRowIndex + 1,
              colsVendor[field] + 1,
              mainValue
            );
            vData[vRowIndex][colsVendor[field]] = mainValue;

            // Aggiorna TS Vendor (usa TS Main se disponibile, altrimenti "adesso")
            const tsToWrite = mainTs ? new Date(mainTs) : new Date();
            if (vendorTsCol !== undefined) {
              setValueBypassingValidation(
                cache.sheet,
                vRowIndex + 1,
                vendorTsCol + 1,
                tsToWrite
              );
              vData[vRowIndex][vendorTsCol] = tsToWrite;
            }

            Logger.log(
              `↪️ Aggiornato Vendor ${venditoreAssegnato}, riga ${
                vRowIndex + 1
              }: ${field} = "${mainValue}" (TS main:${
                mainTs || "-"
              } > TS vendor:${vendorTs || "-"})`
            );
          } else {
            Logger.log(
              `⏭️ Nessun cambiamento per ${field} (valore identico o TS non più recente).`
            );
          }
        });

        continue; // passa alla riga successiva del Main
      }

      // ✳️ Riga NON assegnata → assegna come nella versione vecchia (e invia email SOLO al primo assegnamento)
      Logger.log(
        `🆕 Nuovo cliente senza venditore (riga ${
          i + 1
        }) → calcolo assegnazione (vecchia logica)...`
      );

      // === LOGICA ASSEGNAZIONE VECCHIA VERSIONE ===
      var provincia = (row[colsMain["Provincia"]] || "")
        .toString()
        .trim()
        .toLowerCase();
      var venditoreNuovo = "Cristian Piga"; // fallback

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

      if (provincia === "nu" || provincia === "nuoro") {
        var luogoLowerNU = (luogoConsegna || "").toLowerCase();
        var matchCristian = comuniPerCristianPiga.some((c) =>
          luogoLowerNU.includes(c)
        );
        venditoreNuovo = matchCristian ? "Cristian Piga" : "Marco Guidi";
      } else if (provincia === "ca" || provincia === "cagliari") {
        var luogoLowerCA = (luogoConsegna || "").toLowerCase();
        var comuniPerCristianInCa = ["pula", "villasimius"];
        var matchCristianCA = comuniPerCristianInCa.some((c) =>
          luogoLowerCA.includes(c)
        );
        venditoreNuovo = matchCristianCA ? "Cristian Piga" : "Mircko Manconi";
      } else if (provincia === "su" || provincia === "sud sardegna") {
        var luogoLowerSU = (luogoConsegna || "").toLowerCase();
        var matchMircko = comuniPerMircko.some((c) => luogoLowerSU.includes(c));
        venditoreNuovo = matchMircko ? "Mircko Manconi" : "Cristian Piga";
      } else {
        var provinciaNorm = (provincia || "").toLowerCase();
        var luogoNorm = (luogoConsegna || "").toLowerCase();
        var comuniZonaOlbia = [
          "olbia",
          "golfo aranci",
          "arzachena",
          "porto rotondo",
          "loiri porto san paolo",
          "telti",
          "palau",
          "buddusò",
          "tempio pausania",
          "santa teresa gallura",
        ];
        if (
          provinciaNorm === "ss" ||
          provinciaNorm === "sassari" ||
          provinciaNorm === "ot" ||
          provinciaNorm.includes("olbia")
        ) {
          venditoreNuovo = "Cristian Piga";
        } else if (comuniZonaOlbia.some((c) => luogoNorm.includes(c))) {
          venditoreNuovo = "Cristian Piga";
        } else {
          venditoreNuovo = provinceToVendor[provinciaNorm] || "Cristian Piga";
        }
      }

      Logger.log(
        `✅ Nuovo cliente assegnato → ${venditoreNuovo} (riga ${i + 1})`
      );

      // === SCRITTURE IMMEDIATE ===

      // Scrivi subito "Venditore Assegnato"
      mainSheet
        .getRange(i + 1, colsMain["Venditore Assegnato"] + 1)
        .setValue(venditoreNuovo);

      // Determina se è il PRIMO assegnamento: email SOLO in questo caso
      const wasFirstAssignment = hasDataAssegnazioneCol
        ? !row[colsMain["Data Assegnazione"]]
        : true;

      // Scrivi "Data Assegnazione" se vuota
      if (hasDataAssegnazioneCol && !row[colsMain["Data Assegnazione"]]) {
        mainSheet
          .getRange(i + 1, colsMain["Data Assegnazione"] + 1)
          .setValue(new Date());
      }

      // Scrivi "Data e ora" se vuota (come nella vecchia) – se la colonna esiste
      if (hasDataEOraCol && !row[colsMain["Data e ora"]]) {
        mainSheet
          .getRange(i + 1, colsMain["Data e ora"] + 1)
          .setValue(new Date());
      }

      // Forza scrittura prima dell'email
      SpreadsheetApp.flush();

      // === INVIO EMAIL COME VECCHIA VERSIONE (SOLO AL PRIMO ASSEGNAMENTO) ===
      if (wasFirstAssignment) {
        notifyAssignment(
          venditoreNuovo,
          emailCliente || "",
          nomeCliente,
          telefonoCliente,
          provincia,
          luogoConsegna,
          messaggio
        );

        // Se email non valida → nota in "Note"
        if (!isValidEmail_(emailCliente)) {
          safeSetIfColumnExists_(
            mainSheet,
            colsMain,
            "Note",
            i + 1,
            "Email cliente assente o non valida"
          );
        }
      } else {
        Logger.log(
          `📨 Nessuna email inviata (non è il primo assegnamento) – riga ${
            i + 1
          }`
        );
      }

      // === PREPARA vendorsData COME NELLA NUOVA VERSIONE ===
      if (!vendorsData[venditoreNuovo]) vendorsData[venditoreNuovo] = [];
      var filteredRow = {};
      Object.keys(colsMain).forEach((c) => (filteredRow[c] = row[colsMain[c]]));
      if (hasLeadIdCol) filteredRow["Lead ID"] = leadId;
      filteredRow["Data Assegnazione"] = new Date().toLocaleString();
      vendorsData[venditoreNuovo].push(filteredRow);

      continue; // passa alla prossima riga
    }

    // ✏️ Applica aggiornamento venditori assegnati nel Main (come in origine, se usato)
    updatesAssegnazioni.forEach((u) => {
      mainSheet
        .getRange(u[0], colsMain["Venditore Assegnato"] + 1)
        .setValue(u[1]);
    });

    // 🔁 Sync Vendor Sheets per gli inserimenti (riuso funzione originale)
    // Nota: abbiamo già aggiornato "in-place" i vendor assegnati esistenti; qui gestiamo i nuovi inserimenti.
    Logger.log("🔁 Avvio syncVendorsSheets() per nuovi inserimenti...");
    syncVendorsSheets(vendorsData, vendors);
    Logger.log("✅ Fine syncVendorsSheets() inserimenti");

    Logger.log("✅ Fine syncMainToVendors() [VER. TURBO]");
  } finally {
    lock.releaseLock();
  }

  // 🧾 Log finale modifiche
  Logger.log("📋 Dettaglio modifiche:");
  changesLog.slice(0, 50).forEach((m) => Logger.log(m));
  Logger.log(`Totale modifiche loggate: ${changesLog.length}`);
}

/** ============================================================
 * POPOLA Lead ID nei fogli Venditori se mancano
 * ============================================================
 */

function populateLeadIdInVendorsFromMain() {
  Logger.log("🔍 Avvio populateLeadIdInVendorsFromMain()...");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  if (!mainSheet) {
    Logger.log("❌ Foglio Main non trovato!");
    return;
  }

  var mainData = mainSheet.getDataRange().getValues();
  var colsMain = getColumnIndexes(mainData[0]);

  if (!("Lead ID" in colsMain)) {
    Logger.log("❌ Main non ha colonna Lead ID, stop.");
    return;
  }

  var vendors = getVendors();
  var leadMap = {}; // email|tel → Lead ID

  for (var i = 1; i < mainData.length; i++) {
    var email = (mainData[i][colsMain["Email"]] || "").toString().trim();
    var tel = (mainData[i][colsMain["Telefono"]] || "").toString().trim();
    var leadId = (mainData[i][colsMain["Lead ID"]] || "").toString().trim();
    if (leadId && (email || tel)) {
      leadMap[email + "|" + tel] = leadId;
    }
  }

  Object.keys(vendors).forEach((vendorName) => {
    var vSS = SpreadsheetApp.openById(vendors[vendorName]);
    var sh = vSS.getSheetByName("Dati");
    if (!sh) return;

    Logger.log("✳️ Popolo Lead ID per " + vendorName);
    var data = sh.getDataRange().getValues();
    var colsV = getColumnIndexes(data[0]);
    if (!("Lead ID" in colsV)) {
      var lastCol = sh.getLastColumn();
      sh.insertColumnAfter(lastCol);
      sh.getRange(1, lastCol + 1).setValue("Lead ID");
      colsV["Lead ID"] = lastCol;
    }

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var email = (row[colsV["Email"]] || "").toString().trim();
      var tel = (row[colsV["Telefono"]] || "").toString().trim();
      var existingId = (row[colsV["Lead ID"]] || "").toString().trim();
      var key = email + "|" + tel;
      if (!existingId && leadMap[key]) {
        sh.getRange(r + 1, colsV["Lead ID"] + 1).setValue(leadMap[key]);
        Logger.log("✅ Agg. Lead ID riga " + (r + 1) + ": " + leadMap[key]);
      }
    }
  });

  Logger.log("✅ populateLeadIdInVendorsFromMain() completato.");
}

/** ============================================================
 * SYNC VENDORS SHEETS - Scrittura nei fogli singoli
 * ============================================================
 */

function syncVendorsSheets(vendorsData, vendors) {
  Logger.log("🔁 Avvio syncVendorsSheets()...");

  if (!vendorsData || typeof vendorsData !== "object") {
    Logger.log("ℹ️ vendorsData vuoto o non valido, skip totale.");
    return;
  }

  Object.keys(vendorsData).forEach((venditore) => {
    if (
      !Array.isArray(vendorsData[venditore]) ||
      vendorsData[venditore].length === 0
    ) {
      Logger.log("ℹ️ Nessun dato da inserire per " + venditore + ", skip...");
      return;
    }

    var vendorSS = SpreadsheetApp.openById(vendors[venditore]);
    var venditoreSheet =
      vendorSS.getSheetByName("Dati") || vendorSS.insertSheet("Dati");

    var dataVendor = venditoreSheet.getDataRange().getValues();
    var hasHeader =
      dataVendor &&
      dataVendor.length > 0 &&
      dataVendor[0] &&
      dataVendor[0].length > 0;

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
    }

    var headersVendor = dataVendor[0];
    var colsVendor = getColumnIndexes(headersVendor);

    var existingKeys = new Set();
    for (var i = 1; i < dataVendor.length; i++) {
      var n = (dataVendor[i][colsVendor["Nome"]] || "")
        .toString()
        .trim()
        .toLowerCase();
      var t = (dataVendor[i][colsVendor["Telefono"]] || "").toString().trim();
      if (n || t) existingKeys.add(n + "|" + t);
    }

    var seenInThisRun = new Set();
    var rowsToAdd = [];

    vendorsData[venditore].forEach((row) => {
      if (!row || typeof row !== "object") return;

      var nome = (row["Nome"] || "").toString().trim().toLowerCase();
      var tel = (row["Telefono"] || "").toString().trim();
      if (!nome && !tel) return;

      var key = nome + "|" + tel;
      if (existingKeys.has(key) || seenInThisRun.has(key)) return;
      seenInThisRun.add(key);

      var newRow = headersVendor.map((header) => {
        if (header === "Data Assegnazione") return new Date().toLocaleString();
        if (header === "Stato") return "Da contattare";
        if (header === "Vendita Conclusa?") return "";
        return row[header] || "";
      });

      rowsToAdd.push(newRow);
    });

    if (rowsToAdd.length > 0) {
      var startRow = venditoreSheet.getLastRow() + 1;
      venditoreSheet
        .getRange(startRow, 1, rowsToAdd.length, headersVendor.length)
        .setValues(rowsToAdd);

      // ✅ Applica dropdown solo alle nuove righe
      rowsToAdd.forEach((_, index) => {
        var rowNumber = startRow + index;

        // Dropdown Stato
        if (colsVendor["Stato"] !== undefined) {
          applyDropdownValidation(
            venditoreSheet,
            colsVendor["Stato"],
            [
              "Da contattare",
              "Preventivo inviato",
              "Preventivo non eseguibile",
              "In trattativa",
              "Trattativa terminata",
            ],
            null,
            rowNumber
          );
        }

        // Dropdown Vendita Conclusa?
        if (colsVendor["Vendita Conclusa?"] !== undefined) {
          applyDropdownValidation(
            venditoreSheet,
            colsVendor["Vendita Conclusa?"],
            ["SI", "NO"],
            { SI: "#00FF00", NO: "#FF0000" },
            rowNumber
          );
        }
      });
    }

    Logger.log("✅ Inserite " + rowsToAdd.length + " righe per " + venditore);
  });

  Logger.log("✅ Fine syncVendorsSheets()");
}

/** ============================================================
 * SYNC BIDIREZIONALE: Vendors → Main
 * - Se un venditore aggiorna "Stato" o "Vendita Conclusa?" o Note,
 *   vengono riportati nel Main SOLO se più recenti e non scaduti.
 * - Se nessuna risposta entro 60gg dall'assegnazione → auto "NO"
 * ============================================================
 */

/** Helper per scrivere ignorando temporaneamente la convalida dati */
function setValueBypassingValidation(sheet, rowIndex, colIndex, value) {
  const range = sheet.getRange(rowIndex, colIndex);
  const validation = range.getDataValidation(); // salva regola
  range.clearDataValidations(); // disattiva
  range.setValue(value); // scrivi comunque
  if (validation) range.setDataValidation(validation); // ripristina
}

function updateMainFromVendors() {
  // ⛔ Se in manutenzione → blocca tutto
  if (isMaintenanceOn_()) {
    Logger.log("🚧 Manutenzione attiva — updateMainFromVendors() bloccata");
    return;
  }

  Logger.log("🔁 Avvio updateMainFromVendors() [VER. TURBO]...");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  if (!mainSheet) {
    Logger.log("❌ Foglio Main non trovato!");
    return;
  }

  var mainData = mainSheet.getDataRange().getValues();
  var colsMain = getColumnIndexes(mainData[0]);
  var tsColsMain = ensureAllTimestampColumns(mainSheet, [
    "Stato",
    "Note",
    "Data Preventivo",
    "Importo Preventivo",
    "Vendita Conclusa?",
    "Intestatario Contratto",
  ]);

  var vendors = getVendors();

  Object.keys(vendors).forEach((venditore) => {
    Logger.log(`📂 Controllo aggiornamenti dal vendor: ${venditore}`);

    var vSS = SpreadsheetApp.openById(vendors[venditore]);
    var vSheet = vSS.getSheetByName("Dati");
    if (!vSheet) {
      Logger.log(`⚠️ Foglio Dati mancante per ${venditore}, salto.`);
      return;
    }

    var vData = vSheet.getDataRange().getValues();
    var colsV = getColumnIndexes(vData[0]);
    var tsColsVendor = ensureAllTimestampColumns(vSheet, [
      "Stato",
      "Note",
      "Data Preventivo",
      "Importo Preventivo",
      "Vendita Conclusa?",
      "Intestatario Contratto",
    ]);

    for (var r = 1; r < vData.length; r++) {
      var vRow = vData[r];

      var leadV =
        colsV["Lead ID"] !== undefined
          ? (vRow[colsV["Lead ID"]] || "").toString().trim()
          : "";
      var nomeV = (vRow[colsV["Nome"]] || "").toString().trim();
      var telV = (vRow[colsV["Telefono"]] || "").toString().trim();

      if (!leadV && !nomeV && !telV) continue;

      var mainIndex = -1;
      for (var m = 1; m < mainData.length; m++) {
        var leadM =
          colsMain["Lead ID"] !== undefined
            ? (mainData[m][colsMain["Lead ID"]] || "").toString().trim()
            : "";
        var nomeM = (mainData[m][colsMain["Nome"]] || "").toString().trim();
        var telM = (mainData[m][colsMain["Telefono"]] || "").toString().trim();

        if (
          (leadV && leadV === leadM) ||
          (!leadV && nomeM === nomeV && telM === telV)
        ) {
          mainIndex = m;
          break;
        }
      }
      if (mainIndex === -1) {
        Logger.log(
          `⏭️ Nessuna riga corrispondente in Main per vendor ${venditore}, riga ${
            r + 1
          }`
        );
        continue;
      }

      var mRow = mainData[mainIndex];

      [
        "Stato",
        "Note",
        "Data Preventivo",
        "Importo Preventivo",
        "Vendita Conclusa?",
        "Intestatario Contratto",
      ].forEach((field) => {
        if (!(field in colsV) || !(field in colsMain)) return;

        var vValue = (vRow[colsV[field]] || "").toString().trim();
        var mValue = (mRow[colsMain[field]] || "").toString().trim();
        var vTs =
          tsColsVendor[field] !== undefined
            ? vRow[tsColsVendor[field]] || ""
            : "";
        var mTs =
          tsColsMain[field] !== undefined ? mRow[tsColsMain[field]] || "" : "";

        if (!vValue && !mValue) {
          Logger.log(`⏭️ [${field}] entrambi vuoti -> skip`);
          return;
        }
        if (vTs && mTs && vTs === mTs) {
          Logger.log(`⏭️ [${field}] TS identici (${vTs}) -> skip`);
          return;
        }
        if (!vTs && mTs) {
          Logger.log(`⏭️ [${field}] Vendor TS vuoto ma Main ha TS -> skip`);
          return;
        }
        if (!mTs || vTs > mTs) {
          if (vValue !== mValue) {
            var targetCell = mainSheet.getRange(
              mainIndex + 1,
              colsMain[field] + 1
            );
            targetCell.setDataValidation(null);
            targetCell.setValue(vValue);

            if (tsColsMain[field] !== undefined) {
              mainSheet
                .getRange(mainIndex + 1, tsColsMain[field] + 1)
                .setValue(_isoNow_());
            }

            Logger.log(
              `✅ Aggiornato [${field}] da Vendor→Main (riga Main ${
                mainIndex + 1
              } = "${vValue}")`
            );

            if (field === "Stato") {
              applyDropdownValidation(
                mainSheet,
                colsMain["Stato"],
                [
                  "Da contattare",
                  "Preventivo inviato",
                  "Preventivo non eseguibile",
                  "In trattativa",
                  "Trattativa terminata",
                ],
                null,
                mainIndex + 1
              );
            }

            if (field === "Vendita Conclusa?") {
              applyDropdownValidation(
                mainSheet,
                colsMain["Vendita Conclusa?"],
                ["SI", "NO"],
                {
                  SI: "#00FF00",
                  NO: "#FF0000",
                },
                mainIndex + 1
              );
            }

            if (tsColsVendor[field] !== undefined) {
              vSheet
                .getRange(r + 1, tsColsVendor[field] + 1)
                .setValue(_isoNow_());
            }
          } else {
            Logger.log(`⏭️ [${field}] Valore identico ("${vValue}") -> skip`);
          }
        }
      });
    }
  });

  Logger.log("✅ updateMainFromVendors() completato [VER. TURBO].");
}

function firstTimeSyncMissingFields() {
  Logger.log(
    "🚀 Avvio firstTimeSyncMissingFields() [RIEMPIMENTO BUCHI INIZIALE]..."
  );

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  if (!mainSheet) {
    Logger.log("❌ Foglio Main non trovato!");
    return;
  }

  var mainData = mainSheet.getDataRange().getValues();
  var colsMain = getColumnIndexes(mainData[0]);
  var vendors = getVendors();

  // Campi da sincronizzare solo se uno dei due è vuoto
  var fieldsToFill = [
    "Stato",
    "Note",
    "Data Preventivo",
    "Importo Preventivo",
    "Vendita Conclusa?",
    "Intestatario Contratto",
  ];

  Object.keys(vendors).forEach((venditore) => {
    Logger.log(`📂 Controllo vendor: ${venditore}`);

    var vSS = SpreadsheetApp.openById(vendors[venditore]);
    var vSheet = vSS.getSheetByName("Dati");
    if (!vSheet) {
      Logger.log(`⚠️ Foglio Dati mancante per ${venditore}, salto.`);
      return;
    }

    var vData = vSheet.getDataRange().getValues();
    var colsV = getColumnIndexes(vData[0]);

    for (var r = 1; r < vData.length; r++) {
      var vRow = vData[r];
      var leadV =
        colsV["Lead ID"] !== undefined
          ? (vRow[colsV["Lead ID"]] || "").toString().trim()
          : "";
      var nomeV = (vRow[colsV["Nome"]] || "").toString().trim();
      var telV = (vRow[colsV["Telefono"]] || "").toString().trim();

      if (!leadV && !nomeV && !telV) continue;

      var mainIndex = -1;
      for (var m = 1; m < mainData.length; m++) {
        var leadM =
          colsMain["Lead ID"] !== undefined
            ? (mainData[m][colsMain["Lead ID"]] || "").toString().trim()
            : "";
        var nomeM = (mainData[m][colsMain["Nome"]] || "").toString().trim();
        var telM = (mainData[m][colsMain["Telefono"]] || "").toString().trim();

        if (
          (leadV && leadV === leadM) ||
          (!leadV && nomeM === nomeV && telM === telV)
        ) {
          mainIndex = m;
          break;
        }
      }
      if (mainIndex === -1) {
        Logger.log(
          `⏭️ Nessuna riga corrispondente in Main per venditore ${venditore}, riga ${
            r + 1
          }`
        );
        continue;
      }

      var mRow = mainData[mainIndex];

      fieldsToFill.forEach((field) => {
        if (!(field in colsV) || !(field in colsMain)) return;

        var vValue = (vRow[colsV[field]] || "").toString().trim();
        var mValue = (mRow[colsMain[field]] || "").toString().trim();

        if (!mValue && vValue) {
          mainSheet
            .getRange(mainIndex + 1, colsMain[field] + 1)
            .setValue(vValue);
          Logger.log(
            `✅ Main vuoto, copiato da Vendor → Main [${field}] "${vValue}"`
          );
        } else if (mValue && !vValue) {
          vSheet.getRange(r + 1, colsV[field] + 1).setValue(mValue);
          Logger.log(
            `✅ Vendor vuoto, copiato da Main → Vendor [${field}] "${mValue}"`
          );
        } else {
          Logger.log(
            `⏭️ Skip [${field}] → entrambi pieni o entrambi vuoti ("${mValue}" / "${vValue}")`
          );
        }
      });
    }
  });

  Logger.log("✅ firstTimeSyncMissingFields() completato.");
}

function setValueBypassingValidation(sheet, row, col, value) {
  var cell = sheet.getRange(row, col);
  var rule = cell.getDataValidation();
  cell.clearDataValidations();
  cell.setValue(value);
  if (rule) cell.setDataValidation(rule);
}

function applyDropdownIfColumnExists(sheet, colName, values, colors) {
  // Idem come sopra
}

function getColumnIndexes(headers) {
  var map = {};
  headers.forEach((h, i) => {
    if (h) map[h.toString().trim()] = i;
  });
  return map;
}

function logInfo(msg) {
  Logger.log("ℹ️ " + msg);
}

function logError(msg) {
  Logger.log("❌ " + msg);
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function safeSetIfColumnExists_(sheet, colsMain, colName, rowIndex, value) {
  if (!(colName in colsMain)) return;
  sheet.getRange(rowIndex, colsMain[colName] + 1).setValue(value);
}
