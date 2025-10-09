/** Gmail ↔ Main Reconciliation
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
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
