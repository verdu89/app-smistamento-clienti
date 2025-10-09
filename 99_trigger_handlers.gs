/** Trigger handlers
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */

function onEditInstalled(e) {
  if (!e || !e.source || !e.range) return;

  var sheet = e.source.getActiveSheet();
  var fogliAbilitati = ["Main", "Recensioni Extra"];
  if (!fogliAbilitati.includes(sheet.getName())) return;

  var editedCell = e.range;
  var data = sheet.getDataRange().getValues();
  var cols = getColumnIndexes(data[0]);

  // 🔹 Gestione "Vendita Conclusa?"
  if ("Vendita Conclusa?" in cols) {
    var colVendita = cols["Vendita Conclusa?"] + 1;
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
    }
  }

  // 🔹 Invio email se spuntata "Richiedi Recensione"
  if (
    editedCell.getColumn() === cols["Richiedi Recensione"] + 1 &&
    editedCell.getValue() === true
  ) {
    var row = editedCell.getRow();
    var email = sheet.getRange(row, cols["Email"] + 1).getValue();
    var dataRecensione = sheet
      .getRange(row, cols["Data richiesta recensione"] + 1)
      .getValue();

    if (!email || dataRecensione) return;

    var rawNomeCliente =
      sheet.getRange(row, cols["Nome"] + 1).getValue() || "Cliente";
    var nomeCliente = formatNameProperly(rawNomeCliente.toString().trim());

    const { subject, body } = buildReviewEmail(nomeCliente);
    sendEmail(email, subject, body);

    sheet
      .getRange(row, cols["Data richiesta recensione"] + 1)
      .setValue(new Date().toLocaleDateString());
  }
}

function onEditInstalled_Main(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== "Main") return; // Cambia se il tab ha un altro nome

    var row = e.range.getRow();
    var col = e.range.getColumn();

    // Leggi intestazioni
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var normalizedHeaders = {};
    headers.forEach((h, i) => {
      normalizedHeaders[h.toString().trim().toLowerCase()] = i;
    });

    // Campi da tracciare con timestamp
    var fieldsToTrack = [
      "Stato",
      "Note",
      "Vendita Conclusa?",
      "Data Preventivo",
      "Importo Preventivo",
      "Intestatario Contratto",
    ];

    fieldsToTrack.forEach(function (field) {
      var normalizedField = field.toLowerCase();
      if (
        normalizedHeaders[normalizedField] !== undefined &&
        col === normalizedHeaders[normalizedField] + 1
      ) {
        Logger.log("✏️ Modifica rilevata su '" + field + "' alla riga " + row);

        // Assicura la colonna timestamp adiacente
        var tsCol = ensureTimestampColumnAdjacentHidden(sheet, field);
        if (tsCol !== null) {
          var tsCell = sheet.getRange(row, tsCol + 1);

          // ✅ Rimuovi validazione per evitare che metta un dropdown
          tsCell.setDataValidation(null);

          // ✅ Scrivi timestamp ISO
          tsCell.setValue(new Date().toISOString());

          // ✅ Imposta formato testo per evitare conversioni strane
          tsCell.setNumberFormat("@");

          Logger.log(
            "✅ Timestamp aggiornato per '" +
              field +
              "' in colonna " +
              (tsCol + 1)
          );
        } else {
          Logger.log("⚠️ Nessuna colonna timestamp trovata per " + field);
        }
      }
    });
  } catch (err) {
    Logger.log("❌ Errore in onEditInstalled_Main: " + err);
  }
}

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

    Logger.log("🆕 Creata colonna timestamp " + tsName);

    const newHeaders = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    cols = getColumnIndexes(newHeaders);
  }
  return cols[tsName];
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
