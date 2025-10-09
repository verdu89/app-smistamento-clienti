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

function onEditInstalled_Vendor(e) {
  if (!e || !e.source || !e.range) return;

  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Dati") return; // ✅ Solo foglio "Dati"

  var editedCell = e.range;
  var data = sheet.getDataRange().getValues();
  var cols = getColumnIndexes(data[0]);

  // Campi che tracciamo
  var trackFields = [
    "Stato",
    "Note",
    "Data Preventivo",
    "Importo Preventivo",
    "Vendita Conclusa?",
    "Intestatario Contratto",
  ];

  trackFields.forEach(function (field) {
    if (field in cols) {
      var colIndex = cols[field] + 1; // 1-based
      if (editedCell.getColumn() === colIndex) {
        var tsColumn = ensureTimestampColumnAdjacentHidden(sheet, field);
        if (tsColumn !== null) {
          sheet
            .getRange(editedCell.getRow(), tsColumn + 1)
            .setValue(new Date().toISOString());
          Logger.log(`🕒 Aggiornato timestamp per ${field}`);
        }
      }
    }
  });
}
