/** Email Queue & Processing
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */

function addToEmailQueue(to, subject, body) {
  var sheet = getEmailQueueSheet();
  sheet.appendRow([to, subject, body, 0]);
  logWarning("📌 Email messa in coda per " + to);
}

function buildReviewEmail(nomeCliente) {
  const subject = "Ci racconti com’è andata? 🙌";
  const body = `
Gentile cliente,<br><br>
grazie di cuore per averci scelto! 🙏<br>
Siamo felici di averti aiutato e speriamo che il nostro intervento ti abbia lasciato sereno e soddisfatto.<br><br>
La tua opinione per noi conta moltissimo: ci permette di migliorare ogni giorno e di far conoscere ad altre persone la qualità del nostro lavoro.<br><br>
<b>Ci dedichi un minuto per lasciarci una recensione? 🙌</b><br>
Per noi sarebbe un gesto prezioso e per te un modo semplice per darci una grande mano.<br><br>
<a href="https://maps.app.goo.gl/1gM31niwMtSfPCk16" 
   style="display:inline-block; padding:12px 20px; background:#2563eb; color:#fff; font-weight:bold; border-radius:8px; text-decoration:none;">
✨ Scrivi la tua recensione ✨
</a><br><br>
Il tuo feedback farà davvero la differenza per il nostro team, e ci aiuterà a continuare a offrire il meglio.<br><br>
Con gratitudine,<br>
<b>Il Team Saverplast</b>
`;

  return { subject, body };
}

function getEmailQueueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("EmailQueue");
  if (!sheet) {
    sheet = ss.insertSheet("EmailQueue");
    sheet.appendRow(["Email", "Oggetto", "Corpo", "Tentativi"]);
  }
  return sheet;
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

function sendEmail(to, subject, body) {
  try {
    const res = safeSendEmail_(to, subject, body);
    if (res && res.maintenance) {
      logWarning("📧 Bloccata da manutenzione: " + to + " — " + subject);
      return;
    }
    logInfo("📧 Email inviata a " + to);
  } catch (e) {
    logError("❌ Errore nell'invio email a " + to + ": " + e.message);
    addToEmailQueue(to, subject, body);
  }
}
