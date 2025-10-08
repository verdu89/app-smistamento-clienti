/** Dashboard & Reports
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */


function countLeadsByCalendar(start, end, rawData, colsMap, normalizeFn) {
    // Conta direttamente sulle righe (più preciso per range arbitrari)
    let tot = 0;
    for (let i = 1; i < rawData.length; i++) {
      const d = normalizeFn(getVal(rawData[i], "Data e ora"));
      if (d >= stripTime(start) && d <= stripTime(end)) tot++;
    }
    return tot;
  }


function countLeadsInRange(start, end, map, isYearOrMonthMap) {
    // Se isYearOrMonthMap == true, somma da map (monthMapLead/yearMapLead) su chiavi comprese,
    // altrimenti non usato qui.
    let tot = 0;
    if (isYearOrMonthMap) {
      // Usiamo i dati di data per sicurezza (le mappe non sono per giorno)
      // Qui delego alla funzione che scorre il dataset puntualmente:
      return countLeadsByCalendar(start, end, data, cols, normalizzaData);
    }
    return tot;
  }


function rangeCountLeads(start, end, periodMap /* weekMapLead|Vend */) {
    // Conta sommando chiavi del map che rientrano nel range
    let tot = 0;
    if (periodMap === weekMapLead || periodMap === weekMapVend) {
      // Settimane
      for (const k of Object.keys(periodMap)) {
        const [yy, ww] = k.split("-").map(Number);
        const dateFromKey = weekKeyToDate(yy, ww);
        if (dateInRange(dateFromKey, start, end)) tot += periodMap[k];
      }
    } else {
      // Non usato qui, ma lasciato per simmetria
      for (const k of Object.keys(periodMap)) tot += periodMap[k];
    }
    return tot;
  }


function sendWeeklyReport() {
  aggiornaNumeroPezziInMain(); // ✅ aggiorna campi mancanti

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var data = mainSheet.getDataRange().getValues();
  var colsMain = getColumnIndexes(data[0]);

  var thisMonday = getLastMonday();
  var startDate = new Date(thisMonday);
  startDate.setDate(startDate.getDate() - 7);
  var endDate = new Date(thisMonday);
  endDate.setDate(endDate.getDate() - 1);

  var weekNumber = getWeekNumber(startDate); // 🔢 settimana dei preventivi

  // Imposta orari precisi
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999);

  var clients = [];
  var totalPezzi = 0;

  // 🔢 mappa totali per venditore
  var vendorTotals = {}; // { [venditore]: { pezzi: number, clienti: number } }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateCell = row[colsMain["Data e ora"]];
    if (!dateCell) {
      logInfo(`⚠️ Riga ${i + 1}: campo "Data e ora" vuoto`);
      continue;
    }

    var assignedDate = tryParseDate(dateCell);
    Logger.log(
      `🔍 Riga ${i + 1} – Valore raw: "${dateCell}" ➝ Parsed: ${assignedDate}`
    );

    if (!(assignedDate instanceof Date) || isNaN(assignedDate)) {
      logInfo(`⚠️ Riga ${i + 1}: data non valida -> "${dateCell}"`);
      continue;
    }

    // ✅ Clona l'oggetto Date per azzerare l'orario
    var assignedDateMidnight = new Date(assignedDate);
    assignedDateMidnight.setHours(0, 0, 0, 0);

    if (assignedDateMidnight >= startDate && assignedDateMidnight <= endDate) {
      clients.push(row);

      // ➕ accumula totali per venditore
      var venditore = (row[colsMain["Venditore Assegnato"]] || "-")
        .toString()
        .trim();
      var pezzi = parseInt(row[colsMain["Numero pezzi"]]) || 0;
      totalPezzi += pezzi;

      if (!vendorTotals[venditore]) {
        vendorTotals[venditore] = { pezzi: 0, clienti: 0 };
      }
      vendorTotals[venditore].pezzi += pezzi;
      vendorTotals[venditore].clienti += 1;
    }
  }

  if (clients.length === 0) {
    logInfo("📌 Nessun nuovo cliente registrato la settimana scorsa.");
    return;
  }

  // 🧾 tabella dettagli clienti
  var emailBody = `
  <div style="font-family: Arial; max-width: 800px; margin: auto;">
    <h2 style="text-align:center;">📊 Riepilogo Nuovi Clienti della Settimana</h2>
    <p>🗓️ Settimana <b>#${weekNumber}</b> – dal <b>${startDate.toLocaleDateString()}</b> al <b>${endDate.toLocaleDateString()}</b></p>

    <table border="1" style="border-collapse: collapse; width: 100%; font-size: 12px;">
      <thead style="background-color: #f2f2f2;">
        <tr>
          <th>Data</th>
          <th>Nome</th>
          <th>Telefono</th>
          <th>Email</th>
          <th>Luogo di Consegna</th>
          <th>Venditore Assegnato</th>
          <th>Numero pezzi</th>
          <th>Provenienza contatto</th>
        </tr>
      </thead>
      <tbody>`;

  clients.forEach(function (c) {
    var dataOra = tryParseDate(c[colsMain["Data e ora"]]);
    var dataFormattata = dataOra ? dataOra.toLocaleDateString() : "-";
    var pezzi = parseInt(c[colsMain["Numero pezzi"]]) || 0;

    emailBody += `
      <tr>
        <td>${dataFormattata}</td>
        <td>${c[colsMain["Nome"]] || "-"}</td>
        <td>${c[colsMain["Telefono"]] || "-"}</td>
        <td>${c[colsMain["Email"]] || "-"}</td>
        <td>${c[colsMain["Luogo di Consegna"]] || "-"}</td>
        <td>${c[colsMain["Venditore Assegnato"]] || "-"}</td>
        <td style="text-align:center;">${pezzi}</td>
        <td>${c[colsMain["Provenienza contatto"]] || "Internet"}</td>
      </tr>`;
  });

  emailBody += `
      </tbody>
    </table>

    <br>
    <h3 style="text-align:right; margin: 0;">Totale pezzi richiesti: ${totalPezzi}</h3>
    <hr style="margin: 18px 0; border: none; border-top: 1px solid #ddd;">

    <!-- 🧮 RIEPILOGO PER VENDITORE -->
    <h3 style="margin: 8px 0 6px 0;">🧮 Riepilogo pezzi per venditore</h3>
    <table border="1" style="border-collapse: collapse; width: 100%; font-size: 12px;">
      <thead style="background-color: #f9f9f9;">
        <tr>
          <th>Venditore</th>
          <th>Clienti</th>
          <th>Pezzi</th>
        </tr>
      </thead>
      <tbody>`;

  // Ordina venditori per pezzi desc
  Object.keys(vendorTotals)
    .sort(function (a, b) {
      return vendorTotals[b].pezzi - vendorTotals[a].pezzi;
    })
    .forEach(function (v) {
      emailBody += `
        <tr>
          <td>${v}</td>
          <td style="text-align:center;">${vendorTotals[v].clienti}</td>
          <td style="text-align:center;"><b>${vendorTotals[v].pezzi}</b></td>
        </tr>`;
      Logger.log(
        `👤 ${v}: ${vendorTotals[v].clienti} clienti, ${vendorTotals[v].pezzi} pezzi`
      );
    });

  emailBody += `
      </tbody>
    </table>

    <p style="font-size: 10px; text-align: center; margin-top: 30px;">Impaginato per stampa su foglio A4</p>
  </div>`;

  sendEmail(
    "newsaverplast@gmail.com",
    "📊 [Riepilogo settimanale] Nuovi Clienti",
    emailBody
  );

  logInfo(
    `✅ Report inviato: ${clients.length} clienti, ${totalPezzi} pezzi totali.`
  );
}


function updateDashboardFromMain() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  sheet.clear();

  const mainSheet = ss.getSheetByName("Main");
  const data = mainSheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    Logger.log("Nessun dato.");
    return;
  }

  const headers = data[0];
  const cols = getColumnIndexes(headers);
  const today = stripTime(new Date());

  /* ==========================
   *  Funzioni supporto locali
   * ========================== */
  function parseCustomDate(val) {
    if (val instanceof Date && !isNaN(val)) return stripTime(val);
    if (typeof val === "string" && val.trim() !== "") {
      const parsed = new Date(val);
      if (!isNaN(parsed)) return stripTime(parsed);
    }
    return null;
  }
  function stripTime(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  function normalizzaData(val) {
    // Se mancante/illeggibile, assegna la data di oggi per coerenza dei KPI temporali
    return parseCustomDate(val) || stripTime(today);
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
  function getWeekNumber(d) {
    const _d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    _d.setUTCDate(_d.getUTCDate() + 4 - (_d.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(_d.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil(((_d - yearStart) / 86400000 + 1) / 7);
    return weekNo;
  }
  function getLastMonday(fromDate) {
    const d = new Date(fromDate || new Date());
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    return stripTime(new Date(d.setDate(diff)));
  }
  function getVal(row, key) {
    const idx = cols[key];
    return typeof idx === "number" && idx >= 0 ? row[idx] : "";
  }
  function fmtPerc(n) {
    return isFinite(n) ? (n * 100).toFixed(1) + "%" : "-";
  }
  function safeDiv(a, b) {
    return b > 0 ? a / b : 0;
  }

  /* ==========================
   *  Variabili accumulo
   * ========================== */
  let leadTotali = 0,
    preventivi = 0,
    vendite = 0,
    totalePezzi = 0;
  let tempoTotaleRisposta = 0,
    risposteValide = 0;

  const provenienze = {}; // { fonte: count }
  const statoDistribuzione = {}; // { stato: count }
  const venditori = {}; // { nome: {lead, preventivi, vendite, tempi[] } }
  const venditeProvenienza = {}; // { fonte: vendite }

  // Raggruppamenti per trend & medie
  const weekMapLead = {}; // { "YYYY-WW": count }
  const weekMapVend = {}; // { "YYYY-WW": count }
  const monthMapLead = {}; // { "YYYY-MM": count }
  const monthMapVend = {}; // { "YYYY-MM": count }
  const yearMapLead = {}; // { "YYYY": count }
  const yearMapVend = {}; // { "YYYY": count }

  // Periodi chiave
  const inizioAnno = new Date(today.getFullYear(), 0, 1);
  const inizioMese = new Date(today.getFullYear(), today.getMonth(), 1);
  const thisMon = getLastMonday(today);
  const prevMon = new Date(thisMon);
  prevMon.setDate(prevMon.getDate() - 7);
  const prevSun = new Date(prevMon);
  prevSun.setDate(prevMon.getDate() + 6);

  // Confronti (settimana precedente alla settimana chiusa)
  const prevPrevMon = new Date(prevMon);
  prevPrevMon.setDate(prevMon.getDate() - 7);
  const prevPrevSun = new Date(prevPrevMon);
  prevPrevSun.setDate(prevPrevMon.getDate() + 6);

  let leadAnno = 0,
    leadMese = 0,
    leadSettimanali = 0;
  let venditeSettimanali = 0;
  let primaDataLead = null;

  /* ==========================
   *  Ciclo dati – normalizzazione e conteggi
   * ========================== */
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const stato =
      (getVal(row, "Stato") || "").toString().trim() || "Non specificato";
    const venditore = (getVal(row, "Venditore Assegnato") || "")
      .toString()
      .trim();
    const dataAssegnazione = normalizzaData(getVal(row, "Data e ora")); // <= NORMALIZZAZIONE
    const dataPreventivo = parseCustomDate(getVal(row, "Data Preventivo")); // non forziamo default per non falsare i tempi
    const vendConclusaStr = (getVal(row, "Vendita Conclusa?") || "")
      .toString()
      .trim()
      .toUpperCase();
    const pezzi = parseInt(getVal(row, "Numero pezzi")) || 0;
    const provenienza = normalizzaProvenienza(
      getVal(row, "Provenienza contatto") || "Internet"
    );

    const isVendita =
      vendConclusaStr === "SI" ||
      (vendConclusaStr === "" && stato === "Trattativa terminata");

    // KPI globali
    leadTotali++;
    totalePezzi += pezzi;

    provenienze[provenienza] = (provenienze[provenienza] || 0) + 1;
    statoDistribuzione[stato] = (statoDistribuzione[stato] || 0) + 1;

    if (!primaDataLead || dataAssegnazione < primaDataLead)
      primaDataLead = dataAssegnazione;
    if (dataAssegnazione >= inizioAnno) leadAnno++;
    if (dataAssegnazione >= inizioMese) leadMese++;
    if (dataAssegnazione >= prevMon && dataAssegnazione <= prevSun)
      leadSettimanali++;

    if (stato === "Preventivo inviato") preventivi++;
    if (isVendita) {
      vendite++;
      venditeProvenienza[provenienza] =
        (venditeProvenienza[provenienza] || 0) + 1;
      if (dataAssegnazione >= prevMon && dataAssegnazione <= prevSun)
        venditeSettimanali++;
    }

    // Venditori (contiamo sempre la riga)
    if (venditore) {
      if (!venditori[venditore])
        venditori[venditore] = {
          lead: 0,
          preventivi: 0,
          vendite: 0,
          tempi: [],
        };
      venditori[venditore].lead++;
      if (stato === "Preventivo inviato") venditori[venditore].preventivi++;
      if (isVendita) venditori[venditore].vendite++;
      if (dataPreventivo) {
        const diff = Math.round(
          (dataPreventivo - dataAssegnazione) / (1000 * 60 * 60 * 24)
        );
        if (!isNaN(diff) && diff >= 0 && diff <= 60) {
          venditori[venditore].tempi.push(diff);
          tempoTotaleRisposta += diff;
          risposteValide++;
        }
      }
    }

    // Raggruppamenti per medie e trend
    const w = getWeekNumber(dataAssegnazione);
    const y = dataAssegnazione.getFullYear();
    const m = dataAssegnazione.getMonth() + 1;
    const weekKey = `${y}-${String(w).padStart(2, "0")}`;
    const monthKey = `${y}-${String(m).padStart(2, "0")}`;
    const yearKey = `${y}`;

    weekMapLead[weekKey] = (weekMapLead[weekKey] || 0) + 1;
    monthMapLead[monthKey] = (monthMapLead[monthKey] || 0) + 1;
    yearMapLead[yearKey] = (yearMapLead[yearKey] || 0) + 1;

    if (isVendita) {
      weekMapVend[weekKey] = (weekMapVend[weekKey] || 0) + 1;
      monthMapVend[monthKey] = (monthMapVend[monthKey] || 0) + 1;
      yearMapVend[yearKey] = (yearMapVend[yearKey] || 0) + 1;
    }
  }

  /* ==========================
   *  KPI Globali
   * ========================== */
  const conversionRate =
    preventivi > 0 ? ((vendite / preventivi) * 100).toFixed(1) + "%" : "-";
  const pezziMedi =
    leadTotali > 0 ? (totalePezzi / leadTotali).toFixed(2) : "-";
  const tempoMedioRisposta =
    risposteValide > 0 ? Math.round(tempoTotaleRisposta / risposteValide) : "-";

  // Medie arrivi lead – giornaliere/settimanali/mensili/annuali
  const giorniStorici = primaDataLead
    ? getWorkingDaysInRange(primaDataLead, today)
    : 1;
  const giorniAnno = getWorkingDaysInRange(inizioAnno, today);
  const giorniMese = getWorkingDaysInRange(inizioMese, today);

  const mediaGiornalieraStorica = (leadTotali / giorniStorici).toFixed(2);
  const mediaGiornalieraAnno = (leadAnno / giorniAnno).toFixed(2);
  const mediaGiornalieraMese = (leadMese / giorniMese).toFixed(2);
  const mediaGiornalieraSettimana = (leadSettimanali / 5).toFixed(2);

  const nSettimaneStoriche = Object.keys(weekMapLead).length || 1;
  const nMesiStorici = Object.keys(monthMapLead).length || 1;
  const nAnniStorici = Object.keys(yearMapLead).length || 1;

  const mediaSettimanaleStorica = (leadTotali / nSettimaneStoriche).toFixed(2);
  const mediaMensileStorica = (leadTotali / nMesiStorici).toFixed(2);
  const mediaAnnualeStorica = (leadTotali / nAnniStorici).toFixed(2);

  // Confronti vs periodo precedente
  // Settimana chiusa vs settimana precedente
  const settLeadPrevPrev = rangeCountLeads(
    prevPrevMon,
    prevPrevSun,
    weekMapLead
  );
  const settVendPrevPrev = rangeCountLeads(
    prevPrevMon,
    prevPrevSun,
    weekMapVend
  );
  const settConvPrev =
    leadSettimanali > 0 ? venditeSettimanali / leadSettimanali : 0;
  const settConvPrevPrev =
    settLeadPrevPrev > 0 ? settVendPrevPrev / settLeadPrevPrev : 0;

  // Mese (MTD) vs mese precedente allo stesso numero di giorni lavorativi
  const prevMonth = new Date(inizioMese);
  prevMonth.setMonth(prevMonth.getMonth() - 1);
  const endMTD = today; // fino ad oggi
  const giorniLavorativiMTD = getWorkingDaysInRange(inizioMese, endMTD);
  const endPrevMonthAligned = alignPrevPeriodEnd(
    prevMonth,
    giorniLavorativiMTD
  );
  const leadsMTD = countLeadsInRange(inizioMese, endMTD, monthMapLead, true);
  const leadsPrevAligned = countLeadsByCalendar(
    prevMonth,
    endPrevMonthAligned,
    data,
    cols,
    normalizzaData
  );

  // Anno (YTD) vs anno precedente allo stesso numero di giorni lavorativi
  const prevYear = new Date(inizioAnno);
  prevYear.setFullYear(prevYear.getFullYear() - 1);
  const giorniLavorativiYTD = getWorkingDaysInRange(inizioAnno, today);
  const endPrevYearAligned = alignPrevPeriodEnd(
    prevYear,
    giorniLavorativiYTD,
    "year"
  );
  const leadsYTD = countLeadsInRange(inizioAnno, today, yearMapLead, true);
  const leadsPrevYearAligned = countLeadsByCalendar(
    prevYear,
    endPrevYearAligned,
    data,
    cols,
    normalizzaData
  );

  /* ==========================
   *  Scrittura DASHBOARD
   * ========================== */
  // Titolo
  sheet
    .getRange("B1")
    .setValue("📊 DASHBOARD PREMIUM – LEAD & VENDITE (riunioni)")
    .setFontSize(18)
    .setFontWeight("bold");
  sheet
    .getRange("B2")
    .setValue(
      "Aggiornata al: " +
        Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy")
    )
    .setFontStyle("italic");
  sheet.appendRow([" "]);

  // KPI Globali
  sheet
    .getRange(sheet.getLastRow() + 1, 2)
    .setValue("📌 KPI Globali")
    .setFontWeight("bold");
  sheet.appendRow([" "]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 6)
    .setValues([
      [
        "Totale Lead",
        "Preventivi",
        "Vendite",
        "Conversion Rate",
        "Tempo medio risposta (gg)",
        "Pezzi medi/Lead",
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 6)
    .setValues([
      [
        leadTotali,
        preventivi,
        vendite,
        conversionRate,
        tempoMedioRisposta,
        pezziMedi,
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() - 1, 2, 1, 6)
    .setFontWeight("bold")
    .setBackground("#0b5394")
    .setFontColor("white")
    .setHorizontalAlignment("center");
  sheet
    .getRange(sheet.getLastRow(), 2, 1, 6)
    .setFontWeight("bold")
    .setBackground("#cfe2f3")
    .setHorizontalAlignment("center");
  sheet.appendRow([" "]);

  // Medie arrivi lead – giornaliere / settimanali / mensili / annuali
  sheet
    .getRange(sheet.getLastRow() + 1, 2)
    .setValue("🟡 Medie Arrivi Lead")
    .setFontWeight("bold");
  sheet.appendRow([" "]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 8)
    .setValues([
      [
        "Giornaliera Storica",
        "Giornaliera YTD",
        "Giornaliera MTD",
        "Giornaliera Sett. Chiusa",
        "Settimanale Storica",
        "Mensile Storica",
        "Annuale Storica",
        "Settimana Chiusa: Media/giorno",
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 8)
    .setValues([
      [
        mediaGiornalieraStorica,
        mediaGiornalieraAnno,
        mediaGiornalieraMese,
        mediaGiornalieraSettimana,
        mediaSettimanaleStorica,
        mediaMensileStorica,
        mediaAnnualeStorica,
        mediaGiornalieraSettimana,
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() - 1, 2, 1, 8)
    .setFontWeight("bold")
    .setBackground("#f1c232")
    .setHorizontalAlignment("center");
  sheet
    .getRange(sheet.getLastRow(), 2, 1, 8)
    .setFontWeight("bold")
    .setBackground("#fff2cc")
    .setHorizontalAlignment("center");
  sheet.appendRow([" "]);

  // Riepilogo settimana chiusa (prevMon - prevSun) + confronto
  sheet
    .getRange(sheet.getLastRow() + 1, 2)
    .setValue(
      `🟢 Riepilogo settimana chiusa (${fmtDate(prevMon)} - ${fmtDate(
        prevSun
      )})`
    )
    .setFontWeight("bold");
  sheet.appendRow([" "]);
  const convSett =
    leadSettimanali > 0
      ? ((venditeSettimanali / leadSettimanali) * 100).toFixed(1) + "%"
      : "-";
  const convSettPrev =
    settLeadPrevPrev > 0
      ? ((settVendPrevPrev / settLeadPrevPrev) * 100).toFixed(1) + "%"
      : "-";
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 6)
    .setValues([
      [
        "Lead (sett.)",
        "Vendite (sett.)",
        "Conv. (sett.)",
        "Lead sett. precedente",
        "Vendite sett. prec.",
        "Conv. sett. prec.",
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 6)
    .setValues([
      [
        leadSettimanali,
        venditeSettimanali,
        convSett,
        settLeadPrevPrev,
        settVendPrevPrev,
        convSettPrev,
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() - 1, 2, 1, 6)
    .setFontWeight("bold")
    .setBackground("#38761d")
    .setFontColor("white")
    .setHorizontalAlignment("center");
  sheet
    .getRange(sheet.getLastRow(), 2, 1, 6)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");
  sheet.appendRow([" "]);

  // Analisi Lead per periodo
  sheet
    .getRange(sheet.getLastRow() + 1, 2)
    .setValue("🟠 Analisi Lead per periodo")
    .setFontWeight("bold");
  sheet.appendRow([" "]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 5)
    .setValues([
      ["Periodo", "Lead", "Vendite", "Conversion Rate", "Media Lead/giorno"],
    ]);

  const convStorico =
    leadTotali > 0 ? ((vendite / leadTotali) * 100).toFixed(1) + "%" : "-";
  const mediaStorica = (leadTotali / giorniStorici).toFixed(2);
  const convAnno =
    leadAnno > 0
      ? (
          (countInMap(yearMapVend, String(today.getFullYear())) / leadAnno) *
          100
        ).toFixed(1) + "%"
      : "-";
  const convMese =
    leadMese > 0
      ? (
          (countInMap(
            monthMapVend,
            `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(
              2,
              "0"
            )}`
          ) /
            leadMese) *
          100
        ).toFixed(1) + "%"
      : "-";
  const convSettimanale =
    leadSettimanali > 0
      ? ((venditeSettimanali / leadSettimanali) * 100).toFixed(1) + "%"
      : "-";

  const mediaAnno = (leadAnno / giorniAnno).toFixed(2);
  const mediaMese = (leadMese / giorniMese).toFixed(2);
  const mediaSett = (leadSettimanali / 5).toFixed(2);

  sheet.appendRow(["Storico", leadTotali, vendite, convStorico, mediaStorica]);
  sheet.appendRow([
    "Anno (YTD)",
    leadAnno,
    countInMap(yearMapVend, String(today.getFullYear())),
    convAnno,
    mediaAnno,
  ]);
  sheet.appendRow([
    "Mese (MTD)",
    leadMese,
    countInMap(
      monthMapVend,
      `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`
    ),
    convMese,
    mediaMese,
  ]);
  sheet.appendRow([
    "Settimana chiusa",
    leadSettimanali,
    venditeSettimanali,
    convSettimanale,
    mediaSett,
  ]);

  const leadStart = sheet.getLastRow() - 3;
  sheet
    .getRange(leadStart - 1, 2, 1, 5)
    .setFontWeight("bold")
    .setBackground("#e69138")
    .setFontColor("white")
    .setHorizontalAlignment("center");
  sheet
    .getRange(leadStart, 2, 4, 5)
    .setBackground("#fce5cd")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
  sheet.appendRow([" "]);

  // Confronti MTD e YTD (lead)
  sheet
    .getRange(sheet.getLastRow() + 1, 2)
    .setValue("🔵 Confronti Lead (MTD/YTD vs periodo allineato)")
    .setFontWeight("bold");
  sheet.appendRow([" "]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 5)
    .setValues([
      [
        "Periodo",
        "Lead attuali",
        "Lead periodo prec. allineato",
        "Δ Lead",
        "Δ %",
      ],
    ]);
  const deltaMTD = leadsMTD - leadsPrevAligned;
  const deltaMTDPerc = fmtPerc(safeDiv(deltaMTD, leadsPrevAligned));
  const deltaYTD = leadsYTD - leadsPrevYearAligned;
  const deltaYTDPerc = fmtPerc(safeDiv(deltaYTD, leadsPrevYearAligned));
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 5)
    .setValues([["MTD", leadsMTD, leadsPrevAligned, deltaMTD, deltaMTDPerc]]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 5)
    .setValues([
      ["YTD", leadsYTD, leadsPrevYearAligned, deltaYTD, deltaYTDPerc],
    ]);
  sheet
    .getRange(sheet.getLastRow() - 2, 2, 2, 5)
    .setBackground("#d9e1f2")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
  sheet
    .getRange(sheet.getLastRow() - 3, 2, 1, 5)
    .setFontWeight("bold")
    .setBackground("#2e75b6")
    .setFontColor("white")
    .setHorizontalAlignment("center");
  sheet.appendRow([" "]);

  // Performance Venditori
  sheet
    .getRange(sheet.getLastRow() + 1, 2)
    .setValue("⚫ Performance Venditori")
    .setFontWeight("bold");
  sheet.appendRow([" "]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 7)
    .setValues([
      [
        "Venditore",
        "Lead gestite",
        "Preventivi",
        "Vendite",
        "% Chiusura su Lead",
        "% Chiusura su Preventivi",
        "Tempo medio risposta (gg)",
      ],
    ]);

  const vendArray = Object.keys(venditori)
    .map((nome) => {
      const v = venditori[nome];
      const chiusuraLead = v.lead > 0 ? v.vendite / v.lead : 0;
      const chiusuraPrev = v.preventivi > 0 ? v.vendite / v.preventivi : 0;
      const tempoMedio =
        v.tempi.length > 0
          ? Math.round(v.tempi.reduce((a, b) => a + b, 0) / v.tempi.length)
          : "-";
      return {
        nome,
        lead: v.lead,
        preventivi: v.preventivi,
        vendite: v.vendite,
        chiusuraLead,
        chiusuraPrev,
        chiusuraLeadTxt: fmtPerc(chiusuraLead),
        chiusuraPrevTxt: fmtPerc(chiusuraPrev),
        tempoMedio,
      };
    })
    .sort((a, b) => b.chiusuraLead - a.chiusuraLead);

  vendArray.forEach((v, idx) => {
    let nome = v.nome;
    if (idx === 0) nome = "🥇 " + nome;
    else if (idx === 1) nome = "🥈 " + nome;
    else if (idx === 2) nome = "🥉 " + nome;
    sheet.appendRow([
      nome,
      v.lead,
      v.preventivi,
      v.vendite,
      v.chiusuraLeadTxt,
      v.chiusuraPrevTxt,
      v.tempoMedio,
    ]);
  });

  const vendStart = sheet.getLastRow() - vendArray.length;
  sheet
    .getRange(vendStart - 1, 2, 1, 7)
    .setFontWeight("bold")
    .setBackground("#666666")
    .setFontColor("white")
    .setHorizontalAlignment("center");
  if (vendArray.length > 0)
    sheet
      .getRange(vendStart, 2, vendArray.length, 7)
      .setBackground("#f3f3f3")
      .setHorizontalAlignment("center");
  sheet.appendRow([" "]);

  /* ==========================
   *  Grafici
   * ========================== */
  let chartStart = sheet.getLastRow() + 5;
  sheet
    .getRange(chartStart, 2)
    .setValue("📈 Analisi Grafica")
    .setFontWeight("bold")
    .setFontSize(14);

  // 1) Distribuzione per Stato (PIE)
  const statoRow = chartStart + 2;
  sheet
    .getRange(statoRow, 2)
    .setValue("📊 Distribuzione per Stato")
    .setFontWeight("bold");
  const statoKeys = Object.keys(statoDistribuzione);
  statoKeys.forEach((k, i) => {
    sheet.getRange(statoRow + i + 1, 2).setValue(k);
    sheet.getRange(statoRow + i + 1, 3).setValue(statoDistribuzione[k]);
  });
  if (statoKeys.length > 0) {
    const chart1 = sheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(statoRow + 1, 2, statoKeys.length, 2))
      .setPosition(statoRow, 6, 0, 0)
      .setOption("title", "Distribuzione per Stato")
      .build();
    sheet.insertChart(chart1);
  }

  // 2) Provenienza Lead (PIE)
  const provRow = statoRow + statoKeys.length + 8;
  sheet
    .getRange(provRow, 2)
    .setValue("📊 Provenienza Lead")
    .setFontWeight("bold");
  const provKeys = Object.keys(provenienze);
  provKeys.forEach((k, i) => {
    sheet.getRange(provRow + i + 1, 2).setValue(k);
    sheet.getRange(provRow + i + 1, 3).setValue(provenienze[k]);
  });
  if (provKeys.length > 0) {
    const chart2 = sheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(provRow + 1, 2, provKeys.length, 2))
      .setPosition(provRow, 6, 0, 0)
      .setOption("title", "Provenienza Lead")
      .build();
    sheet.insertChart(chart2);
  }

  // 3) Vendite per Provenienza (BAR)
  const vendRow = provRow + provKeys.length + 8;
  sheet
    .getRange(vendRow, 2)
    .setValue("📈 Vendite per Provenienza")
    .setFontWeight("bold");
  const vendKeys = Object.keys(venditeProvenienza);
  vendKeys.forEach((k, i) => {
    sheet.getRange(vendRow + i + 1, 2).setValue(k);
    sheet.getRange(vendRow + i + 1, 3).setValue(venditeProvenienza[k]);
  });
  if (vendKeys.length > 0) {
    const chart3 = sheet
      .newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(vendRow + 1, 2, vendKeys.length, 2))
      .setPosition(vendRow, 6, 0, 0)
      .setOption("title", "Vendite per Provenienza")
      .build();
    sheet.insertChart(chart3);
  }

  // 4) Trend ultime 12 settimane (COLUMN, Lead + Vendite)
  const trendRow = vendRow + vendKeys.length + 10;
  sheet
    .getRange(trendRow, 2)
    .setValue("📉 Trend – Ultime 12 settimane")
    .setFontWeight("bold");

  const sortedWeeks = Object.keys(weekMapLead).sort(sortPeriodKeys);
  const last12Weeks = sortedWeeks.slice(-12);
  const tableWeekRow = trendRow + 2;
  sheet
    .getRange(tableWeekRow, 2, 1, 3)
    .setValues([["Settimana", "Lead", "Vendite"]]);
  last12Weeks.forEach((wk, i) => {
    sheet
      .getRange(tableWeekRow + i + 1, 2, 1, 3)
      .setValues([[wk, weekMapLead[wk] || 0, weekMapVend[wk] || 0]]);
  });
  sheet
    .getRange(tableWeekRow, 2, 1, 3)
    .setFontWeight("bold")
    .setBackground("#cfe2f3");
  if (last12Weeks.length > 0) {
    const chart4 = sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(tableWeekRow, 2, last12Weeks.length + 1, 3))
      .setPosition(tableWeekRow, 6, 0, 0)
      .setOption("title", "Lead & Vendite – Ultime 12 settimane")
      .build();
    sheet.insertChart(chart4);
  }

  // 5) Trend ultimi 12 mesi (COLUMN, Lead + Vendite)
  const monthTrendRow = tableWeekRow + last12Weeks.length + 10;
  sheet
    .getRange(monthTrendRow, 2)
    .setValue("📉 Trend – Ultimi 12 mesi")
    .setFontWeight("bold");

  const sortedMonths = Object.keys(monthMapLead).sort(sortPeriodKeys);
  const last12Months = sortedMonths.slice(-12);
  const tableMonthRow = monthTrendRow + 2;
  sheet
    .getRange(tableMonthRow, 2, 1, 3)
    .setValues([["Mese", "Lead", "Vendite"]]);
  last12Months.forEach((mk, i) => {
    sheet
      .getRange(tableMonthRow + i + 1, 2, 1, 3)
      .setValues([[mk, monthMapLead[mk] || 0, monthMapVend[mk] || 0]]);
  });
  sheet
    .getRange(tableMonthRow, 2, 1, 3)
    .setFontWeight("bold")
    .setBackground("#d9ead3");
  if (last12Months.length > 0) {
    const chart5 = sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(tableMonthRow, 2, last12Months.length + 1, 3))
      .setPosition(tableMonthRow, 6, 0, 0)
      .setOption("title", "Lead & Vendite – Ultimi 12 mesi")
      .build();
    sheet.insertChart(chart5);
  }

  // Executive Summary
  const summaryRow = tableMonthRow + last12Months.length + 10;
  sheet
    .getRange(summaryRow, 2)
    .setValue("🟣 Executive Summary")
    .setFontWeight("bold")
    .setFontSize(12);
  sheet.appendRow([" "]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 4)
    .setValues([["Periodo", "Lead", "Vendite", "Conversion Rate"]]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 4)
    .setValues([["Storico", leadTotali, vendite, convStorico]]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 4)
    .setValues([
      [
        "YTD",
        leadAnno,
        countInMap(yearMapVend, String(today.getFullYear())),
        convAnno,
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 4)
    .setValues([
      [
        "MTD",
        leadMese,
        countInMap(
          monthMapVend,
          `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(
            2,
            "0"
          )}`
        ),
        convMese,
      ],
    ]);
  sheet
    .getRange(sheet.getLastRow() + 1, 2, 1, 4)
    .setValues([
      [
        "Settimana chiusa",
        leadSettimanali,
        venditeSettimanali,
        convSettimanale,
      ],
    ]);
  const sumStart = sheet.getLastRow() - 3;
  sheet
    .getRange(sumStart - 1, 2, 1, 4)
    .setFontWeight("bold")
    .setBackground("#674ea7")
    .setFontColor("white")
    .setHorizontalAlignment("center");
  sheet
    .getRange(sumStart, 2, 4, 4)
    .setBackground("#d9d2e9")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");

  logInfo("✅ Dashboard Premium aggiornata con successo");

  /* ==========================
   *  Funzioni di appoggio interne (con accesso a variabili locali)
   * ========================== */
  function fmtDate(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  function sortPeriodKeys(a, b) {
    // Ordina "YYYY-XX" correttamente
    const [ya, xa] = a.split("-").map(Number);
    const [yb, xb] = b.split("-").map(Number);
    return ya === yb ? xa - xb : ya - yb;
  }
  function countInMap(map, key) {
    return map[key] || 0;
  }
  function dateInRange(d, start, end) {
    return d >= stripTime(start) && d <= stripTime(end);
  }
  function rangeCountLeads(start, end, periodMap /* weekMapLead|Vend */) {
    // Conta sommando chiavi del map che rientrano nel range
    let tot = 0;
    if (periodMap === weekMapLead || periodMap === weekMapVend) {
      // Settimane
      for (const k of Object.keys(periodMap)) {
        const [yy, ww] = k.split("-").map(Number);
        const dateFromKey = weekKeyToDate(yy, ww);
        if (dateInRange(dateFromKey, start, end)) tot += periodMap[k];
      }
    } else {
      // Non usato qui, ma lasciato per simmetria
      for (const k of Object.keys(periodMap)) tot += periodMap[k];
    }
    return tot;
  }
  function weekKeyToDate(year, week) {
    // Ritorna il lunedì di quella settimana ISO
    const simple = new Date(year, 0, 1 + (week - 1) * 7);
    const dow = simple.getDay();
    const ISOweekStart = new Date(simple);
    const diff =
      dow <= 4 ? simple.getDate() - dow + 1 : simple.getDate() + 8 - dow;
    ISOweekStart.setDate(diff);
    return stripTime(ISOweekStart);
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
  function countLeadsInRange(start, end, map, isYearOrMonthMap) {
    // Se isYearOrMonthMap == true, somma da map (monthMapLead/yearMapLead) su chiavi comprese,
    // altrimenti non usato qui.
    let tot = 0;
    if (isYearOrMonthMap) {
      // Usiamo i dati di data per sicurezza (le mappe non sono per giorno)
      // Qui delego alla funzione che scorre il dataset puntualmente:
      return countLeadsByCalendar(start, end, data, cols, normalizzaData);
    }
    return tot;
  }
  function countLeadsByCalendar(start, end, rawData, colsMap, normalizeFn) {
    // Conta direttamente sulle righe (più preciso per range arbitrari)
    let tot = 0;
    for (let i = 1; i < rawData.length; i++) {
      const d = normalizeFn(getVal(rawData[i], "Data e ora"));
      if (d >= stripTime(start) && d <= stripTime(end)) tot++;
    }
    return tot;
  }
}
