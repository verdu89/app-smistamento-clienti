/** Utilities
 * Auto-generated split from smistamento-clienti.gs
 * Keep functions unchanged; moved only for organization.
 */


// === Canonical helpers (deduped) ===
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
    if (keyN) normMap.set(keyN, val);
  });

  // 1) match esatto su exactNormalized (se fornito)
  if (options && options.exactNormalized) {
    for (const ex of options.exactNormalized) {
      if (normMap.has(ex)) return normMap.get(ex);
    }
  }

  // 2) fallback: include tutte le parole chiave indicate
  const must = (options && options.mustInclude ? options.mustInclude : []).map(normalize);
  for (const fd of fieldData) {
    const n = normalize(fd.name || "");
    if (must.length > 0 && must.every((tok) => n.includes(tok))) {
      return (fd.values || []).join(", ");
    }
  }
  return "";
}

function getLastMonday(fromDate) {
  const d = new Date(fromDate || new Date());
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return stripTime(new Date(d.setDate(diff)));
}

function getWeekNumber(d) {
  const _d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  _d.setUTCDate(_d.getUTCDate() + 4 - (_d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(_d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil(((_d - yearStart) / 86400000 + 1) / 7);
  return weekNo;
}

function md5Hex_(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    str,
    Utilities.Charset.UTF_8
  );
  let hex = "";
  for (let i = 0; i < bytes.length; i++) {
    let h = (bytes[i] & 0xff).toString(16);
    if (h.length === 1) h = "0" + h;
    hex += h;
  }
  return hex;
}

// === Moved utility helpers (paste originals below, unchanged) ===


function applyDropdownIfColumnExists(sheet, columnName, values, colors = null) {
  var headers = sheet.getDataRange().getValues()[0]; // Legge le intestazioni
  var colIndex = headers.indexOf(columnName); // Trova la posizione della colonna

  if (colIndex === -1) {
    Logger.log(`⚠️ La colonna "${columnName}" non esiste nel foglio.`);
    return;
  }

  colIndex += 1; // Converti l'indice da 0-based a 1-based per Google Sheets

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // Nessun dato nel foglio oltre le intestazioni

  var range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);

  if (colors) {
    for (var i = 2; i <= lastRow; i++) {
      var cell = sheet.getRange(i, colIndex);
      var value = cell.getValue().toString().trim();
      if (value in colors) {
        cell.setBackground(colors[value]);
      } else {
        cell.setBackground("#FFFFFF");
      }
    }
  }

  Logger.log(
    `✅ Menu a tendina applicato per "${columnName}" alla colonna ${colIndex}.`
  );
}


function applyDropdownValidation(
  sheet,
  colIndex,
  values,
  colors = null,
  lastRow = null
) {
  if (lastRow === null) {
    lastRow = sheet.getLastRow();
  }
  if (lastRow < 2) return;

  var range = sheet.getRange(lastRow, colIndex + 1, 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);

  if (colors) {
    var cellValue = range.getValue().toString().trim();
    if (cellValue in colors) {
      range.setBackground(colors[cellValue]);
    } else {
      range.setBackground("#FFFFFF");
    }
  }
}


function applyUpdates(sheet, updates, colsMain) {
  updates.forEach((update) => {
    Logger.log(
      "✍️ Scrittura venditore: " + update[1] + " sulla riga " + update[0]
    );

    sheet
      .getRange(update[0], colsMain["Venditore Assegnato"] + 1)
      .setValue(update[1]);
    sheet.getRange(update[0], colsMain["Data e ora"] + 1).setValue(update[2]);
  });
}


function autoResizeAllColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  for (let c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
  }
}


function countInMap(map, key) {
    return map[key] || 0;
  }


function dateInRange(d, start, end) {
    return d >= stripTime(start) && d <= stripTime(end);
  }


function extractEmailFromText_(t) {
  if (!t) return "";
  const m = t.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0] : "";
}


function extractPhones_(t) {
  if (!t) return [];
  // Estrai numeri italiani plausibili (togli spazi/punti; accetta +39 opzionale)
  const cleaned = t.replace(/[\.\-\(\)]/g, " ");
  const digits = cleaned.match(/(\+?\s*39)?\s*\d{6,}/g); // >= 6 cifre utili
  if (!digits) return [];
  return digits.map((x) => x.replace(/\D+/g, ""));
}


function fmtDate(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }


function fmtPerc(n) {
    return isFinite(n) ? (n * 100).toFixed(1) + "%" : "-";
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


function getFirstEmptyRow(sheet) {
  if (!sheet) {
    Logger.log("❌ ERRORE: Il foglio non è definito in getFirstEmptyRow.");
    return 2; // Ritorna la riga 2 come fallback
  }

  var lastRow = sheet.getLastRow(); // Ottiene l'ultima riga "usata"

  if (lastRow === 0) return 2; // Se non ci sono dati, ritorna la prima riga

  var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i].every((cell) => cell === "")) {
      return i + 1; // Trova la prima riga vuota effettiva
    }
  }

  return lastRow + 1; // Se non ci sono righe completamente vuote, ritorna la successiva
}


function getVal(row, key) {
    const idx = cols[key];
    return typeof idx === "number" && idx >= 0 ? row[idx] : "";
  }


function isSimilarText_(a, b) {
  // similarità “leggera”: uno contiene l’altro per almeno 30 caratteri
  if (!a || !b) return false;
  if (a === b) return true;
  const minLen = 30;
  return (
    (a.length >= minLen && b.includes(a.slice(0, minLen))) ||
    (b.length >= minLen && a.includes(b.slice(0, minLen)))
  );
}


function isValidEmail_(email) {
  if (!email || typeof email !== "string") return false;
  const e = email.trim();
  // regex semplice e robusta per casi comuni
  return !!e.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/);
}


function normalizeBody_(text) {
  let t = text
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{2,}/g, "\n")
    .trim();

  // (opzionale) rimuovi footer tipo firma dopo "--"
  // t = t.replace(/--\s*\n.*$/s, "");

  return t;
}


function normalizePhone(tel) {
    if (!tel) return "";
    let clean = String(tel).replace(/\D/g, ""); // toglie tutto tranne cifre
    if (clean.startsWith("39") && clean.length > 10) {
      clean = clean.substring(2); // rimuove prefisso internazionale
    }
    return clean;
  }


function normalizePhone_(p) {
  if (!p) return "";
  let digits = p.replace(/\D+/g, "");
  // rimuovi prefisso 39 ripetuto
  if (digits.startsWith("39") && digits.length > 10)
    digits = digits.slice(digits.length - 10);
  return digits;
}


function normalizeTextForCompare_(text) {
  return (text || "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{2,}/g, "\n")
    .replace(/--\s*\n.*$/s, "") // rimuovi firma semplice dopo “--”
    .trim()
    .toLowerCase();
}


function safeSetIfColumnExists_(sheet, cols, colName, rowIndex, value) {
  if (cols && colName in cols) {
    sheet.getRange(rowIndex, cols[colName] + 1).setValue(value);
  }
}


function stripTime(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }


function tryParseDate(value) {
  if (value instanceof Date && !isNaN(value)) return value;

  if (typeof value === "string") {
    let clean = value
      .trim()
      .replace(",", "") // 🔹 rimuove virgole
      .replace(/\s+/g, " ") // 🔹 normalizza spazi
      .replace(/\./g, "/"); // 🔹 converte punti in slash (es. 9.4.2025 → 9/4/2025)

    // 🔍 dd/MM/yyyy HH:mm:ss
    let match = clean.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/
    );
    if (match) {
      const [_, dd, mm, yyyy, h, m, s] = match;
      return new Date(
        `${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}T${h}:${m}:${s}`
      );
    }

    // 🔍 dd/MM/yyyy
    match = clean.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      const [_, dd, mm, yyyy] = match;
      return new Date(`${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`);
    }

    // 🔍 yyyy-MM-dd HH:mm:ss
    match = clean.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);
    if (match) {
      const [_, yyyy, mm, dd, h, m, s] = match;
      return new Date(`${yyyy}-${mm}-${dd}T${h}:${m}:${s}`);
    }

    // fallback finale (se tutto fallisce)
    const parsed = new Date(clean);
    if (!isNaN(parsed)) return parsed;
  }

  return null;
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


function writeToLogSheet(type, message) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log") || ss.insertSheet("Log");

  logSheet.appendRow([new Date().toLocaleString(), type, message]);

  // Mantiene solo le ultime 500 righe per evitare che il log diventi enorme
  var maxRows = 500;
  var numRows = logSheet.getLastRow();
  if (numRows > maxRows) {
    logSheet.deleteRows(2, numRows - maxRows);
  }
}
