// GLOBALS //
const yachts = [
  { id: "5290589260000102103", name: "Phoebe" },
  { id: "8829461455002103",   name: "My Way" },
  { id: "8645971221402103",  name: "Danai - Skippered" },
  { id: "8836031455002103",  name: "Hawaii 5-0 - Skippered" },
  { id: "9046281455002103",  name: "Red Rose - Skippered" },
  { id: "5327316120000102103", name: "Breezer - Skippered" },
  { id: "6040296640000102103", name: "Penny Lane - Skippered - Green Wave" },
  { id: "6198483800000101909", name: "One Feeling"},
  { id: "2603010730000102103", name: "Blue Horizon - Skippered" },
  { id: "3164676220000102103", name: "Diva - Skippered" },
  { id: "7646950786302103",  name: "Reboot" }
];
const FIRST_SAT = new Date('2024-12-28T00:00:00Z');

const WEEK_MIN = 18;

const BOATS = {
  60456:'Penny Lane - Skippered - Green Wave', 60186:'Breezer - Skippered', 60155:'Phoebe', 59714:'Danai - Skippered',
  51165:'Hawaii 5-0 - Skippered', 60193:'Red Rose - Skippered', 60130:'My Way',38734:'Blue Horizon - Skippered',
  58245:'Reboot',39076:'Diva - Skippered'
};

const LEGEND_UNIFIED = {
    reservation:   ["X", "#FFC7CE"],      // Booked - ροζ
    option:        ["option", "#FFFF99"], // Option - κίτρινο
    free:          ["✓", "#C6EFCE"],      // Free - πράσινο
    unknown:       ["?", "#FFEB9C"],      // Unknown - απαλό κίτρινο
    service:       ["⚒", "#9fa5ab"],     // Service - γκρι
    owner:         ["🔒", "#D9D9D9"]      // Owner - γκρι
  };
  const SEDNA_MAP = {
    'FFDB7C': 'reservation',
    '00FF00': 'free',
    '86A6EF': 'option',
    '92B3FF': 'option',
    'CCCCCC': 'owner',
    '?DEFAULT?': 'unknown'
  };

// SHEET WRITER// 
function CombinedAvailability() {
  // === Helpers μέσα στη συνάρτηση για εύκολο copy-paste ===

  // όπως πριν
  function formatCurrencyGR(val) {
    if (typeof val === "number" && !isNaN(val)) {
      return val.toLocaleString("el-GR", {minimumFractionDigits:2, maximumFractionDigits:2}) + "€";
    }
    if (typeof val === "string") {
      if (val.trim() === "" || val.trim() === "?" || val === "option" || val === "X" || val === "✓") return val;
      let num = parseFloat(val.replace(/[^\d,.-]/g,"").replace(",",".")); 
      if (!isNaN(num)) {
        return num.toLocaleString("el-GR", {minimumFractionDigits:2, maximumFractionDigits:2}) + "€";
      }
    }
    return val || "";
  }

  // NEW: πόσες στήλες είναι τα summary (BOAT..SOURCE)
  const SUMMARY_COLS = 6;

  // NEW: διαβάζει τα ήδη υπάρχοντα labels εβδομάδων από το header
  function readExistingWeeks(sheet, headerRow) {
    const lastCol = sheet.getLastColumn();
    if (lastCol < SUMMARY_COLS + 1) return { weeks: [], map: new Map() };

    const labels = sheet.getRange(headerRow, SUMMARY_COLS + 1, 1, lastCol - SUMMARY_COLS)
                        .getDisplayValues()[0]
                        .map(s => (s || "").trim())
                        .filter(Boolean);
    const map = new Map();
    for (let i = 0; i < labels.length; i++) {
      map.set(labels[i], SUMMARY_COLS + 1 + i); // απόλυτη στήλη
    }
    return { weeks: labels, map };
  }

  // NEW: φροντίζει να υπάρχουν όλες οι (παλιές + νέες) εβδομάδες στο header
  function ensureHeaderWeeks(sheet, headerRow, newWeeks) {
    const { weeks: existingLabels } = readExistingWeeks(sheet, headerRow);
    const existingSet = new Set(existingLabels);
    const wantedLabels = existingLabels.slice();
    for (const w of newWeeks) {
      if (!existingSet.has(w.date)) wantedLabels.push(w.date);
    }

    // αν προστέθηκαν εβδομάδες, γράψε/επαν-γράψε το τμήμα του header για τις εβδομάδες
    if (wantedLabels.length !== existingLabels.length) {
      sheet.getRange(headerRow, SUMMARY_COLS + 1, 1, wantedLabels.length)
           .setValues([wantedLabels]);
    }

    // styling για όλη τη γραμμή header (summary + weeks)
    sheet.getRange(headerRow, 1, 1, SUMMARY_COLS + wantedLabels.length)
      .setFontWeight("bold")
      .setBackground("#003366")
      .setFontColor("white");

    // Επιστρέφουμε νέο map
    const map = new Map();
    for (let i = 0; i < wantedLabels.length; i++) {
      map.set(wantedLabels[i], SUMMARY_COLS + 1 + i);
    }
    return { labels: wantedLabels, map };
  }

  // ==== FETCH & PARSE DATA ====
  const bmBoatData   = parseBoatReservationsFromJsonFile(fetchBookingSheetJson());
  const sednaHtml    = fetchCalendarHtml();
  const sednaData    = parseBoats(sednaHtml);
  const weekMax      = getMaxWeekFromHtml(sednaHtml);
  const totals       = parseCalendarAndFetchPrices(sednaHtml);
  const globalEarliest = findEarliestAvailableWeek(bmBoatData);

  // τρέχον «παράθυρο» εβδομάδων που ήρθαν τώρα (fresh)
  const weeks = [];
  const headerRow = 4;
  for (let w = WEEK_MIN; w <= weekMax; w++) {
    const startDate = new Date(FIRST_SAT.getTime() + (w - 1) * 7 * 864e5);
    weeks.push({
      week: w,
      date: Utilities.formatDate(startDate, Session.getScriptTimeZone(), "dd MMM")
    });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Availability") || ss.insertSheet("Availability");

  // REMOVED: (δεν καθαρίζουμε πλέον όλο το grid – κρατάμε ιστορικό)
  // const firstDataRow = headerRow + 1;
  // const lastRow = sheet.getLastRow();
  // const lastCol = sheet.getMaxColumns();
  // if (lastRow >= firstDataRow) {
  //   sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).clear();
  // }

  // CHANGED: γράφουμε μόνο τα 6 summary headers (σταθερά)
  sheet.getRange(headerRow, 1, 1, SUMMARY_COLS)
    .setValues([["BOAT", "BOOKED", "OPTION", "PROJECTED", "ACTUAL", "SOURCE"]])
    .setFontWeight("bold")
    .setBackground("#003366")
    .setFontColor("white");
  sheet.setFrozenColumns(SUMMARY_COLS);

  // NEW: φτιάχνουμε/ενώνουμε header εβδομάδων (παλιές + νέες)
  const { labels: allWeekLabels, map: weekLabelToCol } = ensureHeaderWeeks(sheet, headerRow, weeks);

  // helper για projected
  function getTotalReservedSumForBoat(boatName) {
    const yacht = yachts.find(y => y.name === boatName);
    if (!yacht) return "";
    const html = fetchYachtDetailsHtml(yacht.id);
    const priceTable = parseYachtPriceTable(html);
    const bmBoat = bmBoatData.find(b => b.name === boatName);
    if (!bmBoat) return "";
    let sum = 0.0;
    for (let w = WEEK_MIN; w <= weekMax; w++) {
      const entry = bmBoat.statusByWeek[w];
      if (entry && entry.status === "reservation") {
        const weekDate = new Date(FIRST_SAT.getTime() + (w - 1) * 7 * 864e5);
        let foundPrice = null;
        for (let period of priceTable) {
          const from = parsePriceDate(period.from);
          const to = parsePriceDate(period.to);
          if (weekDate >= from && weekDate < to) { foundPrice = period.price; break; }
        }
        if (foundPrice) {
          let clean = foundPrice.replace(/[.,](?=\d{3})/g, "")
                                .replace(",", ".")
                                .replace(/[^\d.]/g, "");
          let value = parseFloat(clean);
          if (!isNaN(value)) sum += value;
        }
      }
    }
    return sum === 0 ? "" : sum;
  }

  // (τα δικά σου – τα κρατάω για το loop)
  const koinaNames = Object.values(BOATS);
  const koinaYachts = yachts.filter(y => koinaNames.includes(y.name));
  const monoMmkYachts = yachts.filter(y => !koinaNames.includes(y.name));
  const allYachtsOrdered = koinaYachts.concat(monoMmkYachts);

  // CHANGED: κάνουμε loop στα yachts (ή allYachtsOrdered – όπως προτιμάς)
  yachts.forEach((yacht, i) => {
  const name = yacht.name;
  const bmBoat = bmBoatData.find(b => b.name === name);

  let xCount = 0, optionCount = 0;
  const row1 = [name];   // θα γεμίσει με σύμβολα/νούμερα ανά εβδομάδα (MMK)
  const colors1 = [];

  for (const { week } of weeks) {
    let symbol = "", color = "";
    if (bmBoat) {
      let boatMinWeek = Math.min(...Object.keys(bmBoat.statusByWeek).map(Number));
      if (week < globalEarliest || week < boatMinWeek) {
        [symbol, color] = LEGEND_UNIFIED.unknown;
      } else {
        const entry = bmBoat.statusByWeek[week];
        if (entry) {
          const [mmkSymbol, mmkColor] = LEGEND_UNIFIED[entry.status || "unknown"] || LEGEND_UNIFIED.unknown;
          if (mmkSymbol === "X") {
            let mmkPrice = getMMKPriceForWeek(name, week, bmBoatData, WEEK_MIN, weekMax);
            symbol = (mmkPrice !== "" && mmkPrice != null) ? mmkPrice : "X"; // <-- αριθμός ή "X"
            xCount++;
          } else {
            symbol = mmkSymbol;
          }
          color = mmkColor;
          if (mmkSymbol === "option") optionCount++;
        } else {
          [symbol, color] = LEGEND_UNIFIED.free;
        }
      }
    } else {
      [symbol, color] = LEGEND_UNIFIED.unknown;
    }
    row1.push(symbol);
    colors1.push(color);
  }

  // --- SEDNA γραμμή ---
  const row2 = ["", "", "", "", "", "sedna"];   // θα γεμίσει με τιμές/σύμβολα (Sedna)
  const colors2 = ["", "", "", "", "", ""];
  const koinaNames = Object.values(BOATS);
  const isInSedna = koinaNames.includes(name);
  const sednaBoatId = Object.keys(BOATS).find(id => BOATS[id] === name);
  const sednaArr = (sednaBoatId && sednaData[sednaBoatId]) || [];
  let sednaBooked = 0, sednaOption = 0;

  for (let wi = 0; wi < weeks.length; wi++) {
    let sednaSymbol = "?", sednaColor = LEGEND_UNIFIED.unknown[1];

    if (sednaArr.length && isInSedna) {
      for (const rec of sednaArr) {
        const status = SEDNA_MAP[rec.hex] || "unknown";
        for (const w of rec.weeks) {
          if (w - WEEK_MIN === wi) {
            [sednaSymbol, sednaColor] = LEGEND_UNIFIED[status];
            if (sednaSymbol === "X") {
              let sednaPrice = getSednaBookingPriceForWeek(name, weeks[wi].week, sednaHtml);
              sednaSymbol = (sednaPrice !== "" && sednaPrice != null) ? sednaPrice : "X"; // <-- αριθμός ή "X"
              sednaBooked++;
            }
            if (sednaSymbol === "option") sednaOption++;
            break;
          }
        }
      }
    } else if (!isInSedna) {
      sednaSymbol = "";
      sednaColor = "";
    }
    row2.push(sednaSymbol);
    colors2.push(sednaColor);
  }
  row2[1] = isInSedna ? (sednaBooked || "") : "";
  row2[2] = isInSedna ? (sednaOption || "") : "";

  // === NEW: ΑΘΡΟΙΣΜΑΤΑ από τα weekly cells (χωρίς formulas) ===
  function toNumber(val) {
    if (typeof val === 'number' && isFinite(val)) return val;
    if (typeof val === 'string') {
      const s = val.replace(/[^\d,.-]/g, '').replace(',', '.');
      const n = parseFloat(s);
      return isNaN(n) ? 0 : n;
    }
    return 0;
  }
  // άθροισμα MMK (row1 από τη στήλη SUMMARY_COLS και πέρα)
  const SUMMARY_COLS = 6;
  const mmkSum = row1.slice(SUMMARY_COLS).reduce((a, v) => a + toNumber(v), 0);   // NEW
  // άθροισμα Sedna (row2 από SUMMARY_COLS και πέρα)
  const sednaSum = row2.slice(SUMMARY_COLS).reduce((a, v) => a + toNumber(v), 0); // NEW

  // === ΜΟΡΦΟΠΟΙΗΣΗ νομισμάτων ΜΕΤΑ τον υπολογισμό των αθροισμάτων ===
  let formattedRow1 = row1.slice();
  let formattedRow2 = row2.slice();

  // βάζουμε τα totals πάνω στις στήλες PROJECTED/ACTUAL (χωρίς να βασιζόμαστε σε άλλες συναρτήσεις)
  formattedRow1.splice(1, 0, xCount, optionCount, formatCurrencyGR(mmkSum), formatCurrencyGR(sednaSum), "mmk"); // NEW
  colors1.splice(0, 0, "", "", "", "", "", "");

  for (let j = SUMMARY_COLS; j < formattedRow1.length; j++) {
    formattedRow1[j] = formatCurrencyGR(formattedRow1[j]);
  }
  for (let j = SUMMARY_COLS; j < formattedRow2.length; j++) {
    formattedRow2[j] = formatCurrencyGR(formattedRow2[j]);
  }

  // === από εδώ και κάτω παραμένει όπως στο patched write-block που κρατά ιστορικό ===
  const startRow = headerRow + 1 + 2 * i;
  const totalCols = SUMMARY_COLS + allWeekLabels.length;

  const existing = sheet.getRange(startRow, 1, 2, Math.max(totalCols, sheet.getLastColumn() || totalCols)).getValues();
  let row1Vals = existing[0];
  let row2Vals = existing[1];
  if (!row1Vals || row1Vals.length < totalCols) row1Vals = Array(totalCols).fill("");
  if (!row2Vals || row2Vals.length < totalCols) row2Vals = Array(totalCols).fill("");

  // SUMMARY (γράφουμε πάντα)
  row1Vals[0] = name;
  row1Vals[1] = xCount;
  row1Vals[2] = optionCount;
  row1Vals[3] = formatCurrencyGR(mmkSum);   // NEW
  row1Vals[4] = formatCurrencyGR(sednaSum); // NEW
  row1Vals[5] = "mmk";

  row2Vals[0] = "";
  row2Vals[1] = isInSedna ? (sednaBooked || "") : "";
  row2Vals[2] = isInSedna ? (sednaOption || "") : "";
  row2Vals[3] = "";
  row2Vals[4] = "";
  row2Vals[5] = "sedna";

  // Γράφουμε ΜΟΝΟ τις «φρέσκες» εβδομάδες
  for (let wi = 0; wi < weeks.length; wi++) {
    const label = weeks[wi].date;
    const col = weekLabelToCol.get(label);
    if (!col) continue;
    const idx = col - 1;

    row1Vals[idx] = formattedRow1[SUMMARY_COLS + wi];
    row2Vals[idx] = formattedRow2[SUMMARY_COLS + wi];
  }

  sheet.getRange(startRow, 1, 1, totalCols).setValues([row1Vals]);
  sheet.getRange(startRow + 1, 1, 1, totalCols).setValues([row2Vals]);

  for (let wi = 0; wi < weeks.length; wi++) {
    const label = weeks[wi].date;
    const col = weekLabelToCol.get(label);
    if (!col) continue;
    const bg1 = colors1[SUMMARY_COLS + wi] || "";
    const bg2 = colors2[SUMMARY_COLS + wi] || "";
    sheet.getRange(startRow, col).setBackground(bg1);
    sheet.getRange(startRow + 1, col).setBackground(bg2);
  }

    // Merge (όπως πριν)
    sheet.getRange(startRow, 1, 2, 1).mergeVertically().setVerticalAlignment("middle");
    sheet.getRange(startRow, 4, 2, 1).mergeVertically().setVerticalAlignment("middle").setHorizontalAlignment("center");
    sheet.getRange(startRow, 5, 2, 1).mergeVertically().setVerticalAlignment("middle").setHorizontalAlignment("center");

    // Borders (όπως πριν)
    const lastColForBlock = sheet.getLastColumn();
    sheet.getRange(startRow, 1, 1, lastColForBlock).setBorder(
      true, false, true, false, false, false,
      "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
    sheet.getRange(startRow + 1, 1, 1, lastColForBlock).setBorder(
      false, false, true, false, false, false,
      "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  });

  // Footer border (όπως πριν)
  const lastBlockEndRow = headerRow + 1 + 2 * (yachts.length - 1) + 2;
  const lastColForFooter = sheet.getLastColumn();
  sheet.getRange(lastBlockEndRow, 1, 1, lastColForFooter).setBorder(
    true, false, false, false, false, false,
    "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // auto-resize (όπως πριν)
  for (let col = 1; col <= sheet.getLastColumn(); col++) {
    sheet.autoResizeColumn(col);
  }
}
// SEDNA HELPERS //
const BLANK_CELL = { symbol:'?', color:'#FFEB9C' };
const makeBlank  = () => Array.from({length:COLS}, _ => ({...BLANK_CELL}));

const b64     = s => Utilities.base64Encode(s, Utilities.Charset.UTF_8);
const cookies = r => {
  const raw = r.getAllHeaders()['Set-Cookie'] || r.getAllHeaders()['set-cookie'] || '';
  return Array.isArray(raw) ? raw.join('; ') : raw;
};


// SEDNA FUNCTIONS //

function fetchCalendarHtml() {

  const USER='Archontis', PASS='Vasilis<';

  /* 1. LOGIN */
  const loginResp = UrlFetchApp.fetch(
    'https://client.sednasystem.com/voucher/ajax.asp',
    { method:'post',
      contentType:'application/x-www-form-urlencoded',
      payload: 'dataaction=login&login='+encodeURIComponent(b64(USER))+
               '&password='+encodeURIComponent(b64(PASS)),
      followRedirects:false, muteHttpExceptions:true });
  let cookie = cookies(loginResp);
  Logger.log('LOGIN ➔ %s', loginResp.getResponseCode());

  /* 2. WARM */
  const warmResp = UrlFetchApp.fetch(
    'https://client.sednasystem.com/homecli/default.asp?o171',
    { headers:{Cookie:cookie}, followRedirects:true, muteHttpExceptions:true });
  cookie += '; ' + cookies(warmResp);

  /* 3. LOGO (GET) */
  const logoResp = UrlFetchApp.fetch(
    'https://client.sednasystem.com/homecli/ajx_cli.asp?st=O171&a2d=dashboard&id_logo=o171',
    { headers:{Cookie:cookie}, followRedirects:true, muteHttpExceptions:true });
  cookie += '; ' + cookies(logoResp);
  Logger.log('LOGO ➔ %s', logoResp.getResponseCode());

  /* 4. CALENDAR */
  const planResp = UrlFetchApp.fetch(
    'https://client.sednasystem.com/planning_s/default.asp',
    { headers:{Cookie:cookie}, followRedirects:true, muteHttpExceptions:true });
  Logger.log('PLAN ➔ %s  (bytes=%s)', planResp.getResponseCode(), planResp.getContentText().length);

  return planResp.getContentText();
}


function parseDate(str) {
  const [dd, mm, yyyy] = str.split("/").map(Number);
  return new Date(Date.UTC(yyyy, mm - 1, dd));
}

/**
 * Week number relative to FIRST_SAT (week 1)
 */
function getWeekNumber(date) {
  return Math.floor((date - FIRST_SAT) / (7 * 864e5)) + 1;
}

/**
 * Get minimum week seen in valid entries
 */
function getMinWeekFromHtml(html) {
  const re = /creesej_valid\([^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,"(\d{2}\/\d{2}\/\d{4})"/g;
  let minW = Infinity, m;
  while ((m = re.exec(html))) {
    const w = getWeekNumber(parseDate(m[1]));
    if (w < minW) minW = w;
  }
  return (minW === Infinity) ? WEEK_MIN : minW;
}

/**
 * Get maximum week seen in valid entries
 */
function getMaxWeekFromHtml(html) {
  const re = /creesej_valid\([^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,[^,]+,"(\d{2}\/\d{2}\/\d{4})"/g;
  let maxW = WEEK_MIN, m;
  while ((m = re.exec(html))) {
    const w = getWeekNumber(parseDate(m[1]));
    if (w > maxW) maxW = w;
  }
  return maxW;
}

function parseBoats(html) {
  const result = {};
  const dynamicMax = getMaxWeekFromHtml(html);

  // Για κάθε boat block
  const blockRe = /DT\["BN(\d+)"\]=[\s\S]*?(?=DT\["BN|<\/script>)/g;
  let match;
  while ((match = blockRe.exec(html))) {
    const boatId = match[1];
    if (!BOATS[boatId]) continue;
    const block = match[0];

    // 1) Φτιάχνουμε το inter array με τα valid+vide entries με μεταγλωττισμένες start/end σε εβδομάδα
    const lines = block.match(/creesej_(?:valid|vide)\([^)]*\);/g) || [];
    const inter = lines.map(line => {
      const args = line.match(/"[^"]*"|[^,]+/g).map(s => s.replace(/^"|"$/g,''));
      if (line.startsWith('creesej_valid')) {
        return {
          type:    'valid',
          hex:     args[1].toUpperCase(),
          startW:  getWeekNumber(parseDate(args[10])),
          endW:    getWeekNumber(parseDate(args[11]))
        };
      } else {
        return {
          type:    'vide',
          hex:     args[1].toUpperCase()
        };
      }
    });

    // 2) Κάνουμε ένα μόνο pass, γεμίζοντας entries με σωστά weeks
    const entries = [];
    let lastValidEndW = null;
    for (let i = 0; i < inter.length; i++) {
      const e = inter[i];
      if (e.type === 'valid') {
        // all weeks [startW, endW)
        const weeks = [];
        for (let w = e.startW; w < e.endW; w++) weeks.push(w);
        entries.push({ hex: e.hex, weeks });
        lastValidEndW = e.endW;
      } else if (e.type === 'vide' && lastValidEndW !== null) {
        // find next valid after this vide
        const next = inter.slice(i+1).find(x => x.type === 'valid');
        // gapEnd exclusive
        const gapEnd = next ? next.startW : (dynamicMax + 1);
        if (gapEnd > lastValidEndW) {
          const weeks = [];
          for (let w = lastValidEndW; w < gapEnd; w++) weeks.push(w);
          entries.push({ hex: e.hex, weeks });
          lastValidEndW = gapEnd;
        }
      }
    }

    if (entries.length) {
      result[boatId] = entries;
    }
  }

  Logger.log('parseBoats result: %s', JSON.stringify(result, null, 2));
  return result;
}

function parseCalendarAndFetchPrices(calendarHtml) {
  let totals = {};
  for (let bid in BOATS) totals[BOATS[bid]] = 0;

  const bookingRegex = /creesej_valid\((.*?)\);/g;
  let match;
  while (match = bookingRegex.exec(calendarHtml)) {
    let params = match[1].split(",").map(s => s.trim().replace(/^"|"$/g,""));

    let id_command = params[5];
    let boat_id = params[9];
    let date_from = params[10];
    let date_to = params[11];
    let color = params[1];

    if (color !== "FFDB7C") continue;
    if (!(boat_id in BOATS)) continue;

    let boatName = BOATS[boat_id] || "Unknown Boat";
    let contractUrl = `https://client.sednasystem.com/Operator/Special/171/contractNew.asp?typ_command=ope&id_command=${id_command}&type_doc=Fyly_contract3&id_lang_cli=0&print_lang_cli=0&print_comm=1&subope=0&nwst=O171&group=&booking_no=&serial_no=&vers=pol&flown=1`;

    try {
      let html = UrlFetchApp.fetch(contractUrl, {muteHttpExceptions: true}).getContentText();
      let priceStr = extractPriceFromHtml(html);

      Logger.log('-------------------------------');
      Logger.log('Boat: ' + boatName);
      Logger.log('Dates: ' + date_from + ' - ' + date_to);
      Logger.log('Price: ' + priceStr);
      Logger.log('URL: ' + contractUrl);

      if (priceStr && priceStr !== "Price not found") {
        let numeric = priceStr.replace(/[^\d.,]/g, '').replace(',', '.');
        let value = parseFloat(numeric);

        if (!isNaN(value)) totals[boatName] += value;
      }
    } catch (e) {
      Logger.log('Error fetching: ' + contractUrl + " for " + boatName);
    }
  }

  // Print
  Logger.log("==== ΣΥΝΟΛΟ ΑΝΑ ΣΚΑΦΟΣ ====");
  for (let boat in totals) {
    Logger.log(boat + ": " + totals[boat] + "€");
  }

  return totals; // <------ ΕΠΙΣΤΡΕΦΕΙ ΤΟ ΑΝΤΙΚΕΙΜΕΝΟ!
}


function extractPriceFromHtml(html) {
  // 1. Βρες το block (in words): <b>...euros and zero cent</b>
  let match = html.match(/\(in words\):\s*<b>(.*?)euros?/i);
  if (match) {
    var text = match[1].replace(/[\s]+/g, " ").trim();
    var amount = wordsToNumber(text);
    return amount + "€";
  }
  // 2. Fallback σε αριθμητική αναζήτηση αν δεν υπάρχει το παραπάνω
  let priceMatch = html.match(/Charter[\s\S]*?freight[\s\S]*?in[\s\S]*?total[\s\S]*?EUR[\s\u00A0]*([\d\.,]+)/i);
  if (priceMatch) return priceMatch[1] + "€";
  priceMatch = html.match(/EUR[\s\u00A0]?([\d\.,]+)/i);
  if (priceMatch) return priceMatch[1] + "€";
  priceMatch = html.match(/([\d\.,]+)\s*€/);
  if (priceMatch) return priceMatch[1] + "€";
  priceMatch = html.match(/Total\s*Amount[^0-9]*([\d\.,]+)\s?€?/i);
  if (priceMatch) return priceMatch[1] + "€";
  return "Price not found";
}


function wordsToNumber(words) {
  // Χρησιμοποιεί ένα πολύ απλό "parser" για αγγλικούς αριθμούς
  var smallNumbers = {
    zero: 0, one: 1, two: 2, three: 3, four: 4, five: 5,
    six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
    eleven: 11, twelve: 12, thirteen: 13, fourteen: 14, fifteen: 15,
    sixteen: 16, seventeen: 17, eighteen: 18, nineteen: 19
  };
  var tens = {
    twenty: 20, thirty: 30, forty: 40, fifty: 50, sixty: 60,
    seventy: 70, eighty: 80, ninety: 90
  };
  var multipliers = {
    hundred: 100, thousand: 1000, million: 1000000
  };

  var tokens = words.replace(/ and /g, ' ').replace(/-/g, ' ').split(/\s+/);
  var n = 0, group = 0;
  for (var i=0; i<tokens.length; ++i) {
    var w = tokens[i].toLowerCase();
    if (smallNumbers[w] !== undefined) group += smallNumbers[w];
    else if (tens[w] !== undefined) group += tens[w];
    else if (w === "hundred") group *= 100;
    else if (w === "thousand") { n += group * 1000; group = 0; }
    else if (w === "million") { n += group * 1000000; group = 0; }
  }
  return n + group;
}

// MMK FUNCTIONS // 

function fetchBookingSheetJson() {
  // ΒΗΜΑ 1: LOGIN
  var loginUrl = "https://portal.booking-manager.com/wbm2/app/login_register/";
  var payload = {
    "is_post_back": "1",
    "refPage": "",
    "login_email": "christos.georgopoulos@gmail.com",
    "login_password": "kyngIq-7qanqy-dumpab"
  };
  var options = {
    "method": "post",
    "payload": payload,
    "followRedirects": false,
    "muteHttpExceptions": true
  };

  var loginResp = UrlFetchApp.fetch(loginUrl, options);
  var cookies = "";
  var headers = loginResp.getAllHeaders();
  if (headers["Set-Cookie"]) {
    if (Array.isArray(headers["Set-Cookie"])) {
      cookies = headers["Set-Cookie"].map(function(cookie) {
        return cookie.split(";")[0];
      }).join("; ");
    } else {
      cookies = headers["Set-Cookie"].split(";")[0];
    }
  } else {
    Logger.log("No cookie found in login response.");
    return;
  }

  Logger.log("Cookies: " + cookies);

  // ΒΗΜΑ 2: Fetch Booking Sheet JSON
  // --- ΠΡΟΣΑΡΜΟΣΕ το παρακάτω URL στις ημερομηνίες/filters που θες! ---
  var bookingSheetUrl = "https://portal.booking-manager.com/wbm2/page.html?responseType=JSON&view=BookingSheetData&companyid=7690&from=1754082000000&to=1786136399059&timeZoneOffsetInMins=-180&fromFormatted=2025-08-02%2000:00&toFormatted=2026-08-07%2023:59&daily=false&filter_discounts=false&isOnHubSpot=false&resultsPage=1&filter_country=GR&filter_region=35&filter_region=10&filter_region=7&filter_service=2103&filter_service=1909&filter_base=13&filter_base=4945797760000100000&filter_base=216&filter_base=1935994390000100000&filterlocationdistance=5000&filter_year=2025&filter_month=7&filter_date=3&filter_duration=7&filter_flexibility=on_day&filter_service_type=all&filter_model=3947847730000100000&filter_model=1399966290000100000&filter_model=800608360000100000&filter_model=4030569900000100000&filter_model=1305064610000100000&filter_model=780746060000100000&filter_length_ft=0-2000&filter_cabins=0-2000&filter_berths=0-2000&filter_heads=0-2000&filter_price=0-10001000&filter_yachtage=0-7&filter_year_from=2018&filter_availability_status=-1";

  var sheetOptions = {
    "method": "get",
    "headers": {
      "Cookie": cookies,
      "Accept": "*/*", // όπως φαίνεται στα headers σου
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
    },
    "muteHttpExceptions": true
  };

  var sheetResp = UrlFetchApp.fetch(bookingSheetUrl, sheetOptions);
  var content = sheetResp.getContentText(); // string
  var bytes = sheetResp.getContent();       // blob

  Logger.log("Μέγεθος χαρακτήρων JSON: " + content.length);
  Logger.log("Μέγεθος σε bytes: " + bytes.length);
  Logger.log("Πρώτα 200 χαρακτήρες:\n" + content.substr(0, 200));
  return sheetResp.getContentText();
  
}

/** fetch yacht details (html) **/
function fetchYachtDetailsHtml(yachtId) {
  // 1. Login και πάρε το cookie (ή χρησιμοποίησε το ίδιο session όπως πριν)
  var loginUrl = "https://portal.booking-manager.com/wbm2/app/login_register/";
  var payload = {
    "is_post_back": "1",
    "refPage": "",
    "login_email": "christos.georgopoulos@gmail.com",
    "login_password": "kyngIq-7qanqy-dumpab"
  };
  var options = {
    "method": "post",
    "payload": payload,
    "followRedirects": false,
    "muteHttpExceptions": true
  };

  var loginResp = UrlFetchApp.fetch(loginUrl, options);
  var cookies = "";
  var headers = loginResp.getAllHeaders();
  if (headers["Set-Cookie"]) {
    if (Array.isArray(headers["Set-Cookie"])) {
      cookies = headers["Set-Cookie"].map(function(cookie) {
        return cookie.split(";")[0];
      }).join("; ");
    } else {
      cookies = headers["Set-Cookie"].split(";")[0];
    }
  } else {
    Logger.log("No cookie found in login response.");
    return;
  }

  // 2. Φτιάξε το url με το yacht id που παίρνει ως όρισμα
  var yachtUrl = "https://portal.booking-manager.com/wbm2/page.html?view=YachtDetails&templateType=responsive&companyid=7690&yachtId=" 
    + yachtId + "&addMargins=true&setlang=en&setCurrency=EUR";

  // 3. Fetch το HTML
  var htmlOptions = {
    "method": "get",
    "headers": {
      "Cookie": cookies,
      "Accept": "*/*",
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
    },
    "muteHttpExceptions": true
  };

  var htmlResp = UrlFetchApp.fetch(yachtUrl, htmlOptions);
  var html = htmlResp.getContentText();
  return html;
}


function parseBoatReservationsFromJsonFile(jsonText) {
  const json = JSON.parse(jsonText);
  const MS_IN_WEEK = 7 * 24 * 60 * 60 * 1000;

  // Φτιάξε ένα Set με όλα τα έγκυρα ονόματα από τον πίνακα yachts
  const validNames = new Set(yachts.map(y => y.name));

  const boats = [];

  json.boats.forEach(boat => {
    // **ΚΡΑΤΑΜΕ μόνο αν το όνομα υπάρχει στα yachts**
    if (!validNames.has(boat.name)) return;

    const statusByWeek = {};
    (boat.reservations || []).forEach(res => {
      const startDate = new Date(res.dateFrom);
      const endDate = new Date(res.dateTo);

      // 💥 Αφαιρούμε 1 ημέρα για να μην περιληφθεί η εβδομάδα του dateTo
      endDate.setDate(endDate.getDate() - 1);

      const status = res.status;
      const startWeek = Math.floor((startDate - FIRST_SAT) / MS_IN_WEEK) + 1;
      const endWeek = Math.floor((endDate - FIRST_SAT) / MS_IN_WEEK) + 1;

      for (let w = startWeek; w <= endWeek; w++) {
        statusByWeek[w] = {
          status: status,
          hex: res.color || "",
          start: res.dateFrom,
          end: res.dateTo
        };
      }
    });
    boats.push({
      name: boat.name,
      statusByWeek: statusByWeek
    });
  });

  return boats;
}
/** parse yacht details (html) **/
function parseYachtPriceTable(html) {
  const results = [];
  const rowRe = /<tr>\s*<td class="row date">([^<]+)<\/td>\s*<td class="row date">([^<]+)<\/td>\s*<td class="row price">([^<]+)<\/td>/g;
  let m;
  while ((m = rowRe.exec(html)) !== null) {
    results.push({
      from: m[1].trim(),
      to: m[2].trim(),
      price: m[3].trim()
    });
  }
  return results;
}

function formatDate(date) {
  return date.toISOString().split('T')[0];
}

function findEarliestAvailableWeek(boatData) {
  let minWeek = Infinity;
  boatData.forEach(boat => {
    Object.keys(boat.statusByWeek).forEach(weekStr => {
      const week = parseInt(weekStr, 10);
      if (week < minWeek) minWeek = week;
    });
  });
  return minWeek === Infinity ? null : minWeek;
}

function parsePriceDate(d) {
    if (!d) return new Date("2100-01-01"); // dummy μελλοντική
    const [day, month, year] = d.split(".");
    return new Date(`${year}-${month}-${day}T00:00:00Z`);
  }

/**
 * Παίρνει boat name και εβδομάδα και επιστρέφει την τιμή καταλόγου για αυτήν την εβδομάδα (MMK).
 * Επιστρέφει "" αν δεν υπάρχει τιμή για αυτή την εβδομάδα.
 */
function getMMKPriceForWeek(boatName, weekNumber, boatData, weekMin, weekMax) {
  const boat = boatData.find(b => b.name === boatName);
  if (!boat) return null;

  const yacht = yachts.find(y => y.name === boatName);
  if (!yacht) return null;

  const html = fetchYachtDetailsHtml(yacht.id);
  const priceTable = parseYachtPriceTable(html);

  const entry = boat.statusByWeek[weekNumber];
  if (!entry || entry.status !== "reservation") return null;

  // Βρίσκουμε τη σωστή περίοδο τιμής για αυτή την εβδομάδα
  const weekDate = new Date(FIRST_SAT.getTime() + (weekNumber - 1) * 7 * 864e5);
  let foundPrice = null;
  for (let period of priceTable) {
    const from = parsePriceDate(period.from);
    const to = parsePriceDate(period.to);
    if (weekDate >= from && weekDate < to) {
      foundPrice = period.price;
      break;
    }
  }
  if (foundPrice) {
    let clean = foundPrice.replace(/[.,](?=\d{3})/g, "")
                          .replace(",", ".")
                          .replace(/[^\d.]/g, "");
    let value = parseFloat(clean);
    if (!isNaN(value)) return value;  // Επιστρέφει αριθμό!
  }
  return null;
}

/**
 * Παίρνει boat name και εβδομάδα και επιστρέφει την τιμή κράτησης (Sedna).
 * Επιστρέφει "" αν δεν βρει τιμή για αυτήν την εβδομάδα/σκάφος.
 */
function getSednaBookingPriceForWeek(boatName, weekNumber, calendarHtml) {
  const boatId = Object.keys(BOATS).find(id => BOATS[id] === boatName);
  if (!boatId) return null;

  const re = new RegExp(`creesej_valid\\(([^\\)]*)\\);`, 'g');
  let match;
  while ((match = re.exec(calendarHtml))) {
    const params = match[1].split(",").map(s => s.trim().replace(/^"|"$/g,""));
    let color = params[1];
    let id_command = params[5];
    let id_boat = params[9];
    let date_from = params[10];
    let date_to = params[11];
    if (id_boat !== boatId) continue;
    if (color !== "FFDB7C") continue; // Booked μόνο

    // Βρίσκουμε τις εβδομάδες αυτής της κράτησης
    const startDate = parseDate(date_from);
    const endDate = parseDate(date_to);
    endDate.setDate(endDate.getDate() - 1);
    const startWeek = getWeekNumber(startDate);
    const endWeek = getWeekNumber(endDate);

    if (weekNumber >= startWeek && weekNumber <= endWeek) {
      let contractUrl = `https://client.sednasystem.com/Operator/Special/171/contractNew.asp?typ_command=ope&id_command=${id_command}&type_doc=Fyly_contract3&id_lang_cli=0&print_lang_cli=0&print_comm=1&subope=0&nwst=O171&group=&booking_no=&serial_no=&vers=pol&flown=1`;
      try {
        let html = UrlFetchApp.fetch(contractUrl, {muteHttpExceptions: true}).getContentText();
        let priceStr = extractPriceFromHtml(html);
        if (priceStr && priceStr !== "Price not found") {
          let numeric = priceStr.replace(/[^\d.,]/g, '').replace(',', '.');
          let value = parseFloat(numeric);
          if (!isNaN(value)) {
            let numWeeks = endWeek - startWeek + 1;
            let perWeek = value / numWeeks;
            return perWeek; // Επιστρέφει αριθμό!
          }
        }
      } catch (e) {
        return null;
      }
    }
  }
  return null;
}

function ReloadCombinedAvailability() {
  var today = new Date();
 var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Availability");
  // Π.χ. κελί F2 (ή όπου θέλεις)
  var loadingCell = sheet.getRange("D2");
  var outpoutCell = sheet.getRange("A27");
  loadingCell.setValue("Loading");
  loadingCell.setBackground("#00FF00"); // Πράσινο

  SpreadsheetApp.flush(); // ΠΕΝΤΑΠΟΛΥ ΣΗΜΑΝΤΙΚΟ! Εμφανίζει αμέσως το loading πριν συνεχίσει ο κώδικας

  // -- Τρέχει το main σου function --
  CombinedAvailability();

  // -- Σβήνει το Loading όταν τελειώσει --
  loadingCell.setValue("");
  loadingCell.setBackground(null);
  outpoutCell.setValue("Last update:" + formattedDate);
  outpoutCell.setHorizontalAlignment("right");
}

function scheduleAndRunCombinedAvailability() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Availability");

  const loadingCell = sheet.getRange("E2");

  // Μικρό status
  loadingCell.setValue("Scheduled");
  loadingCell.setBackground("#00FF00");
  SpreadsheetApp.flush();

  // 1) Καθάρισε παλιά triggers της CombinedAvailability
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "CombinedAvailabilityTriggerWrapper") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 2) Δημιούργησε νέο daily trigger για 08:00 (ώρα project -> Athens)
  ScriptApp.newTrigger("CombinedAvailabilityTriggerWrapper")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

}

function CombinedAvailabilityTriggerWrapper() {
  // Τρέχει την κανονική CombinedAvailability
  CombinedAvailability();

  // Μετά γράφει την ώρα εκτέλεσης
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Availability");
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const outputCell = sheet.getRange("A27"); // διόρθωσα το typo "outpoutCell"

  outputCell.setValue("Last update: " + formattedDate);
  outputCell.setHorizontalAlignment("right");

}

function unscheduleAndRunCombinedAvailability() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Availability");

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "CombinedAvailabilityTriggerWrapper") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  var loadingCell = sheet.getRange("E2");
  loadingCell.setValue("");
  loadingCell.setBackground(null); // Πράσινο

}



