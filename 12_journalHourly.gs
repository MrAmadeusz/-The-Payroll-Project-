/**
 * 1) getRawHourlyJournal()
 *    • Builds a lookup (name or code → LocationCode) from your master locations CSV
 *    • Normalizes headers in the tips file so LocationCode is always used
 *    • Returns combined hourly + tips rows, with LINE_NO & DESCRIPTION filled down
 */
function getRawHourlyJournal() {
  const folder = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_FOLDER_ID);

  // --- 1a) Build locationMap mapping name→LocationCode & code→LocationCode ---
  const locFile  = getCsvFileByKeyword(folder, "location");
  const locRaw   = parseCsvToObjects(locFile);
  let locationMap = {};

  if (locRaw.length) {
    const locHeaders   = Object.keys(locRaw[0]);
    const normalizeKey = h => h.trim().toLowerCase().replace(/\s+/g, "");
    const locHMap      = locHeaders.reduce((m,h) => {
      m[normalizeKey(h)] = h;
      return m;
    }, {});

    // detect the name column and the code column
    const locNameCol = locHMap["location"];            // e.g. "Location"
    const locCodeCol = locHMap["locationcode"];        // e.g. "LocationCode"
    if (!locNameCol || !locCodeCol) {
      throw new Error(
        "Cannot find both Location and LocationCode columns in locations CSV. "
        + "Found headers: " + JSON.stringify(locHeaders)
      );
    }

    // build map so map[name] = code, map[code] = code
    locationMap = locRaw.reduce((m,row) => {
      const name = String(row[locNameCol] || "").trim();
      const code = String(row[locCodeCol] || "").trim();
      if (name) m[name] = code;
      if (code) m[code] = code;
      return m;
    }, {});
  }

  // --- 1b) Load hourly journal rows (unchanged) ---
  const hourlyFile = getCsvFileByKeyword(folder, "hourlyJournal");
  const hourlyRows = parseCsvToObjects(hourlyFile);

  // --- 1c) Build tips rows, always pulling from the “Location” column ---
  const tipsFile = getCsvFileByKeyword(folder, "hourlyTips");
  const rawTips  = parseCsvToObjects(tipsFile);
  let tipsRows   = [];

  if (rawTips.length) {
    const tipHeaders   = Object.keys(rawTips[0]);
    const normalizeKey = h => h.trim().toLowerCase().replace(/\s+/g, "");
    const tipHMap      = tipHeaders.reduce((m,h) => {
      m[normalizeKey(h)] = h;
      return m;
    }, {});

    const payTypeCol   = tipHMap["paymenttype"];
    // Pull the raw code/name from the “Location” column in the tips CSV
    const tipCodeCol   = tipHMap["location"];
    const procDateCol  = tipHMap["processingdate"];
    const payPeriodCol = tipHMap["payperiodnumber"];
    const totalPayCol  = tipHMap["totalpayment"];

    if (!payTypeCol || !tipCodeCol) {
      throw new Error(
        "Cannot find PaymentType or Location column in tips CSV. "
        + "Found headers: " + JSON.stringify(tipHeaders)
      );
    }

    tipsRows = rawTips
      .filter(r => String(r[payTypeCol] || "").trim() === "Tips")
      .map(r => {
        const rawLoc = String(r[tipCodeCol] || "").trim();
        return {
          DONOTIMPORT:   null,
          LINE_NO:       null,
          DOCUMENT:      null,
          JOURNAL:       r[procDateCol] ? "PYRJ" : null,
          DATE:          r[procDateCol],
          REVERSEDATE:   null,
          DESCRIPTION:   null,
          ACCT_NO:       r[payPeriodCol] ? "9431" : null,
          // lookup the true LocationCode; if missing, fall back to the raw text
          LOCATION_ID:   locationMap[rawLoc] || rawLoc,
          DEPT_ID:       null,
          MEMO:          "Tips",
          DEBIT:         r[totalPayCol],
          CREDIT:        null,
          SOURCEENTITY:  null
        };
      });
  }

  // --- 1d) Combine hourly + tips rows, assign LINE_NO, fill-down DESCRIPTION ---
  const combined = hourlyRows
    .concat(tipsRows)
    .filter(r => !String(r.MEMO || "").startsWith("Tips  P"))
    .map((r, i) => ({ ...r, LINE_NO: String(i + 1) }));

  let lastDesc = null;
  combined.forEach(r => {
    if (r.DESCRIPTION) lastDesc = r.DESCRIPTION;
    else               r.DESCRIPTION = lastDesc;
  });

  return combined;
}


/**
 * 2) processHourlyJournalRow(row)
 *    • Normalizes MEMO
 *    • DEPT_ID logic: “Rounding…” → 900 first, then Advance/Classes
 *    • LOCATION_ID logic: non-tips & acct-9 or rounding → 500 (Classes→125)
 */
function processHourlyJournalRow(row) {
  const acct    = String(row.ACCT_NO    || "").trim();
  const memo    = String(row.MEMO       || "").trim();
  let   loc     = String(row.LOCATION_ID|| "").trim();
  let   DEPT_ID = row.DEPT_ID;
  const isTip   = memo === "Tips";

  // DEPT_ID logic
  if (memo.startsWith("Rounding"))              DEPT_ID = "900";
  else if (memo.startsWith("Advance"))          DEPT_ID = null;
  else if (/^P\s*T|^Pt|^Classes/.test(memo))    DEPT_ID = "501";

  // LOCATION_ID logic (skip tips)
  if (!isTip && (acct.startsWith("9") || memo.startsWith("Rounding"))) {
    loc = "500";
    if (memo.startsWith("Classe")) loc = "125";
  }

  return {
    DONOTIMPORT:  row.DONOTIMPORT,
    LINE_NO:      row.LINE_NO,
    DOCUMENT:     row.DOCUMENT,
    JOURNAL:      row.JOURNAL,
    DATE:         row.DATE,
    REVERSEDATE:  row.REVERSEDATE,
    DESCRIPTION:  row.DESCRIPTION,
    ACCT_NO:      row.ACCT_NO,
    LOCATION_ID:  loc,
    DEPT_ID:      DEPT_ID,
    MEMO:         memo,
    DEBIT:        parseAsNullableNumber(row.DEBIT),
    CREDIT:       parseAsNullableNumber(row.CREDIT),
    SOURCEENTITY: row.SOURCEENTITY
  };
}


/**
 * Outputs the transformed hourly journal to a balanced CSV file
 */
function outputTransformedHourlyJournal(data, selectedMonth = "July 2025") {
  if (!data || !data.length) {
    throw new Error("No hourly journal data to export.");
  }

  const folder = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_OUTPUT_FOLDER_ID);
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-ddHHmm");
  const filename = `journalHourly${timestamp}.csv`;

  // Use the post-processor to create balanced journal
  const file = createBalancedJournalFile(data, selectedMonth, filename, folder, 'hourly');
  Logger.log(`✅ File created: ${filename} (${file.getUrl()})`);

  return file;
}


