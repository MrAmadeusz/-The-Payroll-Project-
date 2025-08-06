// 12_journalInvestment.gs

/**
 * Transforms the raw investment payroll journal into the finance-ready format.
 */
function transformInvestmentJournal() {
  const rawData = getRawInvestmentJournal(); // Step 1: load
  const transformed = rawData.map(processInvestmentJournalRow); // Step 2: transform
  outputTransformedInvestmentJournal(transformed); // Step 3: write to sheet
}

/**
 * Loads and parses the raw CSV from Drive - Filename must have "Investment" in the title - will also choose the most recent (edited) file if there are multiples. 
 */
function getRawInvestmentJournal() {
  const folder = DriveApp.getFolderById(CONFIG.INVESTMENT_JOURNAL_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.CSV);

  let latestMatch = null;

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().toLowerCase();

    if (name.includes("investment")) {
      if (!latestMatch || file.getLastUpdated() > latestMatch.getLastUpdated()) {
        latestMatch = file;
      }
    }
  }

  if (!latestMatch) {
    throw new Error("No CSV files containing 'Investment' found in the specified folder.");
  }

  const csvData = Utilities.parseCsv(latestMatch.getBlob().getDataAsString());
  const [headers, ...rows] = csvData;

  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}


/**
 * Applies business rules to transform a single journal row.
 */
function processInvestmentJournalRow(row) {
  if (!row["LINE_NO"]) return null; // Skip if LINE_NO is blank/null

  return {
    "DONOTIMPORT": row["DONOTIMPORT"],
    "LINE_NO": row["LINE_NO"],
    "DOCUMENT": row["DOCUMENT"],
    "JOURNAL": row["JOURNAL"],
    "DATE": row["DATE"],
    "REVERSEDATE": row["REVERSEDATE"],
    "DESCRIPTION": row["DESCRIPTION"],
    "ACCT_NO": row["ACCT_NO"],
    "LOCATION_ID": "516", // Static override
    "DEPT_ID": row["ACCT_NO"].startsWith("3") ? "915" : "",
    "MEMO": row["MEMO"],
    "DEBIT": parseAsNullableNumber(row["DEBIT"]),
    "CREDIT": parseAsNullableNumber(row["CREDIT"]),
    "SOURCEENTITY": row["SOURCEENTITY"]
  };
}

/**
 * Outputs the transformed data to a balanced CSV file
 */
function outputTransformedInvestmentJournal(data, selectedMonth = "July 2025") {
  if (!data || !data.length) {
    throw new Error("No transformed journal data to export.");
  }

  const folder = DriveApp.getFolderById(CONFIG.INVESTMENT_JOURNAL_OUTPUT_FOLDER_ID);
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const filename = `journalInvestment_${timestamp}.csv`;

  // Use the post-processor to create balanced journal
  const file = createBalancedJournalFile(data, selectedMonth, filename, folder, 'investment');
  Logger.log(`✅ File created: ${filename} (${file.getUrl()})`);

  return file;
}
