/**
 * Loader: Pulls rows from CSV file containing form responses (instead of Google Sheet)
 * UPDATED: Now looks for CSV file with "Recharges" keyword
 */
function getRawCrossChargeJournal(selectedMonth) {
  const folder = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_FOLDER_ID);
  const file = getCsvFileByKeyword(folder, "Recharges"); // Changed from getGoogleSheetFileByKeyword
  const rows = parseCsvToObjects(file); // Use existing CSV parser utility

  const payrollPeriodFormat = selectedMonth.replace(" 2025", " 25").trim().toLowerCase();

  return rows.filter(r => {
    const period = (r["Payroll Period"] || "").toString().trim().toLowerCase();
    return period === payrollPeriodFormat;
  });
}

/**
 * Implements a systematic lookup approach with prioritized matching strategies
 * 
 * @param {string} key - The key to look up
 * @param {Object} map - The mapping dictionary
 * @param {string} mapType - Either "location" or "department" for logging
 * @param {string} defaultValue - Value to return if no match is found
 * @returns {string} The mapped value or defaultValue
 */
function rigorousLookup(key, map, mapType, defaultValue = "UNKNOWN") {
  // Normalize the input key
  const normalizedKey = (key || "").toString()
    .replace(/\u00A0/g, " ")  // Replace non-breaking spaces
    .replace(/\s+/g, " ")     // Collapse multiple spaces
    .trim()
    .toLowerCase();
  
  if (!normalizedKey) {
    Logger.log(`❌ Empty ${mapType} key provided`);
    return defaultValue;
  }
  
  // STAGE 1: Direct exact match (most reliable)
  if (map[normalizedKey] !== undefined) {
    Logger.log(`✅ Direct match for ${mapType} '${key}': → '${map[normalizedKey]}'`);
    return map[normalizedKey];
  }
  
  // STAGE 2: Word boundary matching for exact phrases
  // This handles cases like "Leisure Ops" vs "Leisure Ops Marketing"
  for (const mapKey in map) {
    try {
      // Create word boundary regex for exact phrase matching
      const wordPattern = new RegExp(`^${mapKey}$`, 'i');
      if (wordPattern.test(normalizedKey)) {
        Logger.log(`✅ Exact phrase match for ${mapType} '${key}': matched '${mapKey}' → '${map[mapKey]}'`);
        return map[mapKey];
      }
    } catch (e) {
      // Skip any keys that cause regex errors
      continue;
    }
  }
  
  // Log the map contents for debugging (limited to 10 entries)
  const mapEntries = Object.entries(map).slice(0, 10);
  Logger.log(`🔍 Searching ${mapType} map with ${Object.keys(map).length} entries. First 10: ${JSON.stringify(mapEntries)}`);
  
  // STAGE 3: Check if this is a compound term that needs special handling
  // Store the successful match if found
  let matchResult = null;
  
  // Remove any common words that might interfere with matching
  const cleanKey = normalizedKey
    .replace(/\b(and|the|of|in|for|to|at|by|with)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
    
  // STAGE 4: Attempt to find the most specific match by starting with the longest keys
  // Sort map keys by length (longest first) for most specific matching
  const sortedKeys = Object.keys(map).sort((a, b) => b.length - a.length);
  
  for (const mapKey of sortedKeys) {
    // Only consider keys with 3+ characters to avoid overly general matches
    if (mapKey.length < 3) continue;
    
    // First try exact word boundary match
    try {
      const boundaryPattern = new RegExp(`\\b${mapKey}\\b`, 'i');
      if (boundaryPattern.test(normalizedKey)) {
        Logger.log(`✅ Word boundary match for ${mapType} '${key}': matched '${mapKey}' → '${map[mapKey]}'`);
        matchResult = map[mapKey];
        break; // Stop at first match
      }
    } catch (e) {
      // Skip any keys that cause regex errors
      continue;
    }
  }
  
  // If we found a match, return it
  if (matchResult) {
    return matchResult;
  }
  
  // STAGE 5: Last resort - check for significant substring matches
  // But only if the substring is at least 5 characters to avoid false positives
  for (const mapKey of sortedKeys) {
    if (mapKey.length >= 5 && normalizedKey.includes(mapKey)) {
      Logger.log(`⚠️ Substring match for ${mapType} '${key}': matched '${mapKey}' → '${map[mapKey]}'`);
      return map[mapKey];
    }
  }
  
  // No match found
  Logger.log(`❌ No mapping found for ${mapType}: '${key}'`);
  return defaultValue;
}

/**
 * Builds comprehensive location code map with data validation
 */
function buildLocationCodeMap() {
  const folder = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_FOLDER_ID);
  const file = getCsvFileByKeyword(folder, "location");
  const raw = parseCsvToObjects(file);

  // Verify the expected columns exist
  if (raw.length > 0) {
    const firstRow = raw[0];
    if (!firstRow.hasOwnProperty("Location") || !firstRow.hasOwnProperty("LocationCode")) {
      Logger.log("⚠️ Location mapping file is missing expected columns. Found: " + 
                Object.keys(firstRow).join(", "));
    }
  }

  const map = {};
  
  // First, load all exact mappings from the file
  raw.forEach(row => {
    const name = (row["Location"] || "").toString().trim();
    const code = (row["LocationCode"] || "").toString().trim();
    
    if (name && code) {
      const key = name.toLowerCase();
      
      // If we have a duplicate, log it
      if (map[key] && map[key] !== code) {
        Logger.log(`⚠️ Duplicate location mapping: '${name}' maps to both '${map[key]}' and '${code}'`);
      }
      
      map[key] = code;
    }
  });
  
  // Log total mappings loaded
  Logger.log(`📍 Loaded ${Object.keys(map).length} location mappings from file`);
  
  // Add any manual overrides for problematic mappings
  // Only add these if they're truly needed to fix specific mapping issues
  const overrides = {
    // Add specific overrides only if needed
    // Example: "location name": "location code",
  };
  
  // Apply overrides
  for (const [key, value] of Object.entries(overrides)) {
    const normalizedKey = key.toLowerCase().trim();
    if (map[normalizedKey] && map[normalizedKey] !== value) {
      Logger.log(`🔄 Overriding location mapping: '${key}' from '${map[normalizedKey]}' to '${value}'`);
    }
    map[normalizedKey] = value;
  }
  
  return map;
}

/**
 * Builds comprehensive department code map with data validation
 */
function buildDepartmentCodeMap() {
  const folder = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_FOLDER_ID);
  const file = getCsvFileByKeyword(folder, "accounting");
  const raw = parseCsvToObjects(file);
  
  // Verify the expected columns exist
  if (raw.length > 0) {
    const firstRow = raw[0];
    if (!firstRow.hasOwnProperty("Division") || !firstRow.hasOwnProperty("DivisionAccountingCode")) {
      Logger.log("⚠️ Department mapping file is missing expected columns. Found: " + 
                Object.keys(firstRow).join(", "));
    }
  }

  const map = {};
  
  // First, load all exact mappings from the file
  raw.forEach(row => {
    const div = (row["Division"] || "").toString().trim();
    const code = (row["DivisionAccountingCode"] || "").toString().trim();
    
    if (div && code) {
      const key = div.toLowerCase();
      
      // If we have a duplicate, log it
      if (map[key] && map[key] !== code) {
        Logger.log(`⚠️ Duplicate department mapping: '${div}' maps to both '${map[key]}' and '${code}'`);
      }
      
      map[key] = code;
    }
  });
  
  // Log total mappings loaded
  Logger.log(`🏬 Loaded ${Object.keys(map).length} department mappings from file`);
  
  // Add any manual overrides for problematic mappings
  // Only add these if they're truly needed to fix specific mapping issues
  const overrides = {
    // Here we can put confirmed overrides for problematic departments
    "leisure ops": "501", // Specific override for Leisure Ops
    // Add more only if needed
  };
  
  // Apply overrides
  for (const [key, value] of Object.entries(overrides)) {
    const normalizedKey = key.toLowerCase().trim();
    if (map[normalizedKey] && map[normalizedKey] !== value) {
      Logger.log(`🔄 Overriding department mapping: '${key}' from '${map[normalizedKey]}' to '${value}'`);
    }
    map[normalizedKey] = value;
  }
  
  return map;
}

/**
 * Transformer: Expands each row into two journal lines with enhanced rigor
 */
function processCrossChargeJournalRow(row, selectedMonth) {
  // Validate input row
  if (!row) {
    Logger.log("⚠️ Null row provided to processCrossChargeJournalRow");
    return null;
  }
  
  // Normalize input and calculate amount
  const hours = parseFloat((row["Hours Worked "] || "0").toString().trim());
  const rate = parseFloat((row["Rate of Pay Per Hour"] || "0").toString().trim());
  
  if (isNaN(hours) || isNaN(rate)) {
    Logger.log(`⚠️ Invalid hours (${hours}) or rate (${rate}) - original values: Hours="${row["Hours Worked "]}", Rate="${row["Rate of Pay Per Hour"]}"`);
    return null;
  }
  
  const amount = +(hours * rate).toFixed(2);
  if (!amount || amount <= 0) {
    Logger.log(`⚠️ Calculated amount is zero or negative: ${amount}`);
    return null;
  }
  
  // Get mapping dictionaries
  const locationMap = buildLocationCodeMap();
  const deptMap = buildDepartmentCodeMap();
  
  // Log dictionary sizes for debugging
  Logger.log(`📊 Location map contains ${Object.keys(locationMap).length} entries`);
  Logger.log(`📊 Department map contains ${Object.keys(deptMap).length} entries`);
  
  // Extract and normalize keys
  const fromLocKey = (row["Home Hotel "] || "").toString().trim();
  const toLocKey = (row["Hotel - Cost to be transferred to"] || "").toString().trim();
  const fromDeptKey = (row["Home Department "] || "").toString().trim();
  const toDeptKey = (row["Department Costs to be transferred to "] || "").toString().trim();
  
  // Debug log the exact keys we're looking up
  Logger.log(`🔍 LOOKUP fromLocKey='${fromLocKey}'`);
  Logger.log(`🔍 LOOKUP toLocKey='${toLocKey}'`);
  Logger.log(`🔍 LOOKUP fromDeptKey='${fromDeptKey}'`);
  Logger.log(`🔍 LOOKUP toDeptKey='${toDeptKey}'`);
  
  // Use enhanced rigorousLookup function
  const fromLoc = rigorousLookup(fromLocKey, locationMap, "from location");
  const toLoc = rigorousLookup(toLocKey, locationMap, "to location");
  const fromDept = rigorousLookup(fromDeptKey, deptMap, "from department");
  const toDept = rigorousLookup(toDeptKey, deptMap, "to department");
  
  // Log results for every lookup
  Logger.log(`🧩 RESULTS: fromLoc='${fromLoc}', toLoc='${toLoc}'`);
  Logger.log(`🧩 RESULTS: fromDept='${fromDept}', toDept='${toDept}'`);
  
  // Record specific lookup failures
  const unknownMappings = [];
  if (fromLoc === "UNKNOWN") unknownMappings.push(`From Location: "${fromLocKey}"`);
  if (toLoc === "UNKNOWN") unknownMappings.push(`To Location: "${toLocKey}"`);
  if (fromDept === "UNKNOWN") unknownMappings.push(`From Department: "${fromDeptKey}"`);
  if (toDept === "UNKNOWN") unknownMappings.push(`To Department: "${toDeptKey}"`);
  
  if (unknownMappings.length > 0) {
    Logger.log(`❗ MAPPING FAILURES: ${unknownMappings.join(", ")}`);
  }
  
  // Get financial metadata
  const finMeta = getCrossChargeMeta(selectedMonth);
  
  // Create and return the journal entries
  return [
    {
      DONOTIMPORT: "",
      LINE_NO: "",
      DOCUMENT: "",
      JOURNAL: "GJ",
      DATE: finMeta.dateStr,
      REVERSEDATE: "",
      DESCRIPTION: finMeta.description,
      ACCT_NO: "3101",
      LOCATION_ID: fromLoc,
      DEPT_ID: fromDept,
      MEMO: finMeta.memo,
      DEBIT: "",
      CREDIT: amount,
      SOURCEENTITY: ""
    },
    {
      DONOTIMPORT: "",
      LINE_NO: "",
      DOCUMENT: "",
      JOURNAL: "GJ",
      DATE: finMeta.dateStr,
      REVERSEDATE: "",
      DESCRIPTION: finMeta.description,
      ACCT_NO: "3101",
      LOCATION_ID: toLoc,
      DEPT_ID: toDept,
      MEMO: finMeta.memo,
      DEBIT: amount,
      CREDIT: "",
      SOURCEENTITY: ""
    }
  ];
}

/**
 * Creates diagnostics report for troubleshooting
 */
function createMappingDiagnostics(selectedMonth) {
  const rawData = getRawCrossChargeJournal(selectedMonth);
  const locationMap = buildLocationCodeMap();
  const deptMap = buildDepartmentCodeMap();
  
  const diagnostics = {
    totalRows: rawData.length,
    uniqueLocations: {},
    uniqueDepartments: {},
    problemMappings: []
  };
  
  // Analyze each row
  rawData.forEach((row, index) => {
    const fromLocKey = (row["Home Hotel "] || "").toString().trim();
    const toLocKey = (row["Hotel - Cost to be transferred to"] || "").toString().trim();
    const fromDeptKey = (row["Home Department "] || "").toString().trim();
    const toDeptKey = (row["Department Costs to be transferred to "] || "").toString().trim();
    
    // Track unique values
    diagnostics.uniqueLocations[fromLocKey] = true;
    diagnostics.uniqueLocations[toLocKey] = true;
    diagnostics.uniqueDepartments[fromDeptKey] = true;
    diagnostics.uniqueDepartments[toDeptKey] = true;
    
    // Check mappings
    const fromLoc = rigorousLookup(fromLocKey, locationMap, "from location");
    const toLoc = rigorousLookup(toLocKey, locationMap, "to location");
    const fromDept = rigorousLookup(fromDeptKey, deptMap, "from department");
    const toDept = rigorousLookup(toDeptKey, deptMap, "to department");
    
    // Record any problems
    const problems = [];
    if (fromLoc === "UNKNOWN") problems.push(`From Location "${fromLocKey}" not mapped`);
    if (toLoc === "UNKNOWN") problems.push(`To Location "${toLocKey}" not mapped`);
    if (fromDept === "UNKNOWN") problems.push(`From Department "${fromDeptKey}" not mapped`);
    if (toDept === "UNKNOWN") problems.push(`To Department "${toDeptKey}" not mapped`);
    
    if (problems.length > 0) {
      diagnostics.problemMappings.push({
        rowIndex: index + 1,
        problems: problems,
        fromPayroll: row["Payroll No"],
        toHotel: toLocKey,
        fromHotel: fromLocKey,
        fromDept: fromDeptKey,
        toDept: toDeptKey,
        amount: parseFloat(row["Hours Worked "] || 0) * parseFloat(row["Rate of Pay Per Hour"] || 0)
      });
    }
  });
  
  // Convert to arrays
  diagnostics.uniqueLocations = Object.keys(diagnostics.uniqueLocations).sort();
  diagnostics.uniqueDepartments = Object.keys(diagnostics.uniqueDepartments).sort();
  
  // Generate report
  let report = `# Mapping Diagnostics Report\n\n`;
  report += `Total rows processed: ${diagnostics.totalRows}\n`;
  report += `Problem mappings: ${diagnostics.problemMappings.length}\n\n`;
  
  if (diagnostics.problemMappings.length > 0) {
    report += `## Problem Mappings\n\n`;
    diagnostics.problemMappings.forEach(problem => {
      report += `Row ${problem.rowIndex}: Payroll ${problem.fromPayroll}\n`;
      report += `- From: ${problem.fromHotel} / ${problem.fromDept}\n`;
      report += `- To: ${problem.toHotel} / ${problem.toDept}\n`;
      report += `- Amount: ${problem.amount.toFixed(2)}\n`;
      report += `- Issues: ${problem.problems.join(", ")}\n\n`;
    });
  }
  
  report += `## Unique Locations (${diagnostics.uniqueLocations.length})\n\n`;
  report += diagnostics.uniqueLocations.join("\n") + "\n\n";
  
  report += `## Unique Departments (${diagnostics.uniqueDepartments.length})\n\n`;
  report += diagnostics.uniqueDepartments.join("\n") + "\n\n";
  
  // Output to file
  const folder = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_FOLDER_ID);
  const filename = `mapping_diagnostics_${selectedMonth.replace(/\s+/g, "_")}.txt`;
  folder.createFile(filename, report);
  
  return {
    report: report,
    problemCount: diagnostics.problemMappings.length
  };
}

/**
 * Metadata: Assembles values for DATE, DESCRIPTION, MEMO
 */
function getCrossChargeMeta(selectedMonth) {
  const [monthName, year] = selectedMonth.split(" ");
  const date = new Date(`${monthName} 28, ${year}`);
  const period = getFinancialPeriodNumber(monthName);
  return {
    dateStr: Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    description: `P${period} ${year} Cross Charge - ${selectedMonth}`,
    memo: `Cross Charges P${period}`
  };
}

/**
 * Maps month name → UK financial period number
 */
function getFinancialPeriodNumber(monthName) {
  const finYearStart = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"];
  const index = finYearStart.findIndex(m => m.toLowerCase() === monthName.toLowerCase());
  return index >= 0 ? String(index + 1).padStart(2, '0') : "00";
}

/**
 * Exporter: Outputs transformed Cross Charge rows to a balanced CSV file
 */
function outputTransformedCrossChargeJournal(rows, selectedMonth) {
  const flatRows = rows
    .flat()
    .filter(r => r) // Remove any null entries from failed transforms
    .map((r, i) => ({
      ...r,
      LINE_NO: (i + 1).toString()
    }));

  if (!flatRows.length) {
    throw new Error("No cross charge journal rows to export.");
  }

  const folder = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_OUTPUT_FOLDER_ID);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const filename = `crossCharge_journal_${timestamp}.csv`;

  // Use the post-processor to create balanced journal
  const file = createBalancedJournalFile(flatRows, selectedMonth, filename, folder, 'crossCharge');
  
  return file;
}


/**
 * Main function with enhanced error handling
 */
function processCrossChargeJournal(selectedMonth) {
  try {
    // First run diagnostics
    const diagnostics = createMappingDiagnostics(selectedMonth);
    Logger.log(`Diagnostics: Found ${diagnostics.problemCount} potential mapping problems`);
    
    // Step 1: Get raw data from CSV file
    const rawData = getRawCrossChargeJournal(selectedMonth);
    if (!rawData.length) {
      throw new Error(`No cross charge data found for ${selectedMonth}`);
    }
    
    Logger.log(`Found ${rawData.length} cross charge rows for ${selectedMonth}`);
    
    // Step 2: Transform each row into journal entries
    const transformedRows = rawData.map(row => 
      processCrossChargeJournalRow(row, selectedMonth)
    );
    
    // Step 3: Output to CSV
    const outputFile = outputTransformedCrossChargeJournal(transformedRows, selectedMonth);
    
    return {
      success: true,
      message: `✅ Cross charge journal created: ${outputFile.getName()}`,
      file: outputFile,
      diagnosticsReport: `Found ${diagnostics.problemCount} potential mapping issues - see diagnostics file for details`
    };
  } catch (error) {
    Logger.log(`❌ crossCharge failed: ${error.message}`);
    return {
      success: false,
      message: `❌ crossCharge failed: ${error.message}`
    };
  }
}
