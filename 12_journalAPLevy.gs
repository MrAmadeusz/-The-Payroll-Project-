/**
 * 12_journalAPLevy.gs
 * Apprenticeship Levy Journal Processing
 * 
 * Allocates total AP levy proportionally based on Employer NI contributions
 * from hourly and salaried journal outputs
 */

/**
 * Loads and combines hourly + salaried journal outputs, filtering for Employer NI entries
 * @param {string} selectedMonth - Selected payroll month (e.g., "June 2025")
 * @returns {Array<Object>} Combined NI employer entries from both journals
 */
function getRawApLevyJournal(selectedMonth) {
  try {
    console.log(`🔍 Loading AP Levy data for ${selectedMonth}`);
    
    // Map month to financial period
    const [monthName, year] = selectedMonth.split(" ");
    const financialPeriod = getFinancialPeriodNumber(monthName);
    const searchPattern = `NIC employer P${parseInt(financialPeriod)}`; // Remove leading zero
    
    console.log(`📋 Searching for pattern: "${searchPattern}"`);
    
    // Load both journal outputs
    const hourlyData = loadJournalOutput('hourly');
    const salariedData = loadJournalOutput('salaried');
    
    console.log(`📊 Loaded ${hourlyData.length} hourly rows, ${salariedData.length} salaried rows`);
    
    // Combine and filter for Employer NI entries
    const combinedData = [...hourlyData, ...salariedData];
    const niEmployerEntries = combinedData.filter(row => {
      const memo = String(row.MEMO || "").trim();
      return memo === searchPattern;
    });
    
    console.log(`✅ Found ${niEmployerEntries.length} NI employer entries matching "${searchPattern}"`);
    
    if (niEmployerEntries.length === 0) {
      throw new Error(`No NI employer entries found for pattern "${searchPattern}". Ensure hourly and salaried journals have been processed for ${selectedMonth}.`);
    }
    
    // Validate data quality
    const invalidEntries = niEmployerEntries.filter(row => {
      const debit = parseFloat(row.DEBIT || 0);
      return isNaN(debit) || debit <= 0;
    });
    
    if (invalidEntries.length > 0) {
      console.warn(`⚠️ Found ${invalidEntries.length} entries with invalid/zero DEBIT amounts`);
    }
    
    return niEmployerEntries;
    
  } catch (error) {
    console.error(`❌ getRawApLevyJournal failed: ${error.message}`);
    throw error;
  }
}

/**
 * Loads journal output file by type (hourly or salaried)
 * @param {string} journalType - Either 'hourly' or 'salaried'
 * @returns {Array<Object>} Parsed journal data
 */
function loadJournalOutput(journalType) {
  const folderMap = {
    'hourly': CONFIG.HOURLY_JOURNAL_OUTPUT_FOLDER_ID,
    'salaried': CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID
  };
  
  const folderId = folderMap[journalType];
  if (!folderId) {
    throw new Error(`Unknown journal type: ${journalType}`);
  }
  
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.CSV);
  
  // Find the most recent journal file for this type
  let latestFile = null;
  let latestDate = new Date(0);
  const filePattern = new RegExp(`journal${journalType.charAt(0).toUpperCase() + journalType.slice(1)}`, 'i');
  
  while (files.hasNext()) {
    const file = files.next();
    if (filePattern.test(file.getName())) {
      const fileDate = file.getLastUpdated();
      if (fileDate > latestDate) {
        latestDate = fileDate;
        latestFile = file;
      }
    }
  }
  
  if (!latestFile) {
    throw new Error(`No ${journalType} journal output file found. Please run the ${journalType} journal first.`);
  }
  
  console.log(`📂 Loading ${journalType} file: ${latestFile.getName()}`);
  
  // Parse CSV file
  return parseCsvToObjects(latestFile);
}

/**
 * Processes raw NI data to create AP Levy journal entries
 * @param {Array<Object>} niData - NI employer entries
 * @param {number} totalApLevy - Total apprenticeship levy amount
 * @param {string} selectedMonth - Selected payroll month for metadata
 * @returns {Array<Object>} Complete journal entries (credit + debits)
 */
function processApLevyJournalRows(niData, totalApLevy, selectedMonth) {
  try {
    console.log(`💰 Processing AP Levy allocation: £${totalApLevy.toLocaleString()}`);
    
    // Calculate total NI employer contributions
    let totalNI = 0;
    const validEntries = [];
    
    niData.forEach(row => {
      const debitAmount = parseFloat(row.DEBIT || 0);
      if (!isNaN(debitAmount) && debitAmount > 0) {
        totalNI += debitAmount;
        validEntries.push({
          ...row,
          debitAmount: debitAmount
        });
      }
    });
    
    if (totalNI === 0) {
      throw new Error('Total NI employer contributions is zero - cannot calculate proportions');
    }
    
    console.log(`📊 Total NI contributions: £${totalNI.toLocaleString()} across ${validEntries.length} entries`);
    
    // Create journal entries
    const journalEntries = [];
    
    // Get financial metadata
    const [monthName, year] = selectedMonth.split(" ");
    const financialPeriod = getFinancialPeriodNumber(monthName);
    const journalMeta = getApLevyJournalMeta(monthName, year, parseInt(financialPeriod));
    
    // 1. Credit entry (single line for total AP Levy)
    journalEntries.push({
      DONOTIMPORT: "",
      LINE_NO: "1",
      DOCUMENT: "",
      JOURNAL: "PYRJ",
      DATE: journalMeta.dateStr,
      REVERSEDATE: "",
      DESCRIPTION: journalMeta.description,
      ACCT_NO: "9415",
      LOCATION_ID: "500",
      DEPT_ID: "",
      MEMO: journalMeta.memo,
      DEBIT: "",
      CREDIT: totalApLevy.toFixed(2),
      SOURCEENTITY: ""
    });
    
    // 2. Debit entries (proportional allocations)
    validEntries.forEach((entry, index) => {
      const proportion = entry.debitAmount / totalNI;
      const allocation = totalApLevy * proportion;
      
      journalEntries.push({
        DONOTIMPORT: "",
        LINE_NO: (index + 2).toString(),
        DOCUMENT: "",
        JOURNAL: "PYRJ",
        DATE: journalMeta.dateStr,
        REVERSEDATE: "",
        DESCRIPTION: journalMeta.description,
        ACCT_NO: "3119",
        LOCATION_ID: entry.LOCATION_ID || "",
        DEPT_ID: entry.DEPT_ID || "",
        MEMO: journalMeta.memo,
        DEBIT: allocation.toFixed(2),
        CREDIT: "",
        SOURCEENTITY: ""
      });
    });
    
    // Validation: Ensure debits balance with credit
    const totalDebits = journalEntries
      .filter(entry => entry.DEBIT)
      .reduce((sum, entry) => sum + parseFloat(entry.DEBIT), 0);
    
    const totalCredits = journalEntries
      .filter(entry => entry.CREDIT)
      .reduce((sum, entry) => sum + parseFloat(entry.CREDIT), 0);
    
    const difference = Math.abs(totalDebits - totalCredits);
    
    if (difference > 0.02) { // Allow for small rounding differences
      console.warn(`⚠️ Journal entries don't balance: Debits £${totalDebits.toFixed(2)}, Credits £${totalCredits.toFixed(2)}, Difference £${difference.toFixed(2)}`);
    } else {
      console.log(`✅ Journal balanced: ${journalEntries.length} entries totaling £${totalCredits.toFixed(2)}`);
    }
    
    return journalEntries;
    
  } catch (error) {
    console.error(`❌ processApLevyJournalRows failed: ${error.message}`);
    throw error;
  }
}

/**
 * Generates journal metadata for AP Levy entries
 * @param {string} monthName - Month name (e.g., "June")
 * @param {string} year - Year (e.g., "2025")
 * @param {number} financialPeriod - Financial period number (e.g., 3, not "03")
 * @returns {Object} Journal metadata
 */
function getApLevyJournalMeta(monthName, year, financialPeriod) {
  const date = new Date(`${monthName} 28, ${year}`);
  const periodStr = financialPeriod.toString().padStart(2, '0'); // Add leading zero back for display
  
  return {
    dateStr: Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    description: `P${periodStr} ${year} Salary and Hourly - ${monthName} ${year}`,
    memo: `AP Levy P${periodStr}`
  };
}

/**
 * Helper function removed - now getting month from selectedMonth parameter
 */

/**
 * Exports transformed AP Levy journal to balanced CSV
 */
function outputTransformedApLevyJournal(journalEntries, selectedMonth = "July 2025") {
  if (!journalEntries || journalEntries.length === 0) {
    throw new Error("No AP Levy journal entries to export");
  }

  const folder = DriveApp.getFolderById(CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const filename = `journalApLevy_${timestamp}.csv`;

  // Use the post-processor to create balanced journal
  const file = createBalancedJournalFile(journalEntries, selectedMonth, filename, folder, 'apLevy');
  
  console.log(`✅ AP Levy journal exported: ${filename} (${journalEntries.length} entries)`);
  console.log(`📁 File URL: ${file.getUrl()}`);

  return file;
}

/**
 * Main processing function called by the journal handler
 * @param {string} selectedMonth - Selected payroll month
 * @param {number} apLevyAmount - Total apprenticeship levy amount
 * @returns {File} Created journal file
 */
function processApLevyJournal(selectedMonth, apLevyAmount) {
  try {
    console.log(`🚀 Starting AP Levy journal processing for ${selectedMonth}`);
    console.log(`💰 AP Levy amount: £${apLevyAmount.toLocaleString()}`);
    
    // Step 1: Load raw NI data
    const niData = getRawApLevyJournal(selectedMonth);
    
    // Step 2: Process into journal entries
    const journalEntries = processApLevyJournalRows(niData, apLevyAmount, selectedMonth);
    
    // Step 3: Export to CSV
    const outputFile = outputTransformedApLevyJournal(journalEntries);
    
    console.log(`🎉 AP Levy journal processing complete!`);
    
    return outputFile;
    
  } catch (error) {
    console.error(`❌ processApLevyJournal failed: ${error.message}`);
    throw error;
  }
}

/**
 * Test function for development and debugging
 * @param {string} testMonth - Month to test (optional, defaults to current)
 * @param {number} testAmount - Test levy amount (optional, defaults to 1000)
 */
function testApLevyJournal(testMonth = "June 2025", testAmount = 1000) {
  try {
    console.log(`🧪 Testing AP Levy journal with ${testMonth}, £${testAmount}`);
    
    const result = processApLevyJournal(testMonth, testAmount);
    
    console.log(`✅ Test completed successfully`);
    console.log(`📄 File: ${result.getName()}`);
    console.log(`🔗 URL: ${result.getUrl()}`);
    
    return {
      success: true,
      fileName: result.getName(),
      fileUrl: result.getUrl()
    };
    
  } catch (error) {
    console.error(`❌ Test failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Diagnostic function to check available data
 * @param {string} selectedMonth - Month to check
 */
function diagnoseApLevyData(selectedMonth) {
  try {
    console.log(`🔍 Diagnosing AP Levy data for ${selectedMonth}`);
    
    const [monthName, year] = selectedMonth.split(" ");
    const financialPeriod = getFinancialPeriodNumber(monthName);
    const searchPattern = `NIC employer P${parseInt(financialPeriod)}`; // Remove leading zero
    
    console.log(`📋 Search pattern: "${searchPattern}"`);
    
    // Check if output files exist
    let hourlyData = [];
    let salariedData = [];
    
    try {
      hourlyData = loadJournalOutput('hourly');
      console.log(`✅ Hourly journal: ${hourlyData.length} rows`);
    } catch (error) {
      console.log(`❌ Hourly journal: ${error.message}`);
    }
    
    try {
      salariedData = loadJournalOutput('salaried');
      console.log(`✅ Salaried journal: ${salariedData.length} rows`);
    } catch (error) {
      console.log(`❌ Salaried journal: ${error.message}`);
    }
    
    // Check for NI employer entries
    const combinedData = [...hourlyData, ...salariedData];
    const allMemos = [...new Set(combinedData.map(row => row.MEMO).filter(memo => memo))];
    const niMemos = allMemos.filter(memo => memo.includes('NIC') || memo.includes('employer'));
    
    console.log(`📊 Total combined rows: ${combinedData.length}`);
    console.log(`📝 Unique memos containing NIC/employer: ${JSON.stringify(niMemos)}`);
    
    const matchingEntries = combinedData.filter(row => {
      const memo = String(row.MEMO || "").trim();
      return memo === searchPattern;
    });
    
    console.log(`🎯 Entries matching "${searchPattern}": ${matchingEntries.length}`);
    
    if (matchingEntries.length > 0) {
      const totalNI = matchingEntries.reduce((sum, row) => {
        const debit = parseFloat(row.DEBIT || 0);
        return sum + (isNaN(debit) ? 0 : debit);
      }, 0);
      
      console.log(`💰 Total NI amount: £${totalNI.toLocaleString()}`);
      console.log(`🏢 Unique locations: ${[...new Set(matchingEntries.map(r => r.LOCATION_ID))].join(', ')}`);
    }
    
    return {
      searchPattern: searchPattern,
      hourlyRows: hourlyData.length,
      salariedRows: salariedData.length,
      matchingEntries: matchingEntries.length,
      allNiMemos: niMemos
    };
    
  } catch (error) {
    console.error(`❌ Diagnosis failed: ${error.message}`);
    throw error;
  }
}
