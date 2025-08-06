/**
 * 12_journalPostProcessor.gs
 * Universal Journal Post-Processing System
 * 
 * Automatically applies rounding adjustments to ensure perfect debit/credit balance
 * for all journal types without manual intervention.
 * 
 * Integration: Call processJournalForImport() before saving any journal to CSV
 */

/**
 * Main post-processing function - ensures journal balances perfectly
 * @param {Array<Object>} journalEntries - Array of journal entry objects
 * @param {string} selectedMonth - Month selection (e.g., "July 2025")
 * @param {string} journalType - Type for logging (optional)
 * @returns {Array<Object>} Balanced journal entries ready for import
 */
function processJournalForImport(journalEntries, selectedMonth, journalType = 'unknown') {
  try {
    console.log(`🔧 Post-processing ${journalType} journal for import compliance`);
    
    if (!journalEntries || !Array.isArray(journalEntries) || journalEntries.length === 0) {
      console.log(`⚠️ Empty journal provided for ${journalType}`);
      return journalEntries;
    }
    
    // Step 1: Calculate current balance
    const balance = calculateJournalBalance(journalEntries);
    console.log(`📊 Current balance: Debits £${balance.totalDebits.toLocaleString()}, Credits £${balance.totalCredits.toLocaleString()}`);
    
    // Step 2: Determine if adjustment needed
    const imbalance = balance.totalDebits - balance.totalCredits;
    const needsAdjustment = Math.abs(imbalance) > 0.005; // 0.5 pence tolerance
    
    if (!needsAdjustment) {
      console.log(`✅ Journal is balanced - removing any existing rounding lines`);
      return removeExistingRoundingLines(journalEntries, selectedMonth);
    }
    
    console.log(`⚖️ Journal imbalance: £${imbalance.toFixed(3)} (${imbalance > 0 ? 'Credits short' : 'Debits short'})`);
    
    // Step 3: Apply rounding adjustment
    const balancedJournal = applyRoundingAdjustment(journalEntries, selectedMonth, imbalance);
    
    // Step 4: Verify final balance
    const finalBalance = calculateJournalBalance(balancedJournal);
    const finalImbalance = Math.abs(finalBalance.totalDebits - finalBalance.totalCredits);
    
    if (finalImbalance > 0.005) {
      throw new Error(`Failed to balance journal: final imbalance £${finalImbalance.toFixed(3)}`);
    }
    
    console.log(`✅ Journal balanced successfully with rounding adjustment: £${Math.abs(imbalance).toFixed(2)}`);
    return balancedJournal;
    
  } catch (error) {
    console.error(`❌ Journal post-processing failed for ${journalType}: ${error.message}`);
    throw new Error(`Post-processing failed: ${error.message}`);
  }
}

/**
 * Calculates total debits and credits in a journal
 * @param {Array<Object>} journalEntries - Journal entries
 * @returns {Object} Balance totals
 */
function calculateJournalBalance(journalEntries) {
  let totalDebits = 0;
  let totalCredits = 0;
  
  journalEntries.forEach(entry => {
    const debit = parseFloat(entry.DEBIT || 0);
    const credit = parseFloat(entry.CREDIT || 0);
    
    if (!isNaN(debit)) totalDebits += debit;
    if (!isNaN(credit)) totalCredits += credit;
  });
  
  return {
    totalDebits: Math.round(totalDebits * 100) / 100,
    totalCredits: Math.round(totalCredits * 100) / 100,
    entryCount: journalEntries.length
  };
}

/**
 * Extracts period number from selected month
 * @param {string} selectedMonth - e.g., "July 2025"
 * @returns {string} Period number (e.g., "04")
 */
function extractPeriodFromMonth(selectedMonth) {
  if (!selectedMonth || typeof selectedMonth !== 'string') {
    console.warn(`Invalid selectedMonth: ${selectedMonth}`);
    return '01'; // Default fallback
  }
  
  const monthName = selectedMonth.split(' ')[0].trim();
  
  const periodMap = {
    'April': '01',
    'May': '02',
    'June': '03',
    'July': '04',
    'August': '05',
    'September': '06',
    'October': '07',
    'November': '08',
    'December': '09',
    'January': '10',
    'February': '11',
    'March': '12'
  };
  
  const period = periodMap[monthName] || '01';
  console.log(`📅 Extracted period P${period} from ${selectedMonth}`);
  return period;
}

/**
 * Removes existing rounding lines from journal
 * @param {Array<Object>} journalEntries - Journal entries
 * @param {string} selectedMonth - Selected month
 * @returns {Array<Object>} Journal without rounding lines
 */
function removeExistingRoundingLines(journalEntries, selectedMonth) {
  const period = extractPeriodFromMonth(selectedMonth);
  const roundingMemo = `Rounding P${period}`;
  
  const filteredEntries = journalEntries.filter(entry => {
    const memo = String(entry.MEMO || '').trim();
    const isRounding = memo === roundingMemo;
    if (isRounding) {
      console.log(`🗑️ Removing existing rounding line: ${memo}`);
    }
    return !isRounding;
  });
  
  // Renumber line numbers if any were removed
  return renumberJournalLines(filteredEntries);
}

/**
 * Applies rounding adjustment to balance the journal
 * @param {Array<Object>} journalEntries - Journal entries
 * @param {string} selectedMonth - Selected month
 * @param {number} imbalance - Amount to adjust (positive = need more credits)
 * @returns {Array<Object>} Balanced journal entries
 */
function applyRoundingAdjustment(journalEntries, selectedMonth, imbalance) {
  const period = extractPeriodFromMonth(selectedMonth);
  const roundingMemo = `Rounding P${period}`;
  
  // Remove any existing rounding lines first
  let workingEntries = removeExistingRoundingLines(journalEntries, selectedMonth);
  
  // Calculate required credit adjustment (always positive for CREDIT field)
  const creditAdjustment = Math.abs(imbalance);
  
  // Find a sample entry to copy date and description from
  const sampleEntry = workingEntries.find(entry => entry.DATE && entry.DESCRIPTION) || workingEntries[0];
  
  if (!sampleEntry) {
    throw new Error('No valid sample entry found for rounding line template');
  }
  
  // Create rounding adjustment line
  const roundingLine = {
    DONOTIMPORT: "",
    LINE_NO: "", // Will be set during renumbering
    DOCUMENT: "",
    JOURNAL: sampleEntry.JOURNAL || "PYRJ",
    DATE: sampleEntry.DATE || "",
    REVERSEDATE: "",
    DESCRIPTION: sampleEntry.DESCRIPTION || "",
    ACCT_NO: "3101",
    LOCATION_ID: "500",
    DEPT_ID: "900",
    MEMO: roundingMemo,
    DEBIT: "",
    CREDIT: creditAdjustment.toFixed(2),
    SOURCEENTITY: ""
  };
  
  console.log(`💰 Adding rounding line: ${roundingMemo} = £${creditAdjustment.toFixed(2)} CREDIT`);
  
  // Add rounding line to end
  workingEntries.push(roundingLine);
  
  // Renumber all lines
  return renumberJournalLines(workingEntries);
}

/**
 * Renumbers LINE_NO field sequentially
 * @param {Array<Object>} journalEntries - Journal entries
 * @returns {Array<Object>} Renumbered entries
 */
function renumberJournalLines(journalEntries) {
  return journalEntries.map((entry, index) => ({
    ...entry,
    LINE_NO: (index + 1).toString()
  }));
}

/**
 * Validates journal structure before processing
 * @param {Array<Object>} journalEntries - Journal entries to validate
 * @returns {Object} Validation results
 */
function validateJournalStructure(journalEntries) {
  const validation = {
    isValid: true,
    warnings: [],
    errors: []
  };
  
  if (!journalEntries || !Array.isArray(journalEntries)) {
    validation.errors.push('Journal entries must be an array');
    validation.isValid = false;
    return validation;
  }
  
  if (journalEntries.length === 0) {
    validation.warnings.push('Journal is empty');
    return validation;
  }
  
  // Check for required fields
  const requiredFields = ['ACCT_NO', 'LOCATION_ID', 'MEMO'];
  const sampleEntry = journalEntries[0];
  
  requiredFields.forEach(field => {
    if (!sampleEntry.hasOwnProperty(field)) {
      validation.errors.push(`Missing required field: ${field}`);
      validation.isValid = false;
    }
  });
  
  // Check for entries with both debit and credit
  journalEntries.forEach((entry, index) => {
    const debit = parseFloat(entry.DEBIT || 0);
    const credit = parseFloat(entry.CREDIT || 0);
    
    if (debit > 0 && credit > 0) {
      validation.warnings.push(`Line ${index + 1} has both debit and credit amounts`);
    }
    
    if (debit === 0 && credit === 0) {
      validation.warnings.push(`Line ${index + 1} has no debit or credit amount`);
    }
  });
  
  return validation;
}

/**
 * Integration wrapper for existing journal output functions
 * Call this instead of directly saving CSV
 * @param {Array<Object>} journalEntries - Raw journal entries
 * @param {string} selectedMonth - Selected month
 * @param {string} filename - Output filename
 * @param {Folder} folder - Output folder
 * @param {string} journalType - Journal type for logging
 * @returns {File} Created and balanced CSV file
 */
function createBalancedJournalFile(journalEntries, selectedMonth, filename, folder, journalType = 'unknown') {
  try {
    console.log(`📄 Creating balanced journal file: ${filename}`);
    
    // Step 1: Validate structure
    const validation = validateJournalStructure(journalEntries);
    if (!validation.isValid) {
      throw new Error(`Journal validation failed: ${validation.errors.join(', ')}`);
    }
    
    if (validation.warnings.length > 0) {
      console.warn(`⚠️ Journal warnings: ${validation.warnings.join(', ')}`);
    }
    
    // Step 2: Apply post-processing
    const balancedEntries = processJournalForImport(journalEntries, selectedMonth, journalType);
    
    // Step 3: Create CSV content
    const csvContent = createJournalCsvContent(balancedEntries);
    
    // Step 4: Save file
    const file = folder.createFile(filename, csvContent, MimeType.CSV);
    
    console.log(`✅ Balanced journal created: ${filename} (${balancedEntries.length} entries)`);
    console.log(`📁 File URL: ${file.getUrl()}`);
    
    return file;
    
  } catch (error) {
    console.error(`❌ Failed to create balanced journal: ${error.message}`);
    throw error;
  }
}

/**
 * Creates properly formatted CSV content from journal entries
 * @param {Array<Object>} journalEntries - Journal entries
 * @returns {string} CSV content
 */
function createJournalCsvContent(journalEntries) {
  if (!journalEntries || journalEntries.length === 0) {
    throw new Error('No journal entries to convert to CSV');
  }
  
  // Standard journal headers
  const headers = [
    "DONOTIMPORT", "LINE_NO", "DOCUMENT", "JOURNAL", "DATE", "REVERSEDATE",
    "DESCRIPTION", "ACCT_NO", "LOCATION_ID", "DEPT_ID", "MEMO",
    "DEBIT", "CREDIT", "SOURCEENTITY"
  ];
  
  // Build CSV rows
  const csvRows = [headers];
  
  journalEntries.forEach(entry => {
    const row = headers.map(header => {
      const value = entry[header] || "";
      // Proper CSV escaping
      if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
        return `"${value.replace(/"/g, '""')}"`;
      }
      return value;
    });
    csvRows.push(row);
  });
  
  return csvRows.map(row => row.join(",")).join("\r\n");
}

/**
 * Test function to verify post-processing works correctly
 * @param {string} testMonth - Month to test with
 * @returns {Object} Test results
 */
function testJournalPostProcessor(testMonth = "July 2025") {
  try {
    console.log(`🧪 Testing journal post-processor with ${testMonth}`);
    
    // Create test journal with intentional imbalance
    const testJournal = [
      {
        DONOTIMPORT: "", LINE_NO: "1", DOCUMENT: "", JOURNAL: "PYRJ",
        DATE: "28/07/2025", REVERSEDATE: "", DESCRIPTION: "Test Journal",
        ACCT_NO: "3101", LOCATION_ID: "500", DEPT_ID: "501", MEMO: "Test Entry 1",
        DEBIT: "100.00", CREDIT: "", SOURCEENTITY: ""
      },
      {
        DONOTIMPORT: "", LINE_NO: "2", DOCUMENT: "", JOURNAL: "PYRJ",
        DATE: "28/07/2025", REVERSEDATE: "", DESCRIPTION: "Test Journal",
        ACCT_NO: "3102", LOCATION_ID: "501", DEPT_ID: "502", MEMO: "Test Entry 2",
        DEBIT: "50.33", CREDIT: "", SOURCEENTITY: ""
      },
      {
        DONOTIMPORT: "", LINE_NO: "3", DOCUMENT: "", JOURNAL: "PYRJ",
        DATE: "28/07/2025", REVERSEDATE: "", DESCRIPTION: "Test Journal",
        ACCT_NO: "9431", LOCATION_ID: "500", DEPT_ID: "", MEMO: "Test Credit",
        DEBIT: "", CREDIT: "150.00", SOURCEENTITY: ""
      }
    ];
    
    console.log(`📊 Test journal: Debits £150.33, Credits £150.00, Imbalance £0.33`);
    
    // Process journal
    const balanced = processJournalForImport(testJournal, testMonth, 'test');
    
    // Verify balance
    const finalBalance = calculateJournalBalance(balanced);
    const finalImbalance = Math.abs(finalBalance.totalDebits - finalBalance.totalCredits);
    
    // Check for rounding line
    const roundingLine = balanced.find(entry => 
      String(entry.MEMO || '').trim() === 'Rounding P04'
    );
    
    const results = {
      success: finalImbalance <= 0.005,
      originalEntries: testJournal.length,
      finalEntries: balanced.length,
      finalBalance: finalBalance,
      finalImbalance: finalImbalance,
      roundingLineAdded: !!roundingLine,
      roundingAmount: roundingLine ? parseFloat(roundingLine.CREDIT || 0) : 0
    };
    
    if (results.success) {
      console.log(`✅ Test passed: Journal balanced with ${results.finalEntries} entries`);
      if (roundingLine) {
        console.log(`💰 Rounding adjustment: £${results.roundingAmount.toFixed(2)} CREDIT`);
      }
    } else {
      console.log(`❌ Test failed: Final imbalance £${finalImbalance.toFixed(3)}`);
    }
    
    return results;
    
  } catch (error) {
    console.error(`❌ Post-processor test failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Batch test function for multiple scenarios
 * @returns {Object} Comprehensive test results
 */
function testJournalPostProcessorComprehensive() {
  console.log(`🧪 Running comprehensive post-processor tests`);
  
  const testScenarios = [
    { name: 'Balanced Journal', imbalance: 0 },
    { name: 'Small Imbalance', imbalance: 0.03 },
    { name: 'Large Imbalance', imbalance: 5.47 },
    { name: 'Negative Imbalance', imbalance: -2.15 }
  ];
  
  const results = {
    totalTests: testScenarios.length,
    passed: 0,
    failed: 0,
    details: []
  };
  
  testScenarios.forEach(scenario => {
    try {
      console.log(`\n--- Testing: ${scenario.name} ---`);
      
      // Create test journal with specific imbalance
      const testJournal = createTestJournal(scenario.imbalance);
      const processed = processJournalForImport(testJournal, "July 2025", scenario.name);
      const balance = calculateJournalBalance(processed);
      const finalImbalance = Math.abs(balance.totalDebits - balance.totalCredits);
      
      const testResult = {
        scenario: scenario.name,
        success: finalImbalance <= 0.005,
        originalImbalance: scenario.imbalance,
        finalImbalance: finalImbalance,
        entriesAdded: processed.length - testJournal.length
      };
      
      if (testResult.success) {
        results.passed++;
        console.log(`✅ ${scenario.name}: PASSED`);
      } else {
        results.failed++;
        console.log(`❌ ${scenario.name}: FAILED - Final imbalance: £${finalImbalance.toFixed(3)}`);
      }
      
      results.details.push(testResult);
      
    } catch (error) {
      results.failed++;
      results.details.push({
        scenario: scenario.name,
        success: false,
        error: error.message
      });
      console.log(`❌ ${scenario.name}: ERROR - ${error.message}`);
    }
  });
  
  console.log(`\n📊 COMPREHENSIVE TEST RESULTS:`);
  console.log(`Total Tests: ${results.totalTests}`);
  console.log(`Passed: ${results.passed}`);
  console.log(`Failed: ${results.failed}`);
  console.log(`Success Rate: ${((results.passed / results.totalTests) * 100).toFixed(1)}%`);
  
  return results;
}

/**
 * Helper function to create test journal with specific imbalance
 * @param {number} targetImbalance - Desired imbalance amount
 * @returns {Array<Object>} Test journal entries
 */
function createTestJournal(targetImbalance) {
  const baseDebit = 100;
  const baseCredit = 100 + targetImbalance; // This creates the imbalance
  
  return [
    {
      DONOTIMPORT: "", LINE_NO: "1", DOCUMENT: "", JOURNAL: "PYRJ",
      DATE: "28/07/2025", REVERSEDATE: "", DESCRIPTION: "Test Journal",
      ACCT_NO: "3101", LOCATION_ID: "500", DEPT_ID: "501", MEMO: "Test Debit",
      DEBIT: baseDebit.toFixed(2), CREDIT: "", SOURCEENTITY: ""
    },
    {
      DONOTIMPORT: "", LINE_NO: "2", DOCUMENT: "", JOURNAL: "PYRJ", 
      DATE: "28/07/2025", REVERSEDATE: "", DESCRIPTION: "Test Journal",
      ACCT_NO: "9431", LOCATION_ID: "500", DEPT_ID: "", MEMO: "Test Credit",
      DEBIT: "", CREDIT: baseCredit.toFixed(2), SOURCEENTITY: ""
    }
  ];
}
