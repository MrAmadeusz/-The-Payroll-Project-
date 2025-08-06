function showJournalRateCheckSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('12_journalSidebar')
    .setTitle("Export Payroll Journal");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Dispatch journal export by type
 * @param {string} journalType - e.g., 'investment', 'casuals'
 */
function runJournalExport(journalType) {
  try {
    const { loader, transformer, exporter } = getJournalHandlers(journalType);

    const raw = loader();
    const transformed = raw.map(transformer).filter(r => r);
    exporter(transformed);

    return `✅ ${journalType} journal exported with ${transformed.length} rows.`;
  } catch (e) {
    throw new Error(`Export failed for ${journalType}: ${e.message}`);
  }
}

/**
 * Updated getJournalHandlers function to pass selectedMonth parameter
 */

function getJournalHandlers(journalType, selectedMonth, apLevyAmount) {
  switch (journalType) {
    case "investment":
      return {
        loader: getRawInvestmentJournal,
        transformer: processInvestmentJournalRow,
        exporter: (data) => outputTransformedInvestmentJournal(data, selectedMonth)
      };
    case "hourly":
      return {
        loader: getRawHourlyJournal,
        transformer: processHourlyJournalRow,
        exporter: (data) => outputTransformedHourlyJournal(data, selectedMonth)
      };
    case "salaried":
      return {
        loader: getRawSalariedJournal,
        transformer: processSalariedJournalRow,
        exporter: (data) => outputTransformedSalariedJournal(data, selectedMonth)
      };
    case "hourlyAccrual":
      return {
        loader: () => [], // Not used
        transformer: r => r, // Not used
        exporter: () => runHourlyAccrualExport()
      };
    case "crossCharge":
      return {
        loader: () => getRawCrossChargeJournal(selectedMonth),
        transformer: row => processCrossChargeJournalRow(row, selectedMonth),
        exporter: rows => outputTransformedCrossChargeJournal(rows, selectedMonth)
      };
    case "hourlyAccrual":
      return {
        loader: () => [], // Not used
        transformer: r => r, // Not used
        exporter: () => runHourlyAccrualExport()
      };
    case "apLevy":
      return {
        loader: () => getRawApLevyJournal(selectedMonth),
        transformer: () => {}, // Not used for AP Levy
        exporter: (niData) => {
          const journalEntries = processApLevyJournalRows(niData, apLevyAmount, selectedMonth);
          return outputTransformedApLevyJournal(journalEntries, selectedMonth);
        }
      };
    case "ptClasses":
      return {
        loader: () => {}, // Not used - PT/Classes has its own processing flow
        transformer: () => {}, // Not used
        exporter: () => processPTClassesSeparation(selectedMonth)
      };
    default:
      throw new Error("Unsupported journal type: " + journalType);
  }
}

/**
 * Process hourly accrual with complete error handling
 */
function processHourlyAccrualComplete() {
  try {
    const file = runHourlyAccrualExport();
    return {
      success: true,
      message: `Accrual file copied: ${file.getName()}`,
      file: file
    };
  } catch (error) {
    throw new Error(`Hourly accrual processing failed: ${error.message}`);
  }
}

/**
 * Main function to handle multiple journal export with AP Levy prompt support
 * @param {Array<string>} journalTypes - Selected journal types from UI
 * @param {string} selectedMonth - Selected payroll month
 * @returns {string} Combined results from all journal processing
 */
function runMultiJournalExport(journalTypes, selectedMonth) {
  // Handle AP Levy prompt if needed
  let apLevyAmount = null;
  if (journalTypes.includes('apLevy')) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Apprenticeship Levy Input',
      `Enter the total Apprenticeship Levy amount for ${selectedMonth}:\n(Enter as decimal, e.g., 1500.00)`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      throw new Error('AP Levy processing cancelled by user');
    }
    
    const inputText = response.getResponseText().trim();
    apLevyAmount = parseFloat(inputText);
    
    // Validation
    if (isNaN(apLevyAmount) || apLevyAmount <= 0) {
      throw new Error(`Invalid AP Levy amount: "${inputText}". Please enter a positive number.`);
    }
    
    // Additional validation for reasonable range (optional safety check)
    if (apLevyAmount > 100000) {
      const confirmResponse = ui.alert(
        'Large AP Levy Amount',
        `The entered amount (£${apLevyAmount.toLocaleString()}) seems very large. Continue?`,
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResponse !== ui.Button.YES) {
        throw new Error('AP Levy processing cancelled due to large amount');
      }
    }
  }

  // Process each journal type
  return journalTypes.map(type => {
    try {
      const { loader, transformer, exporter } = getJournalHandlers(type, selectedMonth, apLevyAmount);
      
      // Special handling for hourly accrual (doesn't follow standard pattern)
      if (type === 'hourlyAccrual') {
        const file = exporter();
        return `✅ ${type}: ${file.getName()} - ${file.getUrl()}`;
      }
      
      // Special handling for AP Levy (processes entire array, not row by row)
      if (type === 'apLevy') {
        const raw = loader();
        const file = exporter(raw); // Pass raw data directly to exporter
        return `✅ ${type}: AP Levy allocated across ${raw.length} NI entries - ${file.getName()} - ${file.getUrl()}`;
      }
      
      // Special handling for PT/Classes (has its own complete processing flow)
      if (type === 'ptClasses') {
        const result = exporter(); // Returns full result object from processPTClassesSeparation
        if (result.success) {
          let message = `✅ ${type}: 3 files created\n`;
          message += `   • ${result.salariedResult.adjustedJournal.getName()}\n`;
          message += `   • ${result.hourlyResult.adjustedJournal.getName()}\n`;
          message += `   • ${result.combinedJournal.file.getName()}\n`;
          message += `   💰 Total NI allocated: £${(result.salariedResult.totalNIAllocated + result.hourlyResult.totalNIAllocated).toLocaleString()}`;
          message += `   ⚖️ Validation: ${result.validation.isValid ? 'PASSED' : 'FAILED'}`;
          return message;
        } else {
          return `❌ ${type} failed: ${result.error}`;
        }
      }
      
      // Standard processing for other journal types
      const raw = loader();
      const transformed = raw.map(transformer).filter(r => r);
      const file = exporter(transformed);
      
      // Enhanced success message with row count and file info
      return `✅ ${type}: ${transformed.length} rows - ${file.getName()} - ${file.getUrl()}`;
      
    } catch (e) {
      // Enhanced error reporting
      console.error(`Journal export failed for ${type}:`, e);
      return `❌ ${type} failed: ${e.message}`;
    }
  }).join("\n");
}

/**
 * Test function for AP Levy prompt (development/debugging only)
 */
function testApLevyPrompt() {
  try {
    const result = runMultiJournalExport(['apLevy'], 'April 2025');
    console.log('Test result:', result);
    return result;
  } catch (error) {
    console.error('Test failed:', error);
    return `Test failed: ${error.message}`;
  }
}

/**
 * Test function for PT/Classes separation (development/debugging only)
 */
function testPTClassesSeparation() {
  try {
    const result = runMultiJournalExport(['ptClasses'], 'June 2025');
    console.log('PT/Classes test result:', result);
    return result;
  } catch (error) {
    console.error('PT/Classes test failed:', error);
    return `PT/Classes test failed: ${error.message}`;
  }
}

/**
 * Validate journal type selection (helper function)
 * @param {Array<string>} journalTypes - Selected journal types
 * @returns {Object} Validation result
 */
function validateJournalSelection(journalTypes) {
  const validTypes = ['investment', 'hourly', 'salaried', 'hourlyAccrual', 'crossCharge', 'apLevy', 'ptClasses'];
  const invalidTypes = journalTypes.filter(type => !validTypes.includes(type));
  
  return {
    isValid: invalidTypes.length === 0,
    invalidTypes: invalidTypes,
    validTypes: journalTypes.filter(type => validTypes.includes(type))
  };
}

/**
 * Get journal processing summary (helper function for reporting)
 * @param {Array<string>} journalTypes - Journal types being processed
 * @param {string} selectedMonth - Selected month
 * @param {number} apLevyAmount - AP Levy amount if applicable
 * @returns {Object} Processing summary
 */
function getJournalProcessingSummary(journalTypes, selectedMonth, apLevyAmount = null) {
  const validation = validateJournalSelection(journalTypes);
  
  return {
    selectedMonth: selectedMonth,
    journalTypes: journalTypes,
    validJournalTypes: validation.validTypes,
    invalidJournalTypes: validation.invalidTypes,
    includesApLevy: journalTypes.includes('apLevy'),
    includesPTClasses: journalTypes.includes('ptClasses'),
    apLevyAmount: apLevyAmount,
    totalJournals: journalTypes.length,
    processingOrder: journalTypes,
    estimatedProcessingTime: journalTypes.length * 15 // Increased estimate for PT/Classes complexity
  };
}

/**
 * Comprehensive test function for all journal types
 */
function testAllJournalTypes() {
  try {
    console.log('🧪 Testing all journal types...');
    
    const testMonth = 'June 2025';
    const journalTypes = ['investment', 'hourly', 'salaried', 'hourlyAccrual', 'crossCharge', 'ptClasses'];
    
    const results = {};
    
    journalTypes.forEach(type => {
      try {
        console.log(`Testing ${type}...`);
        const result = runMultiJournalExport([type], testMonth);
        results[type] = { success: true, message: result };
        console.log(`✅ ${type}: SUCCESS`);
      } catch (error) {
        results[type] = { success: false, error: error.message };
        console.log(`❌ ${type}: ${error.message}`);
      }
    });
    
    console.log('🎉 All journal type tests completed');
    return results;
    
  } catch (error) {
    console.error('❌ Test all journal types failed:', error);
    return { error: error.message };
  }
}
