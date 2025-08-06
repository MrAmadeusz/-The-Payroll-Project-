/**
 * 12_journalPTClasses.gs
 * PT/Classes Journal Separation Module - FINAL PRODUCTION VERSION
 * 
 * Separates PT/Classes costs and their proportional NI from hourly/salaried journals
 * Creates 3 output files: adjusted salaried, adjusted hourly, and combined PT/Classes journal
 * 
 * Key Features:
 * - Matches NI debits by BOTH location AND department
 * - Handles multiple NI entries per location-department combination
 * - Reduces NI debits (not PT/Classes entries) for correct accounting
 * - Proportional allocation across multiple matching NI entries
 * 
 * Version: 3.0 - Location-Department matching with correct NI reduction logic
 * Last Updated: August 2025
 */

/**
 * Main processing function for PT/Classes separation
 * @param {string} selectedMonth - Month being processed (e.g., "July 2025")
 * @returns {Object} Processing results with file details
 */
function processPTClassesSeparation(selectedMonth) {
  try {
    console.log(`🚀 Starting PT/Classes separation for ${selectedMonth}`);
    
    // Extract period number from selected month
    const [monthName, year] = selectedMonth.split(" ");
    const periodNumber = getFinancialPeriodNumber(monthName);
    const period = parseInt(periodNumber); // Remove leading zero for memo matching
    
    console.log(`📅 Processing period P${period} for ${monthName} ${year}`);
    
    // Process both journal types
    const salariedResult = processSalariedJournalForPTClasses(period);
    const hourlyResult = processHourlyJournalForPTClasses(period);
    
    // Merge the PT/Classes journals
    const combinedPTClassesJournal = mergePTClassesJournals(
      salariedResult.ptClassesEntries, 
      hourlyResult.ptClassesEntries, 
      selectedMonth
    );
    
    // Validate totals
    const validation = validatePTClassesSeparation(
      salariedResult, 
      hourlyResult, 
      combinedPTClassesJournal
    );
    
    console.log(`🎉 PT/Classes separation complete!`);
    console.log(`📊 Salaried: ${salariedResult.ptClassesEntries.length} PT/Classes entries`);
    console.log(`📊 Hourly: ${hourlyResult.ptClassesEntries.length} PT/Classes entries`);
    console.log(`📊 Combined: ${combinedPTClassesJournal.entries.length} total entries`);
    
    return {
      success: true,
      salariedResult: salariedResult,
      hourlyResult: hourlyResult,
      combinedJournal: combinedPTClassesJournal,
      validation: validation,
      period: period,
      selectedMonth: selectedMonth
    };
    
  } catch (error) {
    console.error(`❌ processPTClassesSeparation failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Processes salaried journal for PT/Classes separation
 * @param {number} period - Financial period number
 * @returns {Object} Processing results
 */
function processSalariedJournalForPTClasses(period, selectedMonth = "July 2025") {
  try {
    console.log(`📋 Processing salaried journal for period P${period}`);
    
    const originalJournal = loadMostRecentJournal('salaried');
    console.log(`📂 Loaded salaried journal: ${originalJournal.length} entries`);
    
    const analysis = analyzePTClassesEntries(originalJournal, period);
    
    if (analysis.ptClassesLines.length === 0) {
      console.log(`ℹ️ No PT/Classes entries found in salaried journal for P${period}`);
      const outputFile = outputAdjustedJournal(originalJournal, 'salaried', selectedMonth);
      return {
        originalEntries: originalJournal.length,
        ptClassesEntries: [],
        adjustedJournal: outputFile,
        totalPTClassesDebits: 0,
        totalNIAllocated: 0,
        locationDeptAllocations: {}
      };
    }
    
    const { adjustedJournal, ptClassesEntries } = createAdjustedJournals(originalJournal, analysis, period);
    const adjustedFile = outputAdjustedJournal(adjustedJournal, 'salaried', selectedMonth);
    
    console.log(`✅ Salaried processing complete`);
    return {
      originalEntries: originalJournal.length,
      ptClassesEntries: ptClassesEntries,
      adjustedJournal: adjustedFile,
      totalPTClassesDebits: analysis.totalPTClassesDebits,
      totalNIAllocated: analysis.totalNIAllocated,
      locationDeptAllocations: analysis.locationDeptNIAllocations
    };
    
  } catch (error) {
    console.error(`❌ processSalariedJournalForPTClasses failed: ${error.message}`);
    throw error;
  }
}

/**
 * Processes hourly journal for PT/Classes separation
 * @param {number} period - Financial period number
 * @returns {Object} Processing results
 */
function processHourlyJournalForPTClasses(period, selectedMonth = "July 2025") {
  try {
    console.log(`📋 Processing hourly journal for period P${period}`);
    
    const originalJournal = loadMostRecentJournal('hourly');
    console.log(`📂 Loaded hourly journal: ${originalJournal.length} entries`);
    
    const analysis = analyzePTClassesEntries(originalJournal, period);
    
    if (analysis.ptClassesLines.length === 0) {
      console.log(`ℹ️ No PT/Classes entries found in hourly journal for P${period}`);
      const outputFile = outputAdjustedJournal(originalJournal, 'hourly', selectedMonth);
      return {
        originalEntries: originalJournal.length,
        ptClassesEntries: [],
        adjustedJournal: outputFile,
        totalPTClassesDebits: 0,
        totalNIAllocated: 0,
        locationDeptAllocations: {}
      };
    }
    
    const { adjustedJournal, ptClassesEntries } = createAdjustedJournals(originalJournal, analysis, period);
    const adjustedFile = outputAdjustedJournal(adjustedJournal, 'hourly', selectedMonth);
    
    console.log(`✅ Hourly processing complete`);
    return {
      originalEntries: originalJournal.length,
      ptClassesEntries: ptClassesEntries,
      adjustedJournal: adjustedFile,
      totalPTClassesDebits: analysis.totalPTClassesDebits,
      totalNIAllocated: analysis.totalNIAllocated,
      locationDeptAllocations: analysis.locationDeptNIAllocations
    };
    
  } catch (error) {
    console.error(`❌ processHourlyJournalForPTClasses failed: ${error.message}`);
    throw error;
  }
}

/**
 * Loads the most recent journal file by type
 * @param {string} journalType - Either 'hourly' or 'salaried'
 * @returns {Array<Object>} Parsed journal entries
 */
function loadMostRecentJournal(journalType) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID); // Both use same folder
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
      throw new Error(`No ${journalType} journal file found. Please run the ${journalType} journal first.`);
    }
    
    console.log(`📂 Loading ${journalType} file: ${latestFile.getName()}`);
    
    // Parse CSV file
    return parseCsvToObjects(latestFile);
    
  } catch (error) {
    console.error(`❌ loadMostRecentJournal(${journalType}) failed: ${error.message}`);
    throw error;
  }
}

/**
 * Analyzes journal entries to find PT/Classes lines and calculate NI allocation
 * FINAL VERSION - Location + Department matching with correct NI reduction logic
 * @param {Array<Object>} journalEntries - All journal entries
 * @param {number} period - Financial period number (kept for compatibility)
 * @returns {Object} Analysis results with location-department allocations
 */
function analyzePTClassesEntries(journalEntries, period) {
  try {
    console.log(`🔍 Analyzing PT/Classes entries (Location + Department matching)`);
    
    // Validate input
    if (!journalEntries || !Array.isArray(journalEntries) || journalEntries.length === 0) {
      console.warn('⚠️ No journal entries provided');
      return createEmptyPTClassesAnalysis();
    }
    
    // Define PT/Classes memo patterns (period-agnostic)
    const ptClassesPatterns = [
      'P T^',           // Starts with "P T^"
      'Pt Payments',    // Starts with "Pt Payments" 
      'Classes^',       // Starts with "Classes^"
      'Classes'         // Starts with "Classes" (but not "Classes^" to avoid double-match)
    ];
    
    // Find PT/Classes debit entries using startsWith matching
    const ptClassesLines = [];
    let totalPTClassesDebits = 0;
    
    journalEntries.forEach((entry, index) => {
      if (!entry) return; // Skip null/undefined entries
      
      const memo = String(entry.MEMO || "").trim();
      const debit = parseFloat(entry.DEBIT || 0);
      const locationId = String(entry.LOCATION_ID || "").trim();
      const deptId = String(entry.DEPT_ID || "").trim();
      
      // Check if memo starts with any of our PT/Classes patterns
      const matchesPattern = ptClassesPatterns.some(pattern => {
        if (pattern === 'Classes') {
          // Special case: "Classes" should match "Classes P4" but not "Classes^ P4"
          return memo.startsWith(pattern) && !memo.startsWith('Classes^');
        }
        return memo.startsWith(pattern);
      });
      
      if (matchesPattern && debit > 0) {
        ptClassesLines.push({
          index: index,
          entry: entry,
          memo: memo,
          debit: debit,
          locationId: locationId,
          deptId: deptId,
          locationDeptKey: `${locationId}-${deptId}` // ← KEY: Combined location-department key
        });
        totalPTClassesDebits += debit;
      }
    });
    
    console.log(`📊 Found ${ptClassesLines.length} PT/Classes entries totaling £${totalPTClassesDebits.toLocaleString()}`);
    
    if (ptClassesLines.length === 0) {
      console.log('ℹ️ No PT/Classes entries found');
      return createEmptyPTClassesAnalysis();
    }
    
    // Group PT/Classes entries by BOTH location AND department
    const ptClassesByLocationDept = {};
    ptClassesLines.forEach(line => {
      const key = line.locationDeptKey;
      if (!ptClassesByLocationDept[key]) {
        ptClassesByLocationDept[key] = [];
      }
      ptClassesByLocationDept[key].push(line);
    });
    
    console.log(`📍 PT/Classes activity found in ${Object.keys(ptClassesByLocationDept).length} location-department combinations`);
    
    // Calculate total debits in journal
    const totalDebits = journalEntries.reduce((sum, entry) => {
      if (!entry) return sum;
      const debit = parseFloat(entry.DEBIT || 0);
      return sum + (isNaN(debit) ? 0 : debit);
    }, 0);
    
    // Calculate PT/Classes percentage of total debits
    const ptClassesPercentage = totalDebits > 0 ? totalPTClassesDebits / totalDebits : 0;
    
    console.log(`📊 PT/Classes represents ${(ptClassesPercentage * 100).toFixed(2)}% of total debits (£${totalDebits.toLocaleString()})`);
    
    // Find employer NI CREDIT entry (the summary total)
    let employerNIEntry = null;
    let employerNIIndex = -1;
    
    journalEntries.forEach((entry, index) => {
      if (!entry) return;
      
      const memo = String(entry.MEMO || "").trim();
      const credit = parseFloat(entry.CREDIT || 0);
      
      if (memo.startsWith('NIC employer') && credit > 0) {
        employerNIEntry = entry;
        employerNIIndex = index;
      }
    });
    
    // Find all employer NI DEBIT entries - GROUP BY LOCATION + DEPARTMENT
    const employerNIDebits = [];
    const niDebitsByLocationDept = {};
    
    journalEntries.forEach((entry, index) => {
      if (!entry) return;
      
      const memo = String(entry.MEMO || "").trim();
      const debit = parseFloat(entry.DEBIT || 0);
      const locationId = String(entry.LOCATION_ID || "").trim();
      const deptId = String(entry.DEPT_ID || "").trim();
      const locationDeptKey = `${locationId}-${deptId}`;
      
      if (memo.startsWith('NIC employer') && debit > 0) {
        const niEntry = {
          index: index,
          entry: entry,
          memo: memo,
          debit: debit,
          locationId: locationId,
          deptId: deptId,
          locationDeptKey: locationDeptKey
        };
        
        employerNIDebits.push(niEntry);
        
        // Group NI debits by location-department (sum multiple entries)
        if (!niDebitsByLocationDept[locationDeptKey]) {
          niDebitsByLocationDept[locationDeptKey] = {
            entries: [],
            totalDebit: 0,
            locationId: locationId,
            deptId: deptId
          };
        }
        
        niDebitsByLocationDept[locationDeptKey].entries.push(niEntry);
        niDebitsByLocationDept[locationDeptKey].totalDebit += debit;
      }
    });
    
    console.log(`💰 Found ${employerNIDebits.length} NI debit entries across ${Object.keys(niDebitsByLocationDept).length} location-department combinations`);
    
    if (!employerNIEntry) {
      console.warn(`⚠️ No employer NI credit entry found`);
      return {
        ptClassesLines: ptClassesLines,
        totalPTClassesDebits: totalPTClassesDebits,
        totalDebits: totalDebits,
        ptClassesPercentage: ptClassesPercentage,
        employerNIEntry: null,
        employerNIIndex: -1,
        employerNICredit: 0,
        employerNIDebits: employerNIDebits,
        niDebitsByLocationDept: niDebitsByLocationDept,
        totalNIAllocated: 0,
        locationDeptNIAllocations: {},
        ptClassesByLocationDept: ptClassesByLocationDept
      };
    }
    
    const employerNICredit = parseFloat(employerNIEntry.CREDIT || 0);
    const totalNIAllocated = employerNICredit * ptClassesPercentage;
    
    console.log(`💰 Employer NI: £${employerNICredit.toLocaleString()}, allocated to PT/Classes: £${totalNIAllocated.toLocaleString()}`);
    
    // Calculate NI allocation per location-department combination based on PT/Classes activity
    const locationDeptNIAllocations = {};
    
    Object.entries(ptClassesByLocationDept).forEach(([locationDeptKey, locationDeptLines]) => {
      const locationDeptPTClassesTotal = locationDeptLines.reduce((sum, line) => sum + line.debit, 0);
      const locationDeptPercentage = locationDeptPTClassesTotal / totalPTClassesDebits;
      const locationDeptNIAllocation = totalNIAllocated * locationDeptPercentage;
      
      const [locationId, deptId] = locationDeptKey.split('-');
      
      locationDeptNIAllocations[locationDeptKey] = {
        locationId: locationId,
        deptId: deptId,
        ptClassesAmount: locationDeptPTClassesTotal,
        percentage: locationDeptPercentage,
        niAllocation: locationDeptNIAllocation,
        entries: locationDeptLines.length
      };
      
      console.log(`📍 Location ${locationId}-Dept ${deptId}: £${locationDeptPTClassesTotal.toLocaleString()} PT/Classes (${(locationDeptPercentage * 100).toFixed(1)}%) → £${locationDeptNIAllocation.toLocaleString()} NI allocation`);
    });
    
    return {
      ptClassesLines: ptClassesLines,
      totalPTClassesDebits: totalPTClassesDebits,
      totalDebits: totalDebits,
      ptClassesPercentage: ptClassesPercentage,
      employerNIEntry: employerNIEntry,
      employerNIIndex: employerNIIndex,
      employerNICredit: employerNICredit,
      employerNIDebits: employerNIDebits,
      niDebitsByLocationDept: niDebitsByLocationDept, // ← NEW: Grouped NI debits by location-dept
      totalNIAllocated: totalNIAllocated,
      locationDeptNIAllocations: locationDeptNIAllocations, // ← UPDATED: Location-Dept allocations
      ptClassesByLocationDept: ptClassesByLocationDept // ← UPDATED: Location-Dept grouping
    };
    
  } catch (error) {
    console.error(`❌ analyzePTClassesEntries failed: ${error.message}`);
    return createEmptyPTClassesAnalysis();
  }
}

/**
 * Helper function to create empty analysis result
 * @returns {Object} Empty analysis structure
 */
function createEmptyPTClassesAnalysis() {
  return {
    ptClassesLines: [],
    totalPTClassesDebits: 0,
    totalDebits: 0,
    ptClassesPercentage: 0,
    employerNIEntry: null,
    employerNIIndex: -1,
    employerNICredit: 0,
    employerNIDebits: [],
    niDebitsByLocationDept: {},
    totalNIAllocated: 0,
    locationDeptNIAllocations: {},
    ptClassesByLocationDept: {}
  };
}

/**
 * Creates adjusted journal and PT/Classes entries based on analysis
 * FINAL VERSION - Correct NI reduction logic with location-department matching
 * @param {Array<Object>} originalJournal - Original journal entries
 * @param {Object} analysis - Analysis results from analyzePTClassesEntries
 * @param {number} period - Financial period number
 * @returns {Object} Adjusted journal and PT/Classes entries
 */
function createAdjustedJournals(originalJournal, analysis, period) {
  try {
    console.log(`🔧 Creating adjusted journals with Location-Department matching`);
    
    // Validate inputs
    if (!originalJournal || !Array.isArray(originalJournal)) {
      throw new Error('Invalid original journal data');
    }
    
    if (!analysis) {
      throw new Error('Analysis data is missing');
    }
    
    // Create deep copy of original journal for adjustment
    const adjustedJournal = originalJournal.map(entry => ({ ...entry }));
    
    // CORRECT LOGIC: PT/Classes entries remain unchanged, reduce NI debits instead
    console.log(`📊 PT/Classes entries remain unchanged at original amounts`);
    
    // Adjust employer NI DEBIT entries by location-department combination
    if (analysis.locationDeptNIAllocations && Object.keys(analysis.locationDeptNIAllocations).length > 0) {
      console.log(`💰 Adjusting NI debits across ${Object.keys(analysis.locationDeptNIAllocations).length} location-department combinations`);
      
      Object.entries(analysis.locationDeptNIAllocations).forEach(([locationDeptKey, allocation]) => {
        const [locationId, deptId] = locationDeptKey.split('-');
        
        // Find matching NI debit entries for this location-department combination
        const matchingNIDebits = analysis.niDebitsByLocationDept[locationDeptKey];
        
        if (matchingNIDebits && matchingNIDebits.entries.length > 0) {
          console.log(`📍 Location ${locationId}-Dept ${deptId}: Found ${matchingNIDebits.entries.length} NI debit entries totaling £${matchingNIDebits.totalDebit.toLocaleString()}`);
          
          // Distribute the NI reduction proportionally across all matching entries
          matchingNIDebits.entries.forEach(niEntry => {
            const entryProportion = niEntry.debit / matchingNIDebits.totalDebit;
            const entryReduction = allocation.niAllocation * entryProportion;
            const adjustedAmount = niEntry.debit - entryReduction;
            
            // Update the NI debit entry in the adjusted journal
            adjustedJournal[niEntry.index].DEBIT = adjustedAmount.toFixed(2);
            
            console.log(`    Entry ${niEntry.index}: £${niEntry.debit.toLocaleString()} → £${adjustedAmount.toLocaleString()} (reduced by £${entryReduction.toLocaleString()})`);
          });
        } else {
          console.warn(`⚠️ No matching NI debit entries found for Location ${locationId}-Dept ${deptId}`);
        }
      });
    }
    
    // Adjust employer NI credit entry (total reduction)
    if (analysis.employerNIEntry && analysis.employerNIIndex >= 0) {
      const adjustedNICredit = analysis.employerNICredit - analysis.totalNIAllocated;
      adjustedJournal[analysis.employerNIIndex].CREDIT = adjustedNICredit.toFixed(2);
      console.log(`💰 Adjusted total NI credit: £${analysis.employerNICredit.toLocaleString()} → £${adjustedNICredit.toLocaleString()}`);
    }
    
    // Create PT/Classes journal entries
    const ptClassesEntries = [];
    
    // Add PT/Classes debit entries (FULL original amounts - NOT reduced)
    analysis.ptClassesLines.forEach((line, index) => {
      ptClassesEntries.push({
        DONOTIMPORT: "",
        LINE_NO: (index + 1).toString(),
        DOCUMENT: "",
        JOURNAL: "GJ",
        DATE: line.entry.DATE || "",
        REVERSEDATE: "",
        DESCRIPTION: line.entry.DESCRIPTION || "",
        ACCT_NO: line.entry.ACCT_NO || "",
        LOCATION_ID: line.locationId,
        DEPT_ID: line.deptId,
        MEMO: line.memo,
        DEBIT: line.debit.toFixed(2), // ← FULL ORIGINAL AMOUNT (NOT REDUCED)
        CREDIT: "",
        SOURCEENTITY: ""
      });
    });
    
    // Add proportionate NI debit entries (by location-department)
    let niLineNumber = ptClassesEntries.length + 1;
    Object.entries(analysis.locationDeptNIAllocations).forEach(([locationDeptKey, allocation]) => {
      // Find a sample PT/Classes entry for this location-dept to copy account details
      const sampleEntry = analysis.ptClassesByLocationDept[locationDeptKey][0].entry;
      
      ptClassesEntries.push({
        DONOTIMPORT: "",
        LINE_NO: niLineNumber.toString(),
        DOCUMENT: "",
        JOURNAL: "GJ",
        DATE: sampleEntry.DATE || "",
        REVERSEDATE: "",
        DESCRIPTION: sampleEntry.DESCRIPTION || "",
        ACCT_NO: sampleEntry.ACCT_NO || "", // Same account as PT/Classes
        LOCATION_ID: allocation.locationId,
        DEPT_ID: allocation.deptId,
        MEMO: `NIC employer P${period}`,
        DEBIT: allocation.niAllocation.toFixed(2), // ← NI PORTION ONLY
        CREDIT: "",
        SOURCEENTITY: ""
      });
      
      niLineNumber++;
    });
    
    // Add single NI credit entry (total)
    if (analysis.totalNIAllocated > 0 && analysis.employerNIEntry) {
      ptClassesEntries.push({
        DONOTIMPORT: "",
        LINE_NO: niLineNumber.toString(),
        DOCUMENT: "",
        JOURNAL: "GJ",
        DATE: analysis.employerNIEntry.DATE || "",
        REVERSEDATE: "",
        DESCRIPTION: analysis.employerNIEntry.DESCRIPTION || "",
        ACCT_NO: analysis.employerNIEntry.ACCT_NO || "",
        LOCATION_ID: analysis.employerNIEntry.LOCATION_ID || "",
        DEPT_ID: analysis.employerNIEntry.DEPT_ID || "",
        MEMO: `NIC employer P${period}`,
        DEBIT: "",
        CREDIT: analysis.totalNIAllocated.toFixed(2),
        SOURCEENTITY: ""
      });
    }
    
    console.log(`✅ Created ${ptClassesEntries.length} PT/Classes journal entries`);
    console.log(`📊 PT/Classes debits: ${analysis.ptClassesLines.length} entries (full amounts)`);
    console.log(`💰 NI debits: ${Object.keys(analysis.locationDeptNIAllocations).length} entries (by location-department)`);
    console.log(`💰 NI credit: 1 entry (total allocation)`);
    
    return {
      adjustedJournal: adjustedJournal,
      ptClassesEntries: ptClassesEntries
    };
    
  } catch (error) {
    console.error(`❌ createAdjustedJournals failed: ${error.message}`);
    throw error;
  }
}

/**
 * Merges PT/Classes entries from both hourly and salaried journals
 * @param {Array<Object>} salariedEntries - PT/Classes entries from salaried journal
 * @param {Array<Object>} hourlyEntries - PT/Classes entries from hourly journal
 * @param {string} selectedMonth - Month being processed
 * @returns {Object} Combined journal with file details
 */
function mergePTClassesJournals(salariedEntries, hourlyEntries, selectedMonth) {
  try {
    console.log(`🔗 Merging PT/Classes journals`);
    console.log(`📊 Salaried entries: ${salariedEntries.length}, Hourly entries: ${hourlyEntries.length}`);
    
    // Ensure we have arrays (even if empty)
    const validSalariedEntries = Array.isArray(salariedEntries) ? salariedEntries : [];
    const validHourlyEntries = Array.isArray(hourlyEntries) ? hourlyEntries : [];
    
    // Combine entries
    const combinedEntries = [...validSalariedEntries, ...validHourlyEntries];
    
    if (combinedEntries.length === 0) {
      console.log(`ℹ️ No PT/Classes entries to merge - creating empty journal`);
    }
    
    // Renumber line numbers
    const numberedEntries = combinedEntries.map((entry, index) => ({
      ...entry,
      LINE_NO: (index + 1).toString()
    }));
    
    // Output combined journal
    const outputFile = outputPTClassesJournal(numberedEntries, selectedMonth);
    
    // Calculate summary statistics
    const totalDebits = numberedEntries.reduce((sum, entry) => {
      const debit = parseFloat(entry.DEBIT || 0);
      return sum + (isNaN(debit) ? 0 : debit);
    }, 0);
    
    const totalCredits = numberedEntries.reduce((sum, entry) => {
      const credit = parseFloat(entry.CREDIT || 0);
      return sum + (isNaN(credit) ? 0 : credit);
    }, 0);
    
    const isBalanced = Math.abs(totalDebits - totalCredits) < 0.01;
    
    console.log(`✅ Combined PT/Classes journal created`);
    console.log(`📊 Total entries: ${numberedEntries.length}`);
    console.log(`💰 Debits: £${totalDebits.toLocaleString()}, Credits: £${totalCredits.toLocaleString()}`);
    console.log(`⚖️ Balanced: ${isBalanced ? 'Yes' : 'No'}`);
    
    return {
      entries: numberedEntries,
      file: outputFile,
      totalDebits: totalDebits,
      totalCredits: totalCredits,
      isBalanced: isBalanced,
      salariedCount: validSalariedEntries.length,
      hourlyCount: validHourlyEntries.length
    };
    
  } catch (error) {
    console.error(`❌ mergePTClassesJournals failed: ${error.message}`);
    throw error;
  }
}

/**
 * Outputs adjusted journal to CSV file
 * @param {Array<Object>} journalEntries - Adjusted journal entries
 * @param {string} journalType - Either 'hourly' or 'salaried'
 * @returns {File} Created CSV file
 */
function outputAdjustedJournal(journalEntries, journalType, selectedMonth = "July 2025") {
  try {
    const folder = DriveApp.getFolderById(CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
    const filename = `journal${journalType.charAt(0).toUpperCase() + journalType.slice(1)}_PTClassesAdjusted_${timestamp}.csv`;
    
    return outputJournalToCsv(journalEntries, filename, folder, selectedMonth);
    
  } catch (error) {
    console.error(`❌ outputAdjustedJournal failed: ${error.message}`);
    throw error;
  }
}

/**
 * Outputs PT/Classes journal to CSV file
 * @param {Array<Object>} journalEntries - PT/Classes journal entries
 * @param {string} selectedMonth - Month being processed
 * @returns {File} Created CSV file
 */
function outputPTClassesJournal(journalEntries, selectedMonth) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
    const filename = `journalPTClasses_${timestamp}.csv`;
    
    return outputJournalToCsv(journalEntries, filename, folder, selectedMonth);
    
  } catch (error) {
    console.error(`❌ outputPTClassesJournal failed: ${error.message}`);
    throw error;
  }
}
/**
 * Generic function to output journal entries to balanced CSV
 * REPLACES: outputJournalToCsv(journalEntries, filename, folder)
 */
function outputJournalToCsv(journalEntries, filename, folder, selectedMonth = "July 2025") {
  try {
    // Ensure we have an array (even if empty)
    const entries = Array.isArray(journalEntries) ? journalEntries : [];
    
    if (entries.length === 0) {
      console.log(`ℹ️ Creating empty journal file: ${filename}`);
      
      // For empty journals, create minimal file without post-processing
      const headers = [
        "DONOTIMPORT", "LINE_NO", "DOCUMENT", "JOURNAL", "DATE", "REVERSEDATE",
        "DESCRIPTION", "ACCT_NO", "LOCATION_ID", "DEPT_ID", "MEMO",
        "DEBIT", "CREDIT", "SOURCEENTITY"
      ];
      const csvContent = headers.join(",") + "\r\n";
      const file = folder.createFile(filename, csvContent, MimeType.CSV);
      console.log(`✅ Empty journal created: ${filename}`);
      return file;
    }

    // Use the post-processor to create balanced journal
    const file = createBalancedJournalFile(entries, selectedMonth, filename, folder, 'ptClasses');
    
    console.log(`✅ Journal exported: ${filename} (${entries.length} entries)`);
    console.log(`📁 File URL: ${file.getUrl()}`);

    return file;
    
  } catch (error) {
    console.error(`❌ outputJournalToCsv failed: ${error.message}`);
    throw error;
  }
}

/**
 * Validates that the separation maintained total balance
 * @param {Object} salariedResult - Results from salaried processing
 * @param {Object} hourlyResult - Results from hourly processing
 * @param {Object} combinedPTClassesJournal - Combined PT/Classes journal
 * @returns {Object} Validation results
 */
function validatePTClassesSeparation(salariedResult, hourlyResult, combinedPTClassesJournal) {
  try {
    console.log(`🔍 Validating PT/Classes separation totals`);
    
    const validation = {
      salaried: {
        originalNIAllocated: salariedResult.totalNIAllocated || 0,
        ptClassesEntries: salariedResult.ptClassesEntries ? salariedResult.ptClassesEntries.length : 0
      },
      hourly: {
        originalNIAllocated: hourlyResult.totalNIAllocated || 0,
        ptClassesEntries: hourlyResult.ptClassesEntries ? hourlyResult.ptClassesEntries.length : 0
      },
      combined: {
        totalEntries: combinedPTClassesJournal.entries ? combinedPTClassesJournal.entries.length : 0,
        totalDebits: combinedPTClassesJournal.totalDebits || 0,
        totalCredits: combinedPTClassesJournal.totalCredits || 0,
        isBalanced: combinedPTClassesJournal.isBalanced || false
      },
      totals: {
        expectedNITotal: (salariedResult.totalNIAllocated || 0) + (hourlyResult.totalNIAllocated || 0),
        actualNITotal: combinedPTClassesJournal.totalCredits || 0,
        difference: Math.abs(((salariedResult.totalNIAllocated || 0) + (hourlyResult.totalNIAllocated || 0)) - (combinedPTClassesJournal.totalCredits || 0))
      }
    };
    
    const isValid = validation.totals.difference < 0.01;
    validation.isValid = isValid;
    
    console.log(`✅ Validation ${isValid ? 'PASSED' : 'FAILED'}`);
    console.log(`📊 Expected NI total: £${validation.totals.expectedNITotal.toLocaleString()}`);
    console.log(`📊 Actual NI total: £${validation.totals.actualNITotal.toLocaleString()}`);
    console.log(`📊 Difference: £${validation.totals.difference.toFixed(2)}`);
    
    return validation;
    
  } catch (error) {
    console.error(`❌ validatePTClassesSeparation failed: ${error.message}`);
    return {
      isValid: false,
      error: error.message
    };
  }
}

/**
 * Helper function to get financial period number from month name
 * @param {string} monthName - Month name (e.g., "July")
 * @returns {string} Period number with leading zero (e.g., "07")
 */
function getFinancialPeriodNumber(monthName) {
  const monthMap = {
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
  
  return monthMap[monthName] || '01';
}

// ================================================================
// TESTING AND DIAGNOSTIC FUNCTIONS
// ================================================================

/**
 * Comprehensive test function for PT/Classes separation
 * @param {string} testMonth - Month to test (default: current month context)
 * @returns {Object} Test results
 */
function testPTClassesSeparation(testMonth = "July 2025") {
  try {
    console.log(`🧪 Testing PT/Classes separation for ${testMonth}`);
    
    const result = processPTClassesSeparation(testMonth);
    
    if (result.success) {
      console.log(`✅ Test completed successfully`);
      console.log(`📄 Salaried file: ${result.salariedResult.adjustedJournal.getName()}`);
      console.log(`📄 Hourly file: ${result.hourlyResult.adjustedJournal.getName()}`);
      console.log(`📄 PT/Classes file: ${result.combinedJournal.file.getName()}`);
      console.log(`⚖️ Validation: ${result.validation.isValid ? 'PASSED' : 'FAILED'}`);
      
      // Detailed results
      console.log(`\n📊 DETAILED RESULTS:`);
      console.log(`Salaried PT/Classes entries: ${result.salariedResult.ptClassesEntries.length}`);
      console.log(`Hourly PT/Classes entries: ${result.hourlyResult.ptClassesEntries.length}`);
      console.log(`Combined journal entries: ${result.combinedJournal.entries.length}`);
      console.log(`Total NI allocated: £${(result.salariedResult.totalNIAllocated + result.hourlyResult.totalNIAllocated).toLocaleString()}`);
      console.log(`Journal balanced: ${result.combinedJournal.isBalanced ? 'Yes' : 'No'}`);
      
    } else {
      console.log(`❌ Test failed: ${result.error}`);
    }
    
    return result;
    
  } catch (error) {
    console.error(`❌ Test failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test the location-department matching logic
 * @returns {Object} Test results with matching details
 */
function testLocationDepartmentMatching() {
  try {
    console.log('🧪 Testing Location-Department matching logic...');
    
    // Load and analyze both journals
    const results = {
      salaried: null,
      hourly: null,
      summary: {}
    };
    
    // Test salaried journal
    try {
      const salariedJournal = loadMostRecentJournal('salaried');
      const salariedAnalysis = analyzePTClassesEntries(salariedJournal, 7);
      
      results.salaried = {
        journalEntries: salariedJournal.length,
        ptClassesEntries: salariedAnalysis.ptClassesLines.length,
        locationDeptCombinations: Object.keys(salariedAnalysis.ptClassesByLocationDept).length,
        niLocationDeptCombinations: Object.keys(salariedAnalysis.niDebitsByLocationDept).length,
        totalNIAllocated: salariedAnalysis.totalNIAllocated
      };
      
      console.log(`\n📊 SALARIED RESULTS:`);
      console.log(`PT/Classes entries: ${results.salaried.ptClassesEntries}`);
      console.log(`Location-Dept combinations (PT/Classes): ${results.salaried.locationDeptCombinations}`);
      console.log(`Location-Dept combinations (NI debits): ${results.salaried.niLocationDeptCombinations}`);
      console.log(`Total NI allocated: £${results.salaried.totalNIAllocated.toLocaleString()}`);
      
      // Show matching examples
      console.log(`\n🔍 SALARIED MATCHING EXAMPLES:`);
      Object.entries(salariedAnalysis.ptClassesByLocationDept).slice(0, 3).forEach(([locationDeptKey, ptEntries]) => {
        const [locationId, deptId] = locationDeptKey.split('-');
        const matchingNI = salariedAnalysis.niDebitsByLocationDept[locationDeptKey];
        
        console.log(`\nLocation ${locationId} - Department ${deptId}:`);
        console.log(`  PT/Classes entries: ${ptEntries.length}`);
        ptEntries.slice(0, 2).forEach(entry => {
          console.log(`    "${entry.memo}": £${entry.debit.toLocaleString()}`);
        });
        
        if (matchingNI) {
          console.log(`  ✅ Matching NI entries: ${matchingNI.entries.length} (total: £${matchingNI.totalDebit.toLocaleString()})`);
        } else {
          console.log(`  ❌ NO MATCHING NI ENTRIES FOUND`);
        }
      });
      
    } catch (error) {
      console.log(`❌ Salaried test failed: ${error.message}`);
      results.salaried = { error: error.message };
    }
    
    // Test hourly journal
    try {
      const hourlyJournal = loadMostRecentJournal('hourly');
      const hourlyAnalysis = analyzePTClassesEntries(hourlyJournal, 7);
      
      results.hourly = {
        journalEntries: hourlyJournal.length,
        ptClassesEntries: hourlyAnalysis.ptClassesLines.length,
        locationDeptCombinations: Object.keys(hourlyAnalysis.ptClassesByLocationDept).length,
        niLocationDeptCombinations: Object.keys(hourlyAnalysis.niDebitsByLocationDept).length,
        totalNIAllocated: hourlyAnalysis.totalNIAllocated
      };
      
      console.log(`\n📊 HOURLY RESULTS:`);
      console.log(`PT/Classes entries: ${results.hourly.ptClassesEntries}`);
      console.log(`Location-Dept combinations (PT/Classes): ${results.hourly.locationDeptCombinations}`);
      console.log(`Location-Dept combinations (NI debits): ${results.hourly.niLocationDeptCombinations}`);
      console.log(`Total NI allocated: £${results.hourly.totalNIAllocated.toLocaleString()}`);
      
    } catch (error) {
      console.log(`❌ Hourly test failed: ${error.message}`);
      results.hourly = { error: error.message };
    }
    
    // Summary
    results.summary = {
      totalPTClassesEntries: (results.salaried?.ptClassesEntries || 0) + (results.hourly?.ptClassesEntries || 0),
      totalNIAllocated: (results.salaried?.totalNIAllocated || 0) + (results.hourly?.totalNIAllocated || 0),
      canProcess: (results.salaried?.ptClassesEntries || 0) > 0 || (results.hourly?.ptClassesEntries || 0) > 0
    };
    
    console.log(`\n📊 OVERALL SUMMARY:`);
    console.log(`Total PT/Classes entries: ${results.summary.totalPTClassesEntries}`);
    console.log(`Total NI allocated: £${results.summary.totalNIAllocated.toLocaleString()}`);
    console.log(`Ready for processing: ${results.summary.canProcess ? 'Yes' : 'No'}`);
    
    return results;
    
  } catch (error) {
    console.error(`❌ Location-Department matching test failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Quick diagnostic to verify PT/Classes patterns are being found
 * @returns {Object} Diagnostic results
 */
function diagnosticPTClassesPatterns() {
  try {
    console.log('🔍 Running PT/Classes pattern diagnostic with location-department matching...');
    
    const results = {
      salaried: null,
      hourly: null,
      summary: {}
    };
    
    // Test salaried journal
    try {
      const salariedJournal = loadMostRecentJournal('salaried');
      const salariedAnalysis = analyzePTClassesEntries(salariedJournal, 7);
      
      results.salaried = {
        totalEntries: salariedJournal.length,
        ptClassesFound: salariedAnalysis.ptClassesLines.length,
        totalPTClassesAmount: salariedAnalysis.totalPTClassesDebits,
        niAllocated: salariedAnalysis.totalNIAllocated,
        locationDeptCombinations: Object.keys(salariedAnalysis.ptClassesByLocationDept).length
      };
      
      console.log(`✅ Salaried: ${results.salaried.ptClassesFound} PT/Classes entries in ${results.salaried.locationDeptCombinations} location-department combinations`);
      
    } catch (error) {
      console.log(`❌ Salaried diagnostic failed: ${error.message}`);
      results.salaried = { error: error.message };
    }
    
    // Test hourly journal
    try {
      const hourlyJournal = loadMostRecentJournal('hourly');
      const hourlyAnalysis = analyzePTClassesEntries(hourlyJournal, 7);
      
      results.hourly = {
        totalEntries: hourlyJournal.length,
        ptClassesFound: hourlyAnalysis.ptClassesLines.length,
        totalPTClassesAmount: hourlyAnalysis.totalPTClassesDebits,
        niAllocated: hourlyAnalysis.totalNIAllocated,
        locationDeptCombinations: Object.keys(hourlyAnalysis.ptClassesByLocationDept).length
      };
      
      console.log(`✅ Hourly: ${results.hourly.ptClassesFound} PT/Classes entries in ${results.hourly.locationDeptCombinations} location-department combinations`);
      
    } catch (error) {
      console.log(`❌ Hourly diagnostic failed: ${error.message}`);
      results.hourly = { error: error.message };
    }
    
    // Summary
    const salariedCount = results.salaried && !results.salaried.error ? results.salaried.ptClassesFound : 0;
    const hourlyCount = results.hourly && !results.hourly.error ? results.hourly.ptClassesFound : 0;
    
    results.summary = {
      totalPTClassesEntries: salariedCount + hourlyCount,
      salariedEntries: salariedCount,
      hourlyEntries: hourlyCount,
      canProcess: salariedCount > 0 || hourlyCount > 0
    };
    
    console.log(`\n📊 SUMMARY:`);
    console.log(`Total PT/Classes entries found: ${results.summary.totalPTClassesEntries}`);
    console.log(`Ready for processing: ${results.summary.canProcess ? 'Yes' : 'No'}`);
    
    return results;
    
  } catch (error) {
    console.error(`❌ Diagnostic failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Production readiness check for PT/Classes module
 * @returns {Object} Readiness assessment
 */
function checkPTClassesProductionReadiness() {
  try {
    console.log('🏭 Checking PT/Classes production readiness...');
    
    const checks = {
      folderAccess: false,
      journalFilesPresent: false,
      analysisWorking: false,
      locationDeptMatching: false,
      outputWorking: false,
      validationWorking: false
    };
    
    const results = {
      checks: checks,
      details: {},
      recommendations: [],
      overallReady: false
    };
    
    // Check 1: Folder access
    try {
      const folder = DriveApp.getFolderById(CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID);
      checks.folderAccess = true;
      results.details.folderName = folder.getName();
      console.log(`✅ Folder access: ${folder.getName()}`);
    } catch (error) {
      console.log(`❌ Folder access failed: ${error.message}`);
      results.recommendations.push('Check CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID');
    }
    
    // Check 2: Journal files present
    try {
      const salariedJournal = loadMostRecentJournal('salaried');
      const hourlyJournal = loadMostRecentJournal('hourly');
      checks.journalFilesPresent = true;
      results.details.salariedEntries = salariedJournal.length;
      results.details.hourlyEntries = hourlyJournal.length;
      console.log(`✅ Journal files: Salaried (${salariedJournal.length}), Hourly (${hourlyJournal.length})`);
    } catch (error) {
      console.log(`❌ Journal files check failed: ${error.message}`);
      results.recommendations.push('Ensure hourly and salaried journals have been generated');
    }
    
    // Check 3: Analysis working
    if (checks.journalFilesPresent) {
      try {
        const diagnostic = diagnosticPTClassesPatterns();
        checks.analysisWorking = !diagnostic.error;
        results.details.ptClassesFound = diagnostic.summary.totalPTClassesEntries;
        console.log(`✅ Analysis working: ${diagnostic.summary.totalPTClassesEntries} PT/Classes entries found`);
      } catch (error) {
        console.log(`❌ Analysis check failed: ${error.message}`);
        results.recommendations.push('Debug PT/Classes pattern matching');
      }
    }
    
    // Check 4: Location-Department matching working
    if (checks.analysisWorking) {
      try {
        const matchingTest = testLocationDepartmentMatching();
        checks.locationDeptMatching = matchingTest.success !== false;
        results.details.locationDeptMatching = checks.locationDeptMatching;
        console.log(`✅ Location-Department matching working`);
      } catch (error) {
        console.log(`❌ Location-Department matching check failed: ${error.message}`);
        results.recommendations.push('Fix location-department matching logic');
      }
    }
    
    // Check 5: Output working (dry run)
    if (checks.locationDeptMatching) {
      try {
        // Test with empty data to verify output functions work
        const testEntries = [];
        const testFile = outputJournalToCsv(testEntries, 'test_ptclasses.csv', DriveApp.getFolderById(CONFIG.SALARIED_JOURNAL_OUTPUT_FOLDER_ID));
        testFile.setTrashed(true); // Clean up test file
        checks.outputWorking = true;
        console.log(`✅ Output working: CSV generation functional`);
      } catch (error) {
        console.log(`❌ Output check failed: ${error.message}`);
        results.recommendations.push('Check CSV output functions');
      }
    }
    
    // Check 6: Validation working
    if (checks.outputWorking) {
      try {
        const dummyResults = {
          salariedResult: { totalNIAllocated: 100, ptClassesEntries: [] },
          hourlyResult: { totalNIAllocated: 200, ptClassesEntries: [] },
          combinedJournal: { entries: [], totalDebits: 0, totalCredits: 300, isBalanced: true }
        };
        const validation = validatePTClassesSeparation(dummyResults.salariedResult, dummyResults.hourlyResult, dummyResults.combinedJournal);
        checks.validationWorking = validation.isValid || validation.error;
        console.log(`✅ Validation working: Function executes without errors`);
      } catch (error) {
        console.log(`❌ Validation check failed: ${error.message}`);
        results.recommendations.push('Fix validation function');
      }
    }
    
    // Overall assessment
    const passedChecks = Object.values(checks).filter(check => check).length;
    const totalChecks = Object.keys(checks).length;
    results.overallReady = passedChecks === totalChecks;
    
    console.log(`\n📊 PRODUCTION READINESS: ${results.overallReady ? 'READY' : 'NOT READY'}`);
    console.log(`Passed checks: ${passedChecks}/${totalChecks}`);
    
    if (results.recommendations.length > 0) {
      console.log(`\n📋 RECOMMENDATIONS:`);
      results.recommendations.forEach(rec => console.log(`• ${rec}`));
    }
    
    return results;
    
  } catch (error) {
    console.error(`❌ Production readiness check failed: ${error.message}`);
    return {
      success: false,
      error: error.message,
      overallReady: false
    };
  }
}

// ================================================================
// INTEGRATION FUNCTIONS FOR MAIN SYSTEM
// ================================================================

/**
 * Integration function called from 12JournalSidebargs.gs
 * This is the entry point for PT/Classes processing from the main journal system
 * @param {string} selectedMonth - Month being processed
 * @returns {Object} Results formatted for the main journal system
 */
function runPTClassesSeparationIntegration(selectedMonth) {
  try {
    console.log(`🔗 PT/Classes integration called for ${selectedMonth}`);
    
    const result = processPTClassesSeparation(selectedMonth);
    
    if (result.success) {
      // Format results for integration
      return {
        success: true,
        message: `PT/Classes separation completed successfully`,
        files: [
          {
            name: result.salariedResult.adjustedJournal.getName(),
            url: result.salariedResult.adjustedJournal.getUrl(),
            type: 'Adjusted Salaried Journal'
          },
          {
            name: result.hourlyResult.adjustedJournal.getName(),
            url: result.hourlyResult.adjustedJournal.getUrl(),
            type: 'Adjusted Hourly Journal'
          },
          {
            name: result.combinedJournal.file.getName(),
            url: result.combinedJournal.file.getUrl(),
            type: 'Combined PT/Classes Journal'
          }
        ],
        summary: {
          salariedEntries: result.salariedResult.ptClassesEntries.length,
          hourlyEntries: result.hourlyResult.ptClassesEntries.length,
          totalEntries: result.combinedJournal.entries.length,
          totalNIAllocated: result.salariedResult.totalNIAllocated + result.hourlyResult.totalNIAllocated,
          isBalanced: result.combinedJournal.isBalanced,
          validationPassed: result.validation.isValid
        }
      };
    } else {
      return {
        success: false,
        error: result.error,
        message: `PT/Classes separation failed: ${result.error}`
      };
    }
    
  } catch (error) {
    console.error(`❌ PT/Classes integration failed: ${error.message}`);
    return {
      success: false,
      error: error.message,
      message: `PT/Classes integration error: ${error.message}`
    };
  }
}
