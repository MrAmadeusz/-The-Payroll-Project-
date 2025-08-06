/**
 * 08_startersReport.gs
 * Complete production-ready Starters Report with JSON audit trail and late additions logic
 */

/**
 * SCALABLE Google Sheets audit trail - handles thousands/millions of records
 * Each record is a row, not JSON nonsense
 */

/**
 * Gets or creates the audit trail sheet (proper database structure)
 * @returns {Sheet} The audit trail sheet
 */
function getAuditTrailSheet() {
  const folder = DriveApp.getFolderById(CONFIG.STARTERSLEAVERS_LOG_FOLDER_ID);
  const fileName = 'StartersLeaversAudit';
  
  // Try to find existing sheet
  const files = folder.getFilesByName(fileName);
  let spreadsheet;
  
  if (files.hasNext()) {
    const file = files.next();
    spreadsheet = SpreadsheetApp.openById(file.getId());
  } else {
    // Create new sheet
    spreadsheet = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(spreadsheet.getId());
    
    // Move to correct folder
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }
  
  const sheet = spreadsheet.getActiveSheet();
  
  // Set up headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    const headers = ['Report_Date', 'Employee_No', 'Employee_Date', 'Report_Type', 'Date_Range_Key', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Add data validation and formatting for better performance
    const lastColumn = sheet.getLastColumn();
    const dataRange = sheet.getRange(2, 1, 1000, lastColumn); // Format first 1000 rows
    dataRange.setNumberFormat('@'); // Text format for better performance
  }
  
  return sheet;
}

/**
 * Loads starters audit trail (optimized for thousands of records)
 * @returns {Array} Array of starter audit records
 */
function loadStartersAuditTrail() {
  try {
    const sheet = getAuditTrailSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return []; // Only headers, no data
    }
    
    // Use batch operations for better performance
    const batchSize = 1000;
    let allRecords = [];
    
    for (let startRow = 2; startRow <= lastRow; startRow += batchSize) {
      const endRow = Math.min(startRow + batchSize - 1, lastRow);
      const rowCount = endRow - startRow + 1;
      
      const data = sheet.getRange(startRow, 1, rowCount, 6).getValues();
      
      // Filter for starters only and convert to objects
      const batchRecords = data
        .filter(row => row[3] === 'STARTER' && row[1]) // Filter by Report_Type and ensure Employee_No exists
        .map(row => ({
          reportDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          employeeNo: String(row[1]),
          startDate: String(row[2]),
          reportType: 'STARTER',
          dateRangeKey: String(row[4]),
          timestamp: row[5] ? row[5].toISOString() : ''
        }));
      
      allRecords = allRecords.concat(batchRecords);
    }
    
    return allRecords;
    
  } catch (error) {
    console.error(`Error loading starters audit trail: ${error.message}`);
    return [];
  }
}

/**
 * Updates the starters audit trail (optimized batch insert)
 * @param {Array} newStarters - Array of employee objects to add
 * @param {Object} dateRange - Date range object
 */
function updateStartersAuditTrail(newStarters, dateRange) {
  try {
    if (newStarters.length === 0) return;
    
    const sheet = getAuditTrailSheet();
    
    const reportDate = new Date();
    const dateRangeKey = `${Utilities.formatDate(dateRange.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}_${Utilities.formatDate(dateRange.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
    
    // Prepare new rows in batch
    const newRows = newStarters.map(emp => [
      reportDate,
      emp.EmployeeNumber,
      emp.ContinuousServiceStartDate,
      'STARTER',
      dateRangeKey,
      new Date()
    ]);
    
    // Single batch insert for better performance
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, newRows.length, 6).setValues(newRows);
    
  } catch (error) {
    console.error(`Failed to update starters audit trail: ${error.message}`);
    throw error;
  }
}

/**
 * Loads leavers audit trail (optimized for thousands of records)
 * @returns {Array} Array of leaver audit records
 */
function loadLeaversAuditTrail() {
  try {
    const sheet = getAuditTrailSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return []; // Only headers, no data
    }
    
    // Use batch operations for better performance
    const batchSize = 1000;
    let allRecords = [];
    
    for (let startRow = 2; startRow <= lastRow; startRow += batchSize) {
      const endRow = Math.min(startRow + batchSize - 1, lastRow);
      const rowCount = endRow - startRow + 1;
      
      const data = sheet.getRange(startRow, 1, rowCount, 6).getValues();
      
      // Filter for leavers only
      const batchRecords = data
        .filter(row => row[3] === 'LEAVER' && row[1]) // Filter by Report_Type and ensure Employee_No exists
        .map(row => ({
          reportDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          employeeNo: String(row[1]),
          endDate: String(row[2]),
          reportType: 'LEAVER',
          dateRangeKey: String(row[4]),
          timestamp: row[5] ? row[5].toISOString() : ''
        }));
      
      allRecords = allRecords.concat(batchRecords);
    }
    
    return allRecords;
    
  } catch (error) {
    console.error(`Error loading leavers audit trail: ${error.message}`);
    return [];
  }
}

/**
 * Updates the leavers audit trail (optimized batch insert)
 * @param {Array} newLeavers - Array of employee objects to add
 * @param {Object} dateRange - Date range object
 */
function updateLeaversAuditTrail(newLeavers, dateRange) {
  try {
    if (newLeavers.length === 0) return;
    
    const sheet = getAuditTrailSheet();
    
    const reportDate = new Date();
    const dateRangeKey = `${Utilities.formatDate(dateRange.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}_${Utilities.formatDate(dateRange.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
    
    // Prepare new rows in batch
    const newRows = newLeavers.map(emp => [
      reportDate,
      emp.EmployeeNumber,
      emp.EndDate || emp.LeavingDate || emp.TerminationDate,
      'LEAVER',
      dateRangeKey,
      new Date()
    ]);
    
    // Single batch insert for better performance
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, newRows.length, 6).setValues(newRows);
    
  } catch (error) {
    console.error(`Failed to update leavers audit trail: ${error.message}`);
    throw error;
  }
}

/**
 * Checks if an employee has already been reported for a DIFFERENT date range
 * Uses optimized search for large datasets
 */
function isEmployeeAlreadyReportedElsewhere(employeeNo, auditTrail, currentDateRange) {
  const currentDateRangeKey = `${Utilities.formatDate(currentDateRange.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}_${Utilities.formatDate(currentDateRange.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
  
  return auditTrail.some(record => 
    record.employeeNo === employeeNo && 
    record.reportType === 'STARTER' &&
    record.dateRangeKey !== currentDateRangeKey
  );
}

/**
 * Cleans up old audit records by deleting rows (maintains performance)
 * @param {number} daysToKeep - Number of days to retain records (default: 365)
 * @returns {Object} Cleanup summary
 */
function cleanupOldAuditRecords(daysToKeep = 365) {
  try {
    const sheet = getAuditTrailSheet();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { recordsRemoved: 0, endingRecords: 0 };
    
    // Process in batches to avoid timeout
    const batchSize = 500;
    let totalRemoved = 0;
    
    for (let startRow = lastRow; startRow >= 2; startRow -= batchSize) {
      const endRow = Math.max(startRow - batchSize + 1, 2);
      const rowCount = startRow - endRow + 1;
      
      const data = sheet.getRange(endRow, 1, rowCount, 6).getValues();
      
      // Find rows to delete (working backwards)
      const rowsToDelete = [];
      for (let i = data.length - 1; i >= 0; i--) {
        const recordDate = new Date(data[i][5] || data[i][0]); // Use timestamp or report date
        if (recordDate < cutoffDate) {
          rowsToDelete.push(endRow + i);
        }
      }
      
      // Delete rows (from bottom to top to avoid index shifting)
      rowsToDelete.forEach(rowIndex => {
        sheet.deleteRow(rowIndex);
        totalRemoved++;
      });
    }
    
    return {
      recordsRemoved: totalRemoved,
      endingRecords: sheet.getLastRow() - 1
    };
    
  } catch (error) {
    console.error(`Error during cleanup: ${error.message}`);
    throw error;
  }
}

/**
 * Gets audit trail statistics (optimized for large datasets)
 */
function getAuditTrailStats() {
  try {
    const sheet = getAuditTrailSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { starterCount: 0, leaverCount: 0, totalCount: 0, lastUpdated: 'Never' };
    }
    
    // Use formula-based counting for better performance on large datasets
    const starterCountFormula = `=COUNTIF(D2:D${lastRow},"STARTER")`;
    const leaverCountFormula = `=COUNTIF(D2:D${lastRow},"LEAVER")`;
    
    // Temporarily use a cell to calculate counts
    sheet.getRange('H1').setFormula(starterCountFormula);
    sheet.getRange('H2').setFormula(leaverCountFormula);
    
    const starterCount = sheet.getRange('H1').getValue();
    const leaverCount = sheet.getRange('H2').getValue();
    
    // Clean up temporary formulas
    sheet.getRange('H1:H2').clear();
    
    // Get most recent timestamp from last few rows (more efficient)
    const recentRows = Math.min(10, lastRow - 1);
    const recentData = sheet.getRange(lastRow - recentRows + 1, 6, recentRows, 1).getValues();
    const timestamps = recentData.map(row => new Date(row[0])).filter(date => !isNaN(date));
    const lastUpdated = timestamps.length > 0 ? new Date(Math.max(...timestamps)).toLocaleString() : 'Never';
    
    return {
      starterCount: starterCount,
      leaverCount: leaverCount,
      totalCount: starterCount + leaverCount,
      lastUpdated: lastUpdated
    };
    
  } catch (error) {
    console.error(`Error getting audit trail stats: ${error.message}`);
    return { starterCount: 0, leaverCount: 0, totalCount: 0, lastUpdated: 'Error' };
  }
}

// ===== STARTERS REPORT GENERATION WITH LATE ADDITIONS =====

/**
 * Hybrid generateStartersReport function - preserves original logic + adds late additions
 * @param {Array} data - Employee data array
 * @param {Object} options - Options object (can contain dateRange, lookbackDays)
 * @returns {Array} 2D array with report data
 */
function generateStartersReport(data, options = {}) {
  try {
    const dateRange = options.dateRange || promptForDateRange("Select Date Range for Starters Report");
    const auditTrail = loadStartersAuditTrail();
    
    // PART 1: Original logic - people who started in date range
    const normalStarters = data.filter(emp => {
      const startDate = emp.ContinuousServiceStartDate?.trim();
      return startDate && isDateInRange(startDate, dateRange);
    }).filter(emp => 
      !isEmployeeAlreadyReportedElsewhere(emp.EmployeeNumber, auditTrail, dateRange)
    );
    
    // PART 2: Late additions - people who started before range but never reported
    const lookbackDays = options.lookbackDays || 60;
    const lookbackDate = new Date(dateRange.startDate);
    lookbackDate.setDate(lookbackDate.getDate() - lookbackDays);
    
    const lateAdditions = data.filter(emp => {
      const startDate = emp.ContinuousServiceStartDate?.trim();
      if (!startDate) return false;
      
      const parsedStartDate = parseDMY(startDate);
      if (!parsedStartDate) return false;
      
      // Must have started before report range but after lookback date
      const startedBeforeRange = parsedStartDate < dateRange.startDate;
      const withinLookback = parsedStartDate >= lookbackDate;
      
      // Must never have been reported (not in audit trail at all)
      const neverReported = !auditTrail.some(record => record.employeeNo === emp.EmployeeNumber);
      
      return startedBeforeRange && withinLookback && neverReported;
    });
    
    // Combine both groups
    const allNewStarters = [...normalStarters, ...lateAdditions];
    
    // Remove duplicates (in case someone appears in both lists)
    const uniqueStarters = allNewStarters.filter((emp, index, arr) => 
      arr.findIndex(e => e.EmployeeNumber === emp.EmployeeNumber) === index
    );
    
    // Sort by start date
    uniqueStarters.sort((a, b) => {
      const dateA = parseDMY(a.ContinuousServiceStartDate);
      const dateB = parseDMY(b.ContinuousServiceStartDate);
      return dateA - dateB;
    });
    
    // Build report
    const header = [
      'Employee No', 'Name', 'Email Address', 'Hotel (Unit)', 'Department', 
      'Contracted Hours', 'Start Date', 'Date of Birth', 'Phone Number', 'Job Title'
    ];
    
    const rows = uniqueStarters.map(emp => [
      emp.EmployeeNumber || '',
      `${emp.Firstnames || ''} ${emp.Surname || ''}`.trim(),
      emp.EmailAddress || emp.Email || '',
      emp.Location || emp.Hotel || emp.Unit || '',
      emp.Department || emp.Division || '',
      emp.ContractedHours || emp.Hours || '',
      emp.ContinuousServiceStartDate || '',
      emp.DateOfBirth || emp.DOB || '',
      emp.MobileTel || emp.PhoneNumber || emp.Phone || emp.MobileNumber || '',
      emp.JobTitle || emp.Role || emp.Position || ''
    ]);
    
    // Update audit trail with new starters
    if (uniqueStarters.length > 0) {
      updateStartersAuditTrail(uniqueStarters, dateRange);
    }
    
    // Build summary with breakdown
    const rangeDisplay = formatDateRangeForDisplay(dateRange);
    const titleRow = ['WEEKLY STARTERS REPORT', rangeDisplay, '', '', '', '', '', '', '', ''];
    
    let summaryText = `${uniqueStarters.length} new starters to report`;
    if (lateAdditions.length > 0) {
      summaryText += ` (${normalStarters.length} in period, ${lateAdditions.length} late additions)`;
    }
    
    const summaryRow = ['SUMMARY', summaryText, '', '', '', '', '', '', '', ''];
    const blankRow = ['', '', '', '', '', '', '', '', '', ''];
    
    return [titleRow, summaryRow, blankRow, header, ...rows];
    
  } catch (error) {
    console.error(`Error generating starters report: ${error.message}`);
    throw new Error('Failed to generate starters report: ' + error.message);
  }
}

// ===== UTILITY FUNCTIONS FOR TESTING/MAINTENANCE =====

/**
 * Clears audit trail for testing purposes
 */
function clearAuditTrailForTesting() {
  const initialData = {
    starters: [],
    leavers: [],
    lastCleaned: new Date().toISOString(),
    version: "1.0"
  };
  saveAuditTrail(initialData);
}

/**
 * Backs up current audit trail
 * @returns {Object} Current audit trail data
 */
function backupAuditTrail() {
  return loadAuditTrail();
}

/**
 * Restores audit trail from backup
 * @param {Object} backupData - Backup data to restore
 */
function restoreAuditTrail(backupData) {
  if (backupData) {
    saveAuditTrail(backupData);
  }
}

// ===== ADMIN FUNCTIONS =====

/**
 * Manual cleanup function for admins
 * Removes audit records older than specified days
 * @param {number} daysToKeep - Days to retain (default: 365)
 */
function adminCleanupAuditTrail(daysToKeep = 365) {
  try {
    const ui = SpreadsheetApp.getUi();
    const confirmation = ui.alert(
      'Cleanup Audit Trail',
      `This will remove audit records older than ${daysToKeep} days. Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmation === ui.Button.YES) {
      const result = cleanupOldAuditRecords(daysToKeep);
      ui.alert(
        'Cleanup Complete',
        `Removed ${result.recordsRemoved} old records.\nRecords remaining: ${result.endingRecords}`,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Cleanup Failed', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows audit trail statistics for admins
 */
function showAuditTrailStats() {
  try {
    const auditData = loadAuditTrail();
    const starterCount = auditData.starters?.length || 0;
    const leaverCount = auditData.leavers?.length || 0;
    const lastUpdated = auditData.lastUpdated ? new Date(auditData.lastUpdated).toLocaleString() : 'Never';
    
    const stats = `📊 AUDIT TRAIL STATISTICS
═══════════════════════════

📈 Starter Records: ${starterCount}
📉 Leaver Records: ${leaverCount}
📋 Total Records: ${starterCount + leaverCount}

🕐 Last Updated: ${lastUpdated}
📅 Version: ${auditData.version || '1.0'}
🗂️ Storage: JSON file using existing utilities

💡 Use 'Admin Cleanup Audit Trail' to remove old records if needed.`;
    
    SpreadsheetApp.getUi().alert('Audit Trail Statistics', stats, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Stats Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
