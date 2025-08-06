/**
 * PRE-PAYROLL AUDIT & EXCEPTION MANAGEMENT
 * Handles exception logging, sheet management, and approval workflow
 * Designed for high-volume data with efficient exception-only storage
 */

// Exception sheet configuration
const EXCEPTION_SHEET_CONFIG = {
  headers: [
    'exceptionId',
    'checkName', 
    'employeeNumber',
    'employeeName',
    'issue',
    'severity',
    'details',
    'reviewStatus',
    'reviewedBy',
    'reviewedAt',
    'createdAt',
    'additionalData'
  ],
  sheetName: 'PayrollExceptions',
  folderConfigKey: 'PAYROLL_PREVIEW_FILE_ID' // Store in payroll preview folder (despite name, this is a folder ID)
};

/**
 * Creates a new exceptions sheet for payroll validation
 * @param {string} sessionId - Unique session identifier (timestamp-based)
 * @returns {Object} Sheet creation result with file info
 */
function createExceptionsSheet(sessionId) {
  try {
    console.log(`📊 Creating exceptions sheet for session: ${sessionId}`);
    
    const sheetName = `${EXCEPTION_SHEET_CONFIG.sheetName}_${sessionId}`;
    const spreadsheet = SpreadsheetApp.create(sheetName);
    const sheet = spreadsheet.getActiveSheet();
    
    // Set up headers
    const headers = EXCEPTION_SHEET_CONFIG.headers;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('white');
    
    // Set column widths for readability
    sheet.setColumnWidth(1, 150); // exceptionId
    sheet.setColumnWidth(2, 150); // checkName
    sheet.setColumnWidth(3, 100); // employeeNumber
    sheet.setColumnWidth(4, 150); // employeeName
    sheet.setColumnWidth(5, 300); // issue
    sheet.setColumnWidth(6, 80);  // severity
    sheet.setColumnWidth(7, 250); // details
    sheet.setColumnWidth(8, 100); // reviewStatus
    sheet.setColumnWidth(9, 120); // reviewedBy
    sheet.setColumnWidth(10, 140); // reviewedAt
    sheet.setColumnWidth(11, 140); // createdAt
    sheet.setColumnWidth(12, 200); // additionalData
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Move to configured folder
    const file = DriveApp.getFileById(spreadsheet.getId());
    const targetFolder = DriveApp.getFolderById(CONFIG[EXCEPTION_SHEET_CONFIG.folderConfigKey]);
    targetFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    
    console.log(`✅ Exceptions sheet created: ${sheetName}`);
    
    return {
      success: true,
      sheetId: spreadsheet.getId(),
      sheetName: sheetName,
      url: spreadsheet.getUrl(),
      headerCount: headers.length
    };
    
  } catch (error) {
    console.error(`❌ Failed to create exceptions sheet: ${error.message}`);
    throw new Error(`Exception sheet creation failed: ${error.message}`);
  }
}

/**
 * Writes exceptions to the sheet in batch for performance
 * @param {string} sheetId - Google Sheets file ID
 * @param {Array<Object>} exceptions - Array of exception objects
 * @returns {Object} Write result with count and performance info
 */
function writeExceptionsToSheet(sheetId, exceptions) {
  const startTime = Date.now();
  
  try {
    console.log(`📝 Writing ${exceptions.length} exceptions to sheet...`);
    
    if (!exceptions || exceptions.length === 0) {
      return {
        success: true,
        exceptionsWritten: 0,
        message: "No exceptions to write"
      };
    }
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getActiveSheet();
    const headers = EXCEPTION_SHEET_CONFIG.headers;
    
    // Prepare data rows
    const dataRows = exceptions.map(exception => {
      return headers.map(header => {
        switch(header) {
          case 'exceptionId':
            return exception.exceptionId || generateExceptionId(exception);
          case 'checkName':
            return exception.checkName || 'Unknown Check';
          case 'employeeNumber':
            return exception.employeeNumber || 'Unknown';
          case 'employeeName':
            return exception.employeeName || 'Name Unknown';
          case 'issue':
            return exception.issue || 'No issue description';
          case 'severity':
            return exception.severity || 'warning';
          case 'details':
            return typeof exception.details === 'object' ? 
              JSON.stringify(exception.details) : 
              (exception.details || '');
          case 'reviewStatus':
            return exception.reviewStatus || null;
          case 'reviewedBy':
            return exception.reviewedBy || null;
          case 'reviewedAt':
            return exception.reviewedAt || null;
          case 'createdAt':
            return exception.createdAt || new Date().toISOString();
          case 'additionalData':
            return exception.additionalData ? JSON.stringify(exception.additionalData) : '';
          default:
            return exception[header] || '';
        }
      });
    });
    
    // Write data in batch (much faster than row-by-row)
    const lastRow = sheet.getLastRow();
    const targetRange = sheet.getRange(lastRow + 1, 1, dataRows.length, headers.length);
    targetRange.setValues(dataRows);
    
    // Apply conditional formatting for severity
    applySeverityFormatting(sheet, lastRow + 1, dataRows.length);
    
    const executionTime = Date.now() - startTime;
    console.log(`✅ ${exceptions.length} exceptions written in ${executionTime}ms`);
    
    return {
      success: true,
      exceptionsWritten: exceptions.length,
      executionTime: executionTime,
      totalRows: lastRow + dataRows.length,
      sheetUrl: spreadsheet.getUrl()
    };
    
  } catch (error) {
    console.error(`❌ Failed to write exceptions: ${error.message}`);
    throw new Error(`Exception writing failed: ${error.message}`);
  }
}

/**
 * Reads exceptions from sheet with optional filtering
 * @param {string} sheetId - Google Sheets file ID
 * @param {Object} filters - Optional filters {checkName, severity, reviewStatus}
 * @returns {Array<Object>} Array of exception objects
 */
function readExceptionsFromSheet(sheetId, filters = {}) {
  const startTime = Date.now();
  
  try {
    console.log(`📖 Reading exceptions from sheet with filters:`, filters);
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return []; // No data except headers
    }
    
    const [headers, ...rows] = data;
    const headerIndex = {};
    headers.forEach((header, index) => {
      headerIndex[header] = index;
    });
    
    // Convert rows to objects
    const exceptions = rows.map(row => {
      const exception = {};
      headers.forEach((header, index) => {
        let value = row[index];
        
        // Parse JSON fields
        if (['details', 'additionalData'].includes(header) && typeof value === 'string' && value.startsWith('{')) {
          try {
            value = JSON.parse(value);
          } catch (e) {
            // Keep as string if JSON parsing fails
          }
        }
        
        exception[header] = value;
      });
      return exception;
    });
    
    // Apply filters
    let filteredExceptions = exceptions;
    
    if (filters.checkName) {
      filteredExceptions = filteredExceptions.filter(e => e.checkName === filters.checkName);
    }
    
    if (filters.severity) {
      filteredExceptions = filteredExceptions.filter(e => e.severity === filters.severity);
    }
    
    if (filters.reviewStatus !== undefined) {
      if (filters.reviewStatus === null) {
        filteredExceptions = filteredExceptions.filter(e => !e.reviewStatus);
      } else {
        filteredExceptions = filteredExceptions.filter(e => e.reviewStatus === filters.reviewStatus);
      }
    }
    
    const executionTime = Date.now() - startTime;
    console.log(`✅ Read ${filteredExceptions.length}/${exceptions.length} exceptions in ${executionTime}ms`);
    
    return filteredExceptions;
    
  } catch (error) {
    console.error(`❌ Failed to read exceptions: ${error.message}`);
    throw new Error(`Exception reading failed: ${error.message}`);
  }
}

/**
 * Updates review status for a specific exception
 * @param {string} sheetId - Google Sheets file ID
 * @param {string} exceptionId - Unique exception identifier
 * @param {string} status - Review status ('ok', 'amended', or null)
 * @param {string} reviewedBy - Email of reviewer
 * @returns {Object} Update result
 */
function updateExceptionReviewStatus(sheetId, exceptionId, status, reviewedBy) {
  try {
    console.log(`📝 Updating exception ${exceptionId} to status: ${status}`);
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const [headers, ...rows] = data;
    
    // Find header indices
    const exceptionIdCol = headers.indexOf('exceptionId') + 1;
    const reviewStatusCol = headers.indexOf('reviewStatus') + 1;
    const reviewedByCol = headers.indexOf('reviewedBy') + 1;
    const reviewedAtCol = headers.indexOf('reviewedAt') + 1;
    
    if (exceptionIdCol === 0) {
      throw new Error('exceptionId column not found in sheet');
    }
    
    // Find the exception row
    let rowFound = false;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][exceptionIdCol - 1] === exceptionId) {
        const rowNum = i + 2; // +1 for header, +1 for 0-based index
        
        // Update the status
        if (reviewStatusCol > 0) {
          sheet.getRange(rowNum, reviewStatusCol).setValue(status);
        }
        if (reviewedByCol > 0) {
          sheet.getRange(rowNum, reviewedByCol).setValue(reviewedBy);
        }
        if (reviewedAtCol > 0) {
          sheet.getRange(rowNum, reviewedAtCol).setValue(new Date());
        }
        
        rowFound = true;
        break;
      }
    }
    
    if (!rowFound) {
      throw new Error(`Exception with ID ${exceptionId} not found in sheet`);
    }
    
    console.log(`✅ Exception ${exceptionId} updated successfully`);
    
    return {
      success: true,
      exceptionId: exceptionId,
      status: status,
      reviewedBy: reviewedBy,
      updatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    console.error(`❌ Failed to update exception status: ${error.message}`);
    throw new Error(`Exception update failed: ${error.message}`);
  }
}

/**
 * Gets summary statistics for exceptions by check
 * @param {string} sheetId - Google Sheets file ID
 * @returns {Object} Summary statistics
 */
function getExceptionsSummary(sheetId) {
  try {
    console.log(`📊 Generating exceptions summary for sheet: ${sheetId}`);
    
    const exceptions = readExceptionsFromSheet(sheetId);
    
    const summary = {
      totalExceptions: exceptions.length,
      byCheck: {},
      bySeverity: { critical: 0, warning: 0, info: 0 },
      byStatus: { pending: 0, ok: 0, amended: 0 },
      overallProgress: {
        reviewed: 0,
        pending: 0,
        percentComplete: 0
      }
    };
    
    exceptions.forEach(exception => {
      const checkName = exception.checkName;
      const severity = exception.severity || 'warning';
      const status = exception.reviewStatus;
      
      // By check
      if (!summary.byCheck[checkName]) {
        summary.byCheck[checkName] = {
          total: 0,
          pending: 0,
          reviewed: 0,
          critical: 0,
          warning: 0
        };
      }
      summary.byCheck[checkName].total++;
      
      if (status === 'ok' || status === 'amended') {
        summary.byCheck[checkName].reviewed++;
        summary.overallProgress.reviewed++;
      } else {
        summary.byCheck[checkName].pending++;
        summary.overallProgress.pending++;
      }
      
      if (severity === 'critical') {
        summary.byCheck[checkName].critical++;
      } else {
        summary.byCheck[checkName].warning++;
      }
      
      // By severity
      summary.bySeverity[severity] = (summary.bySeverity[severity] || 0) + 1;
      
      // By status
      if (status === 'ok') {
        summary.byStatus.ok++;
      } else if (status === 'amended') {
        summary.byStatus.amended++;
      } else {
        summary.byStatus.pending++;
      }
    });
    
    // Calculate overall progress percentage
    if (summary.totalExceptions > 0) {
      summary.overallProgress.percentComplete = Math.round(
        (summary.overallProgress.reviewed / summary.totalExceptions) * 100
      );
    }
    
    console.log(`✅ Summary generated: ${summary.totalExceptions} total exceptions, ${summary.overallProgress.percentComplete}% reviewed`);
    
    return summary;
    
  } catch (error) {
    console.error(`❌ Failed to generate summary: ${error.message}`);
    throw new Error(`Summary generation failed: ${error.message}`);
  }
}

/**
 * Creates a unique exception ID
 * @param {Object} exception - Exception object
 * @returns {string} Unique exception ID
 */
function generateExceptionId(exception) {
  const checkPrefix = (exception.checkName || 'unknown').toLowerCase().replace(/[^a-z0-9]/g, '');
  const empNumber = exception.employeeNumber || 'unknown';
  const timestamp = Date.now();
  const random = Math.floor(Math.random() * 1000);
  
  return `${checkPrefix}_${empNumber}_${timestamp}_${random}`;
}

/**
 * Applies conditional formatting for severity levels
 * @param {Sheet} sheet - Google Sheets object
 * @param {number} startRow - Starting row for formatting
 * @param {number} numRows - Number of rows to format
 */
function applySeverityFormatting(sheet, startRow, numRows) {
  try {
    const severityCol = EXCEPTION_SHEET_CONFIG.headers.indexOf('severity') + 1;
    if (severityCol === 0) return; // Column not found
    
    const range = sheet.getRange(startRow, severityCol, numRows, 1);
    
    // Clear existing formatting
    range.clearFormat();
    
    // Create conditional formatting rules
    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('critical')
        .setBackground('#fce8e6')
        .setFontColor('#d93025')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('warning')
        .setBackground('#fef7e0')
        .setFontColor('#f29900')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('info')
        .setBackground('#e8f0fe')
        .setFontColor('#1a73e8')
        .setRanges([range])
        .build()
    ];
    
    sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(rules));
    
  } catch (error) {
    console.log(`⚠️ Could not apply formatting: ${error.message}`);
    // Non-critical error - don't throw
  }
}

/**
 * Archives completed exceptions sheet
 * @param {string} sheetId - Google Sheets file ID to archive
 * @returns {Object} Archive result
 */
function archiveExceptionsSheet(sheetId) {
  try {
    console.log(`📦 Archiving exceptions sheet: ${sheetId}`);
    
    const file = DriveApp.getFileById(sheetId);
    const originalName = file.getName();
    const archivedName = `ARCHIVED_${originalName}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm')}`;
    
    // Rename file to indicate it's archived
    file.setName(archivedName);
    
    // Could move to archive folder if one exists
    // const archiveFolder = DriveApp.getFolderById(CONFIG.ARCHIVE_FOLDER_ID);
    // archiveFolder.addFile(file);
    // currentFolder.removeFile(file);
    
    console.log(`✅ Sheet archived as: ${archivedName}`);
    
    return {
      success: true,
      originalName: originalName,
      archivedName: archivedName,
      archivedAt: new Date().toISOString()
    };
    
  } catch (error) {
    console.error(`❌ Failed to archive sheet: ${error.message}`);
    throw new Error(`Archiving failed: ${error.message}`);
  }
}


/**
 * INTEGRATION HELPER: Create standardized exception object
 * Use this in your validation checks to ensure consistent data structure
 */
function createStandardException(checkName, employeeNumber, employeeName, issue, severity, details = {}) {
  return {
    checkName: checkName,
    employeeNumber: employeeNumber || 'Unknown',
    employeeName: employeeName || 'Name Unknown',
    issue: issue,
    severity: severity || 'warning', // critical | warning | info
    details: details,
    reviewStatus: null,
    reviewedBy: null,
    reviewedAt: null,
    createdAt: new Date().toISOString(),
    exceptionId: null // Will be generated when written to sheet
  };
}
