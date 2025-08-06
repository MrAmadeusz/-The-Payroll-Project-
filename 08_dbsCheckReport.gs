/**
 * 08_dbsCheckReport.gs
 * DBS Check Report: Detects new starters and transfers into GM/Nights/Leisure
 * FIXED: Consistent audit checking logic to prevent duplicate reporting
 */

/**
 * Entry point for DBS Check Report
 * Mirrors report scaffolding used across the reporting engine
 * @param {Array} data - Full employee data array from getAllEmployees()
 * @param {Object} options - Optional parameters (e.g. dateRange)
 * @returns {Array} 2D array suitable for export
 */
function generateDBSCheckReport(data, options = {}) {
  try {
    const dateRange = options.dateRange || promptForDateRange("Select Date Range for DBS Check Report");
    const lookbackDays = options.lookbackDays || 60;
    const auditTrail = loadAuditTrailByType('DBSCHECK');

    const currentEmployees = data.filter(e => e.Status === 'Current');

    // --- PART 1A: Starters in DBS Departments ---
    const inPeriodStarters = currentEmployees.filter(emp => {
      const start = emp.ContinuousServiceStartDate?.trim();
      return (
        start &&
        isDateInRange(start, dateRange) &&
        isDBSRelevant(emp)
      );
    }).filter(emp => !isEmployeeAlreadyReportedForDBS(emp.EmployeeNumber, auditTrail, dateRange));

    // --- PART 1B: Late Starters in DBS Departments ---
    const lookbackDate = new Date(dateRange.startDate);
    lookbackDate.setDate(lookbackDate.getDate() - lookbackDays);

    const lateStarters = currentEmployees.filter(emp => {
      const start = emp.ContinuousServiceStartDate?.trim();
      if (!start || !isDBSRelevant(emp)) return false;
      const parsed = parseDMY(start);
      if (!parsed) return false;
      const beforeRange = parsed < dateRange.startDate;
      const afterLookback = parsed >= lookbackDate;
      const neverReported = !isEmployeeAlreadyReportedForDBS(emp.EmployeeNumber, auditTrail, dateRange);
      return beforeRange && afterLookback && neverReported;
    });

    const starters = [...inPeriodStarters, ...lateStarters];

    // --- PART 2: Transfers into DBS Departments ---
    const snapshot = loadLatestEmployeeSnapshot();
    const snapshotMap = Object.fromEntries(snapshot.map(e => [e.EmployeeNumber, e]));

    const transfers = currentEmployees.filter(emp => {
      const previous = snapshotMap[emp.EmployeeNumber];
      if (!previous || !isDBSRelevant(emp)) return false;

      const currentDivision = (emp.Division || '').toLowerCase();
      const currentLocation = (emp.Location || '').toLowerCase();
      const previousDivision = (previous.Division || '').toLowerCase();
      const previousLocation = (previous.Location || '').toLowerCase();

      const wasNotDBS = !(
        (previousDivision === 'gm') ||
        (previousDivision.startsWith('nights')) ||
        (previousDivision.startsWith('leisure') && previousLocation !== 'support centre')
      );

      const nowIsDBS = isDBSRelevant(emp);
      const notAlreadyReported = !isEmployeeAlreadyReportedForDBS(emp.EmployeeNumber, auditTrail, dateRange);

      return wasNotDBS && nowIsDBS && notAlreadyReported;
    });

    // --- Combine and deduplicate ---
    const allFlagged = [...starters, ...transfers];
    const unique = allFlagged.filter((e, i, arr) => arr.findIndex(x => x.EmployeeNumber === e.EmployeeNumber) === i);

    // --- Sort by start date ---
    unique.sort((a, b) => parseDMY(a.ContinuousServiceStartDate) - parseDMY(b.ContinuousServiceStartDate));

    // --- Build report rows ---
    const header = ['Employee No', 'Name', 'Hotel', 'Email Address', 'Start Date', 'Job Title', 'Date of Birth', 'Phone Number'];
    const rows = unique.map(emp => [
      emp.EmployeeNumber || '',
      `${emp.Firstnames || ''} ${emp.Surname || ''}`.trim(),
      emp.Location || emp.Hotel || '',
      emp.EmailAddress || emp.Email || '',
      emp.ContinuousServiceStartDate || '',
      emp.JobTitle || emp.Role || '',
      emp.DateOfBirth || emp.DOB || '',
      emp.MobileTel || emp.PhoneNumber || emp.Phone || emp.MobileNumber || ''
    ]);

    // --- Debug deduplication ---
    console.log("=== DEDUPLICATION DEBUG ===");
    console.log("Starters found:", starters.length);
    console.log("Transfers found:", transfers.length);
    console.log("Total flagged before dedup:", allFlagged.length);
    console.log("Unique employees after dedup:", unique.length);

    // Check for specific employee
    const targetEmployee = '30014743';
    const startersMatch = starters.filter(e => e.EmployeeNumber === targetEmployee);
    const transfersMatch = transfers.filter(e => e.EmployeeNumber === targetEmployee);
    const allMatch = allFlagged.filter(e => e.EmployeeNumber === targetEmployee);
    const uniqueMatch = unique.filter(e => e.EmployeeNumber === targetEmployee);

    console.log(`Employee ${targetEmployee}:`);
    console.log(`- In starters: ${startersMatch.length}`);
    console.log(`- In transfers: ${transfersMatch.length}`);
    console.log(`- In allFlagged: ${allMatch.length}`);
    console.log(`- In unique: ${uniqueMatch.length}`);

    // --- Audit ---
    if (unique.length > 0) updateAuditTrail(unique, dateRange, 'DBSCHECK');

    // --- Build output ---
    const title = ['WEEKLY DBS CHECK REPORT', formatDateRangeForDisplay(dateRange), '', '', '', '', '', ''];
    const summaryText = `${unique.length} employee(s) flagged for DBS check` +
      (transfers.length > 0
        ? ` (${inPeriodStarters.length} in period, ${lateStarters.length} late starters, ${transfers.length} transfers)`
        : ` (${inPeriodStarters.length} in period, ${lateStarters.length} late starters)`);
    const summary = ['SUMMARY', summaryText, '', '', '', '', '', ''];
    const blank = ['', '', '', '', '', '', '', ''];

    return [title, summary, blank, header, ...rows];

  } catch (err) {
    console.error("DBS Check report error:", err);
    throw new Error("Failed to generate DBS Check report: " + err.message);
  }
}

/**
 * Checks if an employee has already been reported for DBS in ANY previous date range
 * This ensures consistent audit checking logic across all parts of the DBS report
 * @param {string} employeeNo - Employee number to check
 * @param {Array} auditTrail - Array of DBSCHECK audit records
 * @param {Object} currentDateRange - Current report date range
 * @returns {boolean} True if employee has been reported for DBS before (excluding current date range)
 */
function isEmployeeAlreadyReportedForDBS(employeeNo, auditTrail, currentDateRange) {
  const currentDateRangeKey = `${Utilities.formatDate(currentDateRange.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}_${Utilities.formatDate(currentDateRange.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
  
  return auditTrail.some(record => 
    record.employeeNo === employeeNo && 
    record.reportType === 'DBSCHECK' &&
    record.dateRangeKey !== currentDateRangeKey
  );
}

/**
 * Determines if an employee requires DBS check based on division and location
 * @param {Object} emp - Employee object
 * @returns {boolean} True if employee needs DBS check
 */
function isDBSRelevant(emp) {
  const division = (emp.Division || '').toLowerCase();
  const location = (emp.Location || '').toLowerCase();
  return (
    (division === 'gm') ||
    (division.startsWith('nights')) ||
    (division.startsWith('leisure') && location !== 'support centre')
  );
}

/**
 * Loads audit trail records filtered by report type
 * @param {string} type - Report type to filter by (e.g., 'DBSCHECK')
 * @returns {Array} Filtered audit records
 */
function loadAuditTrailByType(type) {
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
      
      // Filter for the specified report type
      const batchRecords = data
        .filter(row => row[3] === type && row[1]) // Filter by Report_Type and ensure Employee_No exists
        .map(row => ({
          reportDate: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          employeeNo: String(row[1]),
          employeeDate: String(row[2]),
          reportType: type,
          dateRangeKey: String(row[4]),
          timestamp: row[5] ? row[5].toISOString() : ''
        }));
      
      allRecords = allRecords.concat(batchRecords);
    }
    
    return allRecords;
    
  } catch (error) {
    console.error(`Error loading ${type} audit trail: ${error.message}`);
    return [];
  }
}

/**
 * Checks if an employee has already been audited for the exact same date range
 * This function is kept for backward compatibility but should not be used for DBS reports
 * @deprecated Use isEmployeeAlreadyReportedForDBS instead for consistent logic
 * @param {string} employeeNo - Employee number
 * @param {Array} auditTrail - Audit trail records
 * @param {Object} range - Date range to check
 * @returns {boolean} True if already audited for this exact date range
 */
function isAlreadyAudited(employeeNo, auditTrail, range) {
  const key = `${Utilities.formatDate(range.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}_${Utilities.formatDate(range.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
  return auditTrail.some(r => r.employeeNo === employeeNo && r.dateRangeKey === key);
}

/**
 * Updates the audit trail with new DBS check records
 * @param {Array} records - Employee records to audit
 * @param {Object} range - Date range of the report
 * @param {string} type - Report type ('DBSCHECK')
 */
function updateAuditTrail(records, range, type) {
  const sheet = getAuditTrailSheet();
  const now = new Date();
  const key = `${Utilities.formatDate(range.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}_${Utilities.formatDate(range.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
  const newRows = records.map(emp => [now, emp.EmployeeNumber, emp.ContinuousServiceStartDate, type, key, now]);
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, newRows.length, 6).setValues(newRows);
}
