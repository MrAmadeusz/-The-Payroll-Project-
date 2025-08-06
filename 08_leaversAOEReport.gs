/**
 * 08_leaversAOEReport.gs
 * Leavers AOE (Amount Owing Employee) Report
 * Identifies leavers with outstanding payments (non-zero ATT values)
 * 
 * IMPORTANT: This uses leavers logic but DOES NOT update the audit trail
 */

/**
 * Generates the Leavers AOE Report
 * Cross-references weekly leavers against Gross to Nett data for outstanding payments
 * 
 * @param {Array} data - Employee data array
 * @param {Object} options - Options object (can contain dateRange, lookbackDays)
 * @returns {Array} 2D array with report data
 */
function generateLeaversAOEReport(data, options = {}) {
  try {
    console.log('💰 Generating Leavers AOE Report...');
    
    const dateRange = options.dateRange || promptForDateRange("Select Date Range for Leavers AOE Report");
    
    // *** GET LEAVERS DATA (same logic as leavers report but NO audit trail updates) ***
    const leavers = getLeaversForAOEReport(data, dateRange, options);
    console.log(`Found ${leavers.length} leavers for the period`);
    
    if (leavers.length === 0) {
      console.log('No leavers found - returning empty report');
      return generateEmptyAOEReport(dateRange);
    }
    
    // *** LOAD LATEST GROSS TO NETT SNAPSHOT ***
    const grossToNettData = loadLatestGrossToNettSnapshot();
    console.log(`Loaded ${grossToNettData.length} records from Gross to Nett snapshot`);
    
    // *** CROSS-REFERENCE AND FILTER ***
    const leaversWithOutstandingPayments = [];
    
    leavers.forEach(leaver => {
      const employeeNumber = normalizeEmployeeNumber(leaver.EmployeeNumber);
      if (!employeeNumber) {
        console.warn(`Invalid employee number for leaver: ${leaver.EmployeeNumber}`);
        return;
      }
      
      // Find matching record in Gross to Nett data
      const grossRecord = grossToNettData.find(record => {
        const recordEmpNo = normalizeEmployeeNumber(record['Emp No.'] || record.EmployeeNumber);
        return recordEmpNo === employeeNumber;
      });
      
      if (grossRecord) {
        const attValue = parseFloat(grossRecord.ATT || grossRecord.att || '0');
        
        if (!isNaN(attValue) && attValue !== 0) {
          console.log(`Leaver ${employeeNumber} has outstanding payment: ${attValue}`);
          leaversWithOutstandingPayments.push({
            ...leaver,
            outstandingAmount: attValue,
            grossToNettRecord: grossRecord
          });
        } else {
          console.log(`Leaver ${employeeNumber} has zero/no outstanding payment`);
        }
      } else {
        console.warn(`Leaver ${employeeNumber} not found in Gross to Nett data`);
      }
    });
    
    console.log(`${leaversWithOutstandingPayments.length} leavers have outstanding payments`);
    
    // *** BUILD REPORT OUTPUT ***
    const header = [
      'Employee Number',
      'First Names', 
      'Surname',
      'Location',
      'Outstanding Amount (ATT)',
      'End Date'
    ];
    
    const rows = leaversWithOutstandingPayments.map(leaver => [
      leaver.EmployeeNumber || '',
      leaver.Firstnames || '',
      leaver.Surname || '',
      leaver.Location || '',
      leaver.outstandingAmount || 0,
      leaver.EndDate || leaver.LeavingDate || leaver.TerminationDate || leaver.LastWorkingDate || ''
    ]);
    
    // Build summary
    const rangeDisplay = formatDateRangeForDisplay(dateRange);
    const titleRow = ['WEEKLY LEAVERS AOE REPORT', rangeDisplay, '', '', '', ''];
    
    const summaryText = `${leaversWithOutstandingPayments.length} leavers with outstanding payments (from ${leavers.length} total leavers)`;
    const summaryRow = ['SUMMARY', summaryText, '', '', '', ''];
    const blankRow = ['', '', '', '', '', ''];
    
    console.log('✅ Leavers AOE report generated successfully');
    
    return [titleRow, summaryRow, blankRow, header, ...rows];
    
  } catch (error) {
    console.error(`Error generating Leavers AOE report: ${error.message}`);
    throw new Error('Failed to generate Leavers AOE report: ' + error.message);
  }
}

/**
 * Gets leavers data using the same logic as the main leavers report
 * BUT CRUCIALLY - does NOT update the audit trail
 * 
 * @param {Array} data - Employee data 
 * @param {Object} dateRange - Date range object
 * @param {Object} options - Options (lookbackDays, etc.)
 * @returns {Array} Array of leaver objects
 */
function getLeaversForAOEReport(data, dateRange, options = {}) {
  try {
    // Load audit trail for duplicate checking (READ ONLY - no updates!)
    const auditTrail = loadLeaversAuditTrail();
    
    // PART 1: Original logic - people who left in date range
    const normalLeavers = data.filter(emp => {
      const endDate = (emp.EndDate || emp.LeavingDate || emp.TerminationDate || emp.LastWorkingDate)?.trim();
      return endDate && isDateInRange(endDate, dateRange);
    }).filter(emp => 
      !isLeaverAlreadyReportedElsewhere(emp.EmployeeNumber, auditTrail, dateRange)
    );
    
    // PART 2: Late additions - people who left before range but never reported
    const lookbackDays = options.lookbackDays || 60;
    const lookbackDate = new Date(dateRange.startDate);
    lookbackDate.setDate(lookbackDate.getDate() - lookbackDays);
    
    const lateAdditions = data.filter(emp => {
      const endDate = (emp.EndDate || emp.LeavingDate || emp.TerminationDate || emp.LastWorkingDate)?.trim();
      if (!endDate) return false;
      
      const parsedEndDate = parseDMY(endDate);
      if (!parsedEndDate) return false;
      
      // Must have left before report range but after lookback date
      const leftBeforeRange = parsedEndDate < dateRange.startDate;
      const withinLookback = parsedEndDate >= lookbackDate;
      
      // Must never have been reported (not in audit trail at all)
      const neverReported = !auditTrail.some(record => record.employeeNo === emp.EmployeeNumber);
      
      return leftBeforeRange && withinLookback && neverReported;
    });
    
    // Combine both groups
    const allNewLeavers = [...normalLeavers, ...lateAdditions];
    
    // Remove duplicates (in case someone appears in both lists)
    const uniqueLeavers = allNewLeavers.filter((emp, index, arr) => 
      arr.findIndex(e => e.EmployeeNumber === emp.EmployeeNumber) === index
    );
    
    // Sort by end date
    uniqueLeavers.sort((a, b) => {
      const dateA = parseDMY(a.EndDate || a.LeavingDate || a.TerminationDate || a.LastWorkingDate);
      const dateB = parseDMY(b.EndDate || b.LeavingDate || b.TerminationDate || b.LastWorkingDate);
      return dateA - dateB;
    });
    
    console.log(`Leavers breakdown: ${normalLeavers.length} in period, ${lateAdditions.length} late additions = ${uniqueLeavers.length} total`);
    
    // *** CRITICAL: NO AUDIT TRAIL UPDATES - this is read-only for AOE purposes ***
    
    return uniqueLeavers;
    
  } catch (error) {
    console.error(`Error getting leavers for AOE report: ${error.message}`);
    throw error;
  }
}

/**
 * Generates an empty report when no leavers are found
 * @param {Object} dateRange - Date range object
 * @returns {Array} Empty report structure
 */
function generateEmptyAOEReport(dateRange) {
  const rangeDisplay = formatDateRangeForDisplay(dateRange);
  const titleRow = ['WEEKLY LEAVERS AOE REPORT', rangeDisplay, '', '', '', ''];
  const summaryRow = ['SUMMARY', 'No leavers found for this period', '', '', '', ''];
  const blankRow = ['', '', '', '', '', ''];
  const header = ['Employee Number', 'First Names', 'Surname', 'Location', 'Outstanding Amount (ATT)', 'End Date'];
  const noDataRow = ['No leavers with outstanding payments', '', '', '', '', ''];
  
  return [titleRow, summaryRow, blankRow, header, noDataRow];
}

/**
 * Test function for the Leavers AOE report
 */
function testLeaversAOEReport() {
  console.log('🧪 Testing Leavers AOE Report...');
  
  try {
    // First ensure we have snapshots
    console.log('Checking for Gross to Nett snapshot...');
    const grossData = loadLatestGrossToNettSnapshot();
    console.log(`✅ Found Gross to Nett snapshot with ${grossData.length} records`);
    
    // Load employee data
    const employees = getAllEmployees();
    console.log(`Loaded ${employees.length} total employees`);
    
    // Generate the report
    const output = generateLeaversAOEReport(employees);
    console.log(`Generated AOE report with ${output.length} rows`);
    
    // Show sample output
    console.log('\nSample output:');
    output.slice(0, 8).forEach((row, index) => {
      console.log(`Row ${index}: ${JSON.stringify(row)}`);
    });
    
    // Create test file
    const testFile = writeReportToStandaloneWorkbook(
      'leaversAOEReport', 
      output, 
      Session.getEffectiveUser().getEmail()
    );
    
    console.log(`✅ Test report created: ${testFile.getUrl()}`);
    
    return {
      success: true,
      totalRows: output.length,
      dataRows: Math.max(0, output.length - 4), // Minus header rows
      url: testFile.getUrl(),
      grossRecords: grossData.length
    };
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Quick analysis of outstanding payments data
 */
function analyzeLeaversAOEData() {
  console.log('🔍 Analyzing Leavers AOE data...');
  
  try {
    const grossData = loadLatestGrossToNettSnapshot();
    console.log(`Loaded ${grossData.length} Gross to Nett records`);
    
    // Analyze ATT values
    const attStats = { zero: 0, positive: 0, negative: 0, invalid: 0, total: grossData.length };
    const outstandingAmounts = [];
    
    grossData.forEach(record => {
      const attValue = parseFloat(record.ATT || record.att || '0');
      
      if (isNaN(attValue)) {
        attStats.invalid++;
      } else if (attValue === 0) {
        attStats.zero++;
      } else if (attValue > 0) {
        attStats.positive++;
        outstandingAmounts.push(attValue);
      } else {
        attStats.negative++;
        outstandingAmounts.push(Math.abs(attValue));
      }
    });
    
    console.log('\n📊 ATT Value Distribution:');
    console.log(`   Zero values: ${attStats.zero}`);
    console.log(`   Positive values: ${attStats.positive}`);
    console.log(`   Negative values: ${attStats.negative}`);
    console.log(`   Invalid values: ${attStats.invalid}`);
    
    if (outstandingAmounts.length > 0) {
      const total = outstandingAmounts.reduce((sum, val) => sum + val, 0);
      const average = total / outstandingAmounts.length;
      const max = Math.max(...outstandingAmounts);
      const min = Math.min(...outstandingAmounts);
      
      console.log(`\n💰 Outstanding Amounts:`);
      console.log(`   Total: £${total.toFixed(2)}`);
      console.log(`   Average: £${average.toFixed(2)}`);
      console.log(`   Max: £${max.toFixed(2)}`);
      console.log(`   Min: £${min.toFixed(2)}`);
    }
    
    return {
      totalRecords: grossData.length,
      attStats: attStats,
      outstandingTotal: outstandingAmounts.reduce((sum, val) => sum + val, 0)
    };
    
  } catch (error) {
    console.error('❌ Analysis failed:', error.message);
    return { error: error.message };
  }
}
