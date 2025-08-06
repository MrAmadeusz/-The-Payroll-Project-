/**
 * 08_currentEmployeesReport.gs
 * Weekly Current Employees Report - Simple snapshot of all current employees - SENT TO LEISU8RE ON THE WEEKLY
 */

/**
 * Generates a report of all current employees as of the report date
 * @param {Array} data - Employee data array
 * @param {Object} options - Options object (dateRange not used but kept for consistency)
 * @returns {Array} 2D array with report data
 */
function generateCurrentEmployeesReport(data, options = {}) {
  try {
    console.log('📋 Generating Current Employees Report...');
    
    // Filter for current employees only
    const currentEmployees = data.filter(emp => {
      const status = (emp.Status || '').toLowerCase().trim();
      return status === 'current';
    });
    
    console.log(`Found ${currentEmployees.length} current employees`);
    
    // Sort by Location, then Division, then Surname for consistent ordering
    currentEmployees.sort((a, b) => {
      const locationA = (a.Location || '').toLowerCase();
      const locationB = (b.Location || '').toLowerCase();
      if (locationA !== locationB) return locationA.localeCompare(locationB);
      
      const divisionA = (a.Division || '').toLowerCase();
      const divisionB = (b.Division || '').toLowerCase();
      if (divisionA !== divisionB) return divisionA.localeCompare(divisionB);
      
      const surnameA = (a.Surname || '').toLowerCase();
      const surnameB = (b.Surname || '').toLowerCase();
      return surnameA.localeCompare(surnameB);
    });
    
    // Build report
    const header = [
      'Employee Number',
      'First Names',
      'Surname', 
      'Location',
      'Division',
      'Start Date'
    ];
    
    const rows = currentEmployees.map(emp => [
      emp.EmployeeNumber || '',
      emp.Firstnames || '',
      emp.Surname || '',
      emp.Location || '',
      emp.Division || '',
      emp.ContinuousServiceStartDate || emp.StartDate || ''
    ]);
    
    // Build output with standard format (title, summary, blank, header, data)
    const reportDate = new Date().toLocaleDateString('en-GB');
    const titleRow = ['WEEKLY CURRENT EMPLOYEES REPORT', `Report Date: ${reportDate}`, '', '', '', ''];
    const summaryText = `${currentEmployees.length} current employees`;
    const summaryRow = ['SUMMARY', summaryText, '', '', '', ''];
    const blankRow = ['', '', '', '', '', ''];
    
    console.log(`✅ Current Employees report generated: ${rows.length} employees`);
    
    return [titleRow, summaryRow, blankRow, header, ...rows];
    
  } catch (error) {
    console.error(`Error generating current employees report: ${error.message}`);
    throw new Error('Failed to generate current employees report: ' + error.message);
  }
}

/**
 * Test function for the current employees report
 */
function testCurrentEmployeesReport() {
  console.log('🧪 Testing Current Employees Report...');
  
  try {
    const employees = getAllEmployees();
    console.log(`Loaded ${employees.length} total employees`);
    
    const output = generateCurrentEmployeesReport(employees);
    console.log(`Generated report with ${output.length} rows`);
    
    // Show sample output
    console.log('Sample output:');
    output.slice(0, 8).forEach((row, index) => {
      console.log(`Row ${index}: ${JSON.stringify(row)}`);
    });
    
    // Create test file
    const testFile = writeReportToStandaloneWorkbook(
      'weeklyCurrentEmployees', 
      output, 
      Session.getEffectiveUser().getEmail()
    );
    
    console.log(`✅ Test report created: ${testFile.getUrl()}`);
    
    return {
      success: true,
      totalEmployees: employees.length,
      currentEmployees: output.length - 4, // Minus header rows
      url: testFile.getUrl()
    };
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
 } 
}
