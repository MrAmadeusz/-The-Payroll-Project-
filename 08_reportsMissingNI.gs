/**
 * Generates a report of current employees missing National Insurance numbers
 * @param {Array} allEmployees - All employee data
 * @param {Object} options - Report options (not used for this report)
 * @returns {Array} Report data as 2D array for Google Sheets
 */
function generateMissingNIReport(allEmployees, options = {}) {
  console.log('🆔 Generating Missing NI Numbers Report...');
  
  // Filter for current employees with missing NI numbers
  const missingNI = allEmployees.filter(employee => {
    // Must be current status
    const isCurrent = (employee.Status || '').toLowerCase() === 'current';
    
    // Must have missing/empty NI Number
    const niNumber = employee.NINumber || '';
    const isMissingNI = niNumber.trim() === '' || niNumber.toLowerCase() === 'null';
    return isCurrent && isMissingNI;
  });
  
  console.log(`Found ${missingNI.length} current employees with missing NI numbers`);
  
  // Build the report output
  const headers = [
    'First Names',
    'Surname', 
    'Employee Number',
    'Location'
  ];
  
  // Add summary row
  const summaryRow = [
    '*=== SUMMARY ===*',
    `${missingNI.length} employees missing NI numbers`,
    `Generated: ${new Date().toLocaleDateString('en-GB')}`,
    `by ${Session.getEffectiveUser().getEmail()}`
  ];
  
  // Add empty row for spacing
  const emptyRow = ['', '', '', ''];
  
  // Create the output array
  const output = [
    summaryRow,
    emptyRow,
    headers
  ];
  
  // Add employee data
  missingNI.forEach(employee => {
    output.push([
      employee.Firstnames || '',
      employee.Surname || '',
      employee.EmployeeNumber || '',
      employee.Location || ''
    ]);
  });
  
  // If no missing NI numbers found, add a message
  if (missingNI.length === 0) {
    output.push(['✅ All current employees have NI numbers recorded', '', '', '']);
  }
  
  console.log(`✅ Missing NI Numbers report generated: ${output.length} rows`);
  return output;
}



function testMissingNIReport() {
  console.log('🧪 Testing Missing NI Numbers Report...');
  
  try {
    const employees = getAllEmployees();
    console.log(`Loaded ${employees.length} employees`);
    
    const output = generateMissingNIReport(employees);
    console.log(`Generated report with ${output.length} rows`);
    
    // Show first few rows for verification
    console.log('Sample output:');
    output.slice(0, 5).forEach((row, index) => {
      console.log(`Row ${index}: ${JSON.stringify(row)}`);
    });
    
    // Create test file
    const testFile = writeReportToStandaloneWorkbook('missingNINumbers', output, Session.getEffectiveUser().getEmail());
    console.log(`✅ Test report created: ${testFile.getUrl()}`);
    
    return {
      success: true,
      rowCount: output.length,
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
