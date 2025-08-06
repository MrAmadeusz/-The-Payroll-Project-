/**
 * Generates a report of current employees with incorrect contract type classifications
 * @param {Array} allEmployees - All employee data
 * @param {Object} options - Report options (not used for this report)
 * @returns {Array} Report data as 2D array for Google Sheets
 */

function generateContractCheckReport(allEmployees, options = {}) {
  console.log('📋 Generating Contract Check Report...');
  
  // Filter for current employees with contract type issues
  const contractIssues = allEmployees.filter(employee => {
    // Must be current status
    const isCurrent = (employee.Status || '').toLowerCase() === 'current';
    if (!isCurrent) return false;
    
    // Exclude employees on old contracts
    const normalizedEmpNo = normalizeEmployeeNumber(employee.EmployeeNumber);
    if (CONFIG.CONTRACT_CHECK_EXCLUSIONS && CONFIG.CONTRACT_CHECK_EXCLUSIONS.includes(normalizedEmpNo)) {
      console.log(`Excluding employee ${normalizedEmpNo} - on old contract`);
      return false;
    }
    
    const payType = (employee.PayType || '').trim();
    const contractHours = parseFloat(employee.ContractHours || '0');
    const contractType = (employee.ContractType || '').trim();
    
    // Rule 1: Hourly + Zero Hours should be Casual
    const rule1 = payType === 'Hourly' && 
                  contractHours === 0 && 
                  contractType !== 'Casual';
    
    // Rule 2: Hourly + Hours should be Flexible  
    const rule2 = payType === 'Hourly' && 
                  contractHours > 0 && 
                  contractType !== 'Flexible';
    
    // Rule 3: Salary should be Full Time
    const rule3 = payType === 'Salary' && 
                  contractType !== 'Full Time';
    
    return rule1 || rule2 || rule3;
  });
  
  console.log(`Found ${contractIssues.length} current employees with contract type issues (after excluding ${CONFIG.CONTRACT_CHECK_EXCLUSIONS?.length || 0} old contracts)`);
  
  // Build the report output
  const headers = [
    'First Names',
    'Surname', 
    'Employee Number',
    'Location',
    'Pay Type',
    'Contract Hours',
    'Current Contract Type',
    'Expected Contract Type',
    'Issue Description'
  ];
  
  // Add summary row
  const summaryRow = [
    '*== SUMMARY ==*',
    `${contractIssues.length} employees with contract issues`,
    `Generated: ${new Date().toLocaleDateString('en-GB')}`,
    `Excluded old contracts: ${CONFIG.CONTRACT_CHECK_EXCLUSIONS?.length || 0}`,
    '', '', '', '', ''
  ];
  
  // Add empty row for spacing
  const emptyRow = ['', '', '', '', '', '', '', '', ''];
  
  // Create the output array
  const output = [
    summaryRow,
    emptyRow,
    headers
  ];
  
  // Add employee data with issue analysis
  contractIssues.forEach(employee => {
    const payType = (employee.PayType || '').trim();
    const contractHours = parseFloat(employee.ContractHours || '0');
    const contractType = (employee.ContractType || '').trim();
    
    // Determine expected contract type and issue description
    let expectedType = '';
    let issueDescription = '';
    
    if (payType === 'Hourly' && contractHours === 0) {
      expectedType = 'Casual';
      issueDescription = 'Hourly worker with 0 hours should be Casual';
    } else if (payType === 'Hourly' && contractHours > 0) {
      expectedType = 'Flexible';
      issueDescription = `Hourly worker with ${contractHours} hours should be Flexible`;
    } else if (payType === 'Salary') {
      expectedType = 'Full Time';
      issueDescription = 'Salaried worker should be Full Time';
    }
    
    output.push([
      employee.Firstnames || '',
      employee.Surname || '',
      employee.EmployeeNumber || '',
      employee.Location || '',
      payType,
      contractHours || 0,
      contractType,
      expectedType,
      issueDescription
    ]);
  });
  
  // If no contract issues found, add a message
  if (contractIssues.length === 0) {
    output.push([
      '✅ All current employees have correct contract type classifications (excluding old contracts)', 
      '', '', '', '', '', '', '', ''
    ]);
  }
  
  console.log(`✅ Contract Check report generated: ${output.length} rows`);
  return output;
}

function analyzeContractCompliance() {
  console.log('🔍 Analyzing contract compliance across all current employees...');
  
  const employees = getAllEmployees();
  const current = employees.filter(emp => (emp.Status || '').toLowerCase() === 'current');
  
  let compliant = 0;
  let issues = 0;
  const breakdown = {
    hourlyZeroCorrect: 0,
    hourlyZeroWrong: 0,
    hourlyHoursCorrect: 0,
    hourlyHoursWrong: 0,
    salaryCorrect: 0,
    salaryWrong: 0
  };
  
  current.forEach(emp => {
    const payType = (emp.PayType || '').trim();
    const contractHours = parseFloat(emp.ContractHours || '0');
    const contractType = (emp.ContractType || '').trim();
    
    if (payType === 'Hourly' && contractHours === 0) {
      if (contractType === 'Casual') {
        breakdown.hourlyZeroCorrect++;
        compliant++;
      } else {
        breakdown.hourlyZeroWrong++;
        issues++;
      }
    } else if (payType === 'Hourly' && contractHours > 0) {
      if (contractType === 'Flexible') {
        breakdown.hourlyHoursCorrect++;
        compliant++;
      } else {
        breakdown.hourlyHoursWrong++;
        issues++;
      }
    } else if (payType === 'Salary') {
      if (contractType === 'Full Time') {
        breakdown.salaryCorrect++;
        compliant++;
      } else {
        breakdown.salaryWrong++;
        issues++;
      }
    }
  });
  
  const complianceRate = ((compliant / (compliant + issues)) * 100).toFixed(1);
  
  console.log(`📊 Contract Compliance Analysis:`);
  console.log(`Total Current Employees: ${current.length}`);
  console.log(`Compliant: ${compliant} (${complianceRate}%)`);
  console.log(`Issues: ${issues}`);
  console.log(`Breakdown:`, breakdown);
  
  return {
    total: current.length,
    compliant: compliant,
    issues: issues,
    complianceRate: complianceRate,
    breakdown: breakdown
  };
}
