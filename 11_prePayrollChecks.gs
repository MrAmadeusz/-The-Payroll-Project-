/**
 * PRE-PAYROLL VALIDATION CHECKS ()
 * Clean, simple validation functions for payroll data
 * Each function returns standardized results format
 * MEMORY OPTIMIZED: All functions read from chunked cache, no parameters needed
 * 
 * FIXES:
 * - Improved cache loading with error handling
 * - Better exception generation with unique IDs
 * - Comprehensive validation logic
 * - Robust error recovery
 */

// Standardized return format for all checks
const CHECK_RESULT_TEMPLATE = {
  checkName: "",
  status: "pass", // pass | warning | fail | error
  exceptions: [], // Array of issue objects
  summary: "",
  executionTime: 0,
  recordsChecked: 0,
  details: {}
};

/**
 * 1. SALARY VALIDATION CHECK (FIXED VERSION)
 * Validates salary calculations before final payroll processing
 */
function checkSalaryValidation() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Salary Validation";
  result.exceptions = [];
  
  try {
    console.log("Running salary validation check on PRE-payroll data...");
    
    // Load data from chunked cache
    const cache = CacheService.getScriptCache();
    const employees = loadChunkedDataFromCache(cache, 'employees');
    const payrollData = loadPayrollDataFromCache();
    
    if (employees.length === 0) {
      throw new Error("No employees found in cache");
    }
    
    if (payrollData.totalRecords === 0) {
      throw new Error("No payroll data found in cache");
    }
    
    // Filter for salary employees only
    const salaryEmployees = employees.filter(emp => 
      normalizePayType(emp.PayType || emp.payType) === "Salary"
    );
    
    console.log(`🔍 Checking ${salaryEmployees.length} salary employees against PRE-payroll data`);
    result.recordsChecked = salaryEmployees.length;
    
    // Get salary data from PRE-payroll files
    const salaryGrossData = payrollData.salaryGrossToNett || [];
    const salaryPaymentsData = payrollData.salaryPaymentsDeductions || [];
    
    console.log(`📊 Found ${salaryGrossData.length} salary gross records, ${salaryPaymentsData.length} salary payment records`);
    
    // Create employee lookup for faster processing
    const employeeIndex = buildDataIndex(salaryEmployees, 'EmployeeNumber');
    
    // Check salary gross to nett calculations
    salaryGrossData.forEach(salaryRecord => {
      const empNumber = normalizeEmployeeNumber(salaryRecord.EmployeeNumber || salaryRecord.employeeNumber);
      const employee = employeeIndex[empNumber]?.[0];
      
      if (!employee) {
        result.exceptions.push(createStandardException(
          'Salary Validation',
          empNumber,
          salaryRecord.EmployeeName || salaryRecord.employeeName || 'Unknown',
          'Salary calculation found for employee not in master data',
          'critical',
          {
            grossAmount: salaryRecord.GrossAmount || salaryRecord.grossAmount,
            sourceFile: salaryRecord._sourceFile
          }
        ));
        return;
      }
      
      // Validate expected salary vs calculated gross
      const expectedSalary = parseFloat(employee.Salary || employee.salary || 0);
      const calculatedGross = parseFloat(salaryRecord.GrossAmount || salaryRecord.grossAmount || 0);
      
      // Allow for small differences (£5 tolerance for monthly calculations)
      const difference = Math.abs(expectedSalary - calculatedGross);
      
      if (difference > 5.00) {
        const severity = difference > 100 ? 'critical' : 'warning';
        
        result.exceptions.push(createStandardException(
          'Salary Validation',
          empNumber,
          employee.Name || employee.EmployeeName,
          `Calculated salary differs from master data: expected £${expectedSalary}, calculated £${calculatedGross}`,
          severity,
          {
            expectedSalary: expectedSalary,
            calculatedGross: calculatedGross,
            difference: difference,
            percentageDiff: expectedSalary > 0 ? ((difference / expectedSalary) * 100).toFixed(2) : 'N/A',
            sourceFile: salaryRecord._sourceFile,
            location: employee.Location,
            department: employee.Department
          }
        ));
      }
      
      // Check for negative or zero gross amounts
      if (calculatedGross <= 0) {
        result.exceptions.push(createStandardException(
          'Salary Validation',
          empNumber,
          employee.Name || employee.EmployeeName,
          `Invalid gross amount: £${calculatedGross}`,
          'critical',
          {
            calculatedGross: calculatedGross,
            expectedSalary: expectedSalary,
            sourceFile: salaryRecord._sourceFile
          }
        ));
      }
    });
    
    // Check for missing salary employees in PRE-payroll data
    salaryEmployees.forEach(employee => {
      const empNumber = normalizeEmployeeNumber(employee.EmployeeNumber);
      const hasGrossRecord = salaryGrossData.some(record => 
        normalizeEmployeeNumber(record.EmployeeNumber || record.employeeNumber) === empNumber
      );
      
      // Only flag as missing if employee is active
      if (!hasGrossRecord && (employee.Status || '').toLowerCase() === 'active') {
        result.exceptions.push(createStandardException(
          'Salary Validation',
          empNumber,
          employee.Name || employee.EmployeeName,
          'Active salary employee missing from PRE-payroll calculations',
          'critical',
          {
            expectedSalary: parseFloat(employee.Salary || employee.salary || 0),
            location: employee.Location,
            department: employee.Department,
            status: employee.Status
          }
        ));
      }
    });
    
    // Set overall result status
    if (result.exceptions.length === 0) {
      result.status = "pass";
      result.summary = `All ${salaryEmployees.length} salary employees validated successfully against PRE-payroll data`;
    } else {
      const criticalCount = result.exceptions.filter(e => e.severity === 'critical').length;
      result.status = criticalCount > 0 ? "fail" : "warning";
      result.summary = `Found ${result.exceptions.length} salary validation issues (${criticalCount} critical)`;
    }
    
  } catch (error) {
    result.status = "error";
    result.summary = `Salary validation failed: ${error.message}`;
    result.error = error.message;
    console.error("Salary validation error:", error);
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

/**
 * 2. NMW COMPLIANCE CHECK (FIXED VERSION)
 * Validates hourly rates meet NMW requirements before payroll processing
 */
function checkNMWCompliance() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "National Minimum Wage Compliance";
  result.exceptions = [];
  
  try {
    console.log("Running NMW compliance check on PRE-payroll data...");
    
    // Load data from chunked cache
    const cache = CacheService.getScriptCache();
    const employees = loadChunkedDataFromCache(cache, 'employees');
    const payrollData = loadPayrollDataFromCache();
    
    if (employees.length === 0) {
      throw new Error("No employees found in cache");
    }
    
    if (payrollData.totalRecords === 0) {
      throw new Error("No payroll data found in cache");
    }
    
    // Filter for hourly employees only
    const hourlyEmployees = employees.filter(emp => 
      normalizePayType(emp.PayType || emp.payType) === "Hourly"
    );
    
    console.log(`🔍 Checking ${hourlyEmployees.length} hourly employees against PRE-payroll data`);
    result.recordsChecked = hourlyEmployees.length;
    
    // Load current NMW rates
    const nmwRates = getCurrentNMWRates();
    
    // Get hourly data from PRE-payroll files
    const hourlyGrossData = payrollData.hourlyGrossToNett || [];
    
    console.log(`⏰ Found ${hourlyGrossData.length} hourly gross records`);
    
    const employeeIndex = buildDataIndex(hourlyEmployees, 'EmployeeNumber');
    
    // Check each hourly calculation
    hourlyGrossData.forEach(hourlyRecord => {
      const empNumber = normalizeEmployeeNumber(hourlyRecord.EmployeeNumber || hourlyRecord.employeeNumber);
      const employee = employeeIndex[empNumber]?.[0];
      
      if (!employee) {
        result.exceptions.push(createStandardException(
          'NMW Compliance',
          empNumber,
          hourlyRecord.EmployeeName || hourlyRecord.employeeName || 'Unknown',
          'Hourly calculation found for employee not in master data',
          'warning',
          { sourceFile: hourlyRecord._sourceFile }
        ));
        return;
      }
      
      // Calculate effective hourly rate from PRE-payroll data
      const grossPay = parseFloat(hourlyRecord.GrossAmount || hourlyRecord.grossAmount || 0);
      const hoursWorked = parseFloat(hourlyRecord.HoursWorked || hourlyRecord.hoursWorked || 
                                   hourlyRecord.TotalHours || hourlyRecord.totalHours || 0);
      
      if (hoursWorked === 0) {
        result.exceptions.push(createStandardException(
          'NMW Compliance',
          empNumber,
          employee.Name || employee.EmployeeName,
          'Zero hours worked but gross pay calculated',
          'warning',
          {
            grossPay: grossPay,
            hoursWorked: hoursWorked,
            sourceFile: hourlyRecord._sourceFile
          }
        ));
        return;
      }
      
      const effectiveRate = grossPay / hoursWorked;
      
      // Determine appropriate NMW rate based on employee data
      const age = calculateAge(employee.DateOfBirth || employee.dateOfBirth);
      const isApprentice = (employee.PayCategory || employee.payCategory || '').toLowerCase().includes('apprentice');
      
      let requiredRate;
      let rateType;
      
      if (isApprentice && age < 19) {
        requiredRate = nmwRates.apprentice;
        rateType = 'Apprentice';
      } else if (age < 18) {
        requiredRate = nmwRates.under18;
        rateType = 'Under 18';
      } else if (age < 21) {
        requiredRate = nmwRates.development;
        rateType = '18-20 Development';
      } else {
        requiredRate = nmwRates.adult;
        rateType = 'Adult (21+)';
      }
      
      // Check compliance (allow 1p tolerance for rounding)
      if (effectiveRate < (requiredRate - 0.01)) {
        const shortfall = requiredRate - effectiveRate;
        
        result.exceptions.push(createStandardException(
          'NMW Compliance',
          empNumber,
          employee.Name || employee.EmployeeName,
          `Calculated rate £${effectiveRate.toFixed(2)}/hr below ${rateType} NMW £${requiredRate.toFixed(2)}/hr`,
          'critical',
          {
            effectiveRate: effectiveRate.toFixed(2),
            requiredRate: requiredRate.toFixed(2),
            rateType: rateType,
            shortfall: shortfall.toFixed(2),
            grossPay: grossPay,
            hoursWorked: hoursWorked,
            age: age,
            isApprentice: isApprentice,
            location: employee.Location,
            sourceFile: hourlyRecord._sourceFile
          }
        ));
      }
    });
    
    // Check for missing hourly employees in PRE-payroll data
    hourlyEmployees.forEach(employee => {
      const empNumber = normalizeEmployeeNumber(employee.EmployeeNumber);
      const hasGrossRecord = hourlyGrossData.some(record => 
        normalizeEmployeeNumber(record.EmployeeNumber || record.employeeNumber) === empNumber
      );
      
      if (!hasGrossRecord && (employee.Status || '').toLowerCase() === 'active') {
        result.exceptions.push(createStandardException(
          'NMW Compliance',
          empNumber,
          employee.Name || employee.EmployeeName,
          'Active hourly employee missing from PRE-payroll calculations',
          'warning',
          {
            location: employee.Location,
            department: employee.Department,
            payType: employee.PayType,
            status: employee.Status
          }
        ));
      }
    });
    
    // Set result status
    if (result.exceptions.length === 0) {
      result.status = "pass";
      result.summary = `All ${hourlyEmployees.length} hourly employees meet NMW requirements in PRE-payroll data`;
    } else {
      const criticalCount = result.exceptions.filter(e => e.severity === 'critical').length;
      result.status = criticalCount > 0 ? "fail" : "warning";
      result.summary = `Found ${result.exceptions.length} NMW compliance issues (${criticalCount} critical)`;
    }
    
  } catch (error) {
    result.status = "error";
    result.summary = `NMW compliance check failed: ${error.message}`;
    result.error = error.message;
    console.error("NMW compliance error:", error);
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

/**
 * 3. MISSING/BLANK PAYMENTS CHECK (FIXED VERSION)
 * Identifies employees with missing or blank payment calculations
 */
function checkMissingBlankPayments() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Missing/Blank Payments";
  result.exceptions = [];
  
  try {
    console.log("Running missing/blank payments check on PRE-payroll data...");
    
    // Load data from chunked cache
    const cache = CacheService.getScriptCache();
    const employees = loadChunkedDataFromCache(cache, 'employees');
    const payrollData = loadPayrollDataFromCache();
    
    if (employees.length === 0) {
      throw new Error("No employees found in cache");
    }
    
    if (payrollData.totalRecords === 0) {
      throw new Error("No payroll data found in cache");
    }
    
    // Get all active employees
    const activeEmployees = employees.filter(emp => 
      (emp.Status || '').toLowerCase() === 'active'
    );
    
    console.log(`🔍 Checking ${activeEmployees.length} active employees for missing payments`);
    result.recordsChecked = activeEmployees.length;
    
    // Combine all PRE-payroll gross data
    const allGrossData = [
      ...(payrollData.hourlyGrossToNett || []),
      ...(payrollData.salaryGrossToNett || [])
    ];
    
    console.log(`📊 Checking against ${allGrossData.length} PRE-payroll gross records`);
    
    // Check each active employee
    activeEmployees.forEach(employee => {
      const empNumber = normalizeEmployeeNumber(employee.EmployeeNumber);
      
      // Look for gross payment record
      const grossRecord = allGrossData.find(record => 
        normalizeEmployeeNumber(record.EmployeeNumber || record.employeeNumber) === empNumber
      );
      
      if (!grossRecord) {
        result.exceptions.push(createStandardException(
          'Missing/Blank Payments',
          empNumber,
          employee.Name || employee.EmployeeName,
          'Active employee has no payment calculation in PRE-payroll data',
          'critical',
          {
            payType: employee.PayType,
            location: employee.Location,
            department: employee.Department,
            status: employee.Status,
            expectedSalary: employee.PayType === 'Salary' ? parseFloat(employee.Salary || 0) : null
          }
        ));
      } else {
        // Check for blank/zero gross amounts
        const grossAmount = parseFloat(grossRecord.GrossAmount || grossRecord.grossAmount || 0);
        
        if (grossAmount === 0 || isNaN(grossAmount)) {
          result.exceptions.push(createStandardException(
            'Missing/Blank Payments',
            empNumber,
            employee.Name || employee.EmployeeName,
            'Employee has zero or blank gross amount in PRE-payroll calculation',
            'critical',
            {
              grossAmount: grossAmount,
              payType: employee.PayType,
              sourceFile: grossRecord._sourceFile,
              location: employee.Location
            }
          ));
        }
        
        // Check for missing required fields based on pay type
        if (employee.PayType === 'Hourly') {
          const hoursWorked = parseFloat(grossRecord.HoursWorked || grossRecord.hoursWorked || 
                                       grossRecord.TotalHours || grossRecord.totalHours || 0);
          
          if (hoursWorked === 0 && grossAmount > 0) {
            result.exceptions.push(createStandardException(
              'Missing/Blank Payments',
              empNumber,
              employee.Name || employee.EmployeeName,
              'Hourly employee has payment but zero hours recorded',
              'warning',
              {
                grossAmount: grossAmount,
                hoursWorked: hoursWorked,
                sourceFile: grossRecord._sourceFile
              }
            ));
          }
        }
      }
    });
    
    // Set result status
    if (result.exceptions.length === 0) {
      result.status = "pass";
      result.summary = `All ${activeEmployees.length} active employees have valid payment calculations`;
    } else {
      const criticalCount = result.exceptions.filter(e => e.severity === 'critical').length;
      result.status = criticalCount > 0 ? "fail" : "warning";
      result.summary = `Found ${result.exceptions.length} missing/blank payment issues (${criticalCount} critical)`;
    }
    
  } catch (error) {
    result.status = "error";
    result.summary = `Missing/blank payments check failed: ${error.message}`;
    result.error = error.message;
    console.error("Missing/blank payments error:", error);
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

/**
 * 4. DUPLICATE BACKPAYS CHECK (FIXED VERSION)
 * Identifies duplicate backpay entries in PRE-payroll data
 */
function checkDuplicateBackpays() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Duplicate Backpays";
  result.exceptions = [];
  
  try {
    console.log("Running duplicate backpays check on PRE-payroll data...");
    
    // Load data from chunked cache
    const cache = CacheService.getScriptCache();
    const employees = loadChunkedDataFromCache(cache, 'employees');
    const payrollData = loadPayrollDataFromCache();
    
    if (payrollData.totalRecords === 0) {
      throw new Error("No payroll data found in cache");
    }
    
    // Get all payment/deduction records that might contain backpays
    const allPaymentRecords = [
      ...(payrollData.hourlyPaymentsDeductions || []),
      ...(payrollData.salaryPaymentsDeductions || [])
    ];
    
    // Filter for backpay-related entries
    const backpayRecords = allPaymentRecords.filter(record => {
      const description = (record.Description || record.description || record.PaymentType || '').toLowerCase();
      return description.includes('backpay') || 
             description.includes('back pay') || 
             description.includes('arrears') ||
             description.includes('adjustment') ||
             description.includes('retro');
    });
    
    console.log(`🔍 Checking ${backpayRecords.length} potential backpay records`);
    result.recordsChecked = backpayRecords.length;
    
    // Group backpays by employee
    const backpaysByEmployee = {};
    
    backpayRecords.forEach(record => {
      const empNumber = normalizeEmployeeNumber(record.EmployeeNumber || record.employeeNumber);
      if (!backpaysByEmployee[empNumber]) {
        backpaysByEmployee[empNumber] = [];
      }
      backpaysByEmployee[empNumber].push(record);
    });
    
    // Check each employee's backpays for duplicates
    Object.keys(backpaysByEmployee).forEach(empNumber => {
      const employeeBackpays = backpaysByEmployee[empNumber];
      
      if (employeeBackpays.length < 2) return; // Can't have duplicates with only 1 payment
      
      const employee = employees.find(emp => 
        normalizeEmployeeNumber(emp.EmployeeNumber) === empNumber
      );
      
      // Look for potential duplicates
      for (let i = 0; i < employeeBackpays.length - 1; i++) {
        for (let j = i + 1; j < employeeBackpays.length; j++) {
          const payment1 = employeeBackpays[i];
          const payment2 = employeeBackpays[j];
          
          const amount1 = parseFloat(payment1.Amount || payment1.amount || payment1.GrossAmount || 0);
          const amount2 = parseFloat(payment2.Amount || payment2.amount || payment2.GrossAmount || 0);
          
          const desc1 = (payment1.Description || payment1.description || payment1.PaymentType || '').toLowerCase();
          const desc2 = (payment2.Description || payment2.description || payment2.PaymentType || '').toLowerCase();
          
          // Check if amounts match exactly and both are positive
          if (Math.abs(amount1 - amount2) < 0.01 && amount1 > 0) {
            // Check if descriptions are similar
            const similarity = calculateStringSimilarity(desc1, desc2);
            
            if (similarity > 0.8) { // 80% similarity threshold
              result.exceptions.push(createStandardException(
                'Duplicate Backpays',
                empNumber,
                employee?.Name || employee?.EmployeeName || 'Unknown',
                `Potential duplicate backpay: £${amount1.toFixed(2)} appears twice with similar descriptions`,
                'warning',
                {
                  amount: amount1,
                  description1: payment1.Description || payment1.description || payment1.PaymentType,
                  description2: payment2.Description || payment2.description || payment2.PaymentType,
                  similarity: (similarity * 100).toFixed(1) + '%',
                  sourceFile1: payment1._sourceFile,
                  sourceFile2: payment2._sourceFile,
                  paymentPeriods: [
                    payment1.PayPeriod || payment1.payPeriod || payment1.Period,
                    payment2.PayPeriod || payment2.payPeriod || payment2.Period
                  ]
                }
              ));
            }
          }
        }
      }
    });
    
    // Set result status
    if (result.exceptions.length === 0) {
      result.status = "pass";
      result.summary = `No duplicate backpays detected in ${backpayRecords.length} backpay records`;
    } else {
      result.status = "warning";
      result.summary = `Found ${result.exceptions.length} potential duplicate backpay entries`;
    }
    
  } catch (error) {
    result.status = "error";
    result.summary = `Duplicate backpays check failed: ${error.message}`;
    result.error = error.message;
    console.error("Duplicate backpays error:", error);
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

/**
 * 5. NET NEGATIVE PAYMENTS CHECK (PRODUCTION IMPLEMENTATION)
 * Validates net payment amounts from PRE-payroll Gross to Nett files
 */
function checkNetNegativePayments() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Net Negative Payments";
  result.exceptions = [];
  
  try {
    console.log("Running net negative payments check on PRE-payroll data...");
    
    // Load data from chunked cache
    const cache = CacheService.getScriptCache();
    const employees = loadChunkedDataFromCache(cache, 'employees');
    const payrollData = loadPayrollDataFromCache();
    
    if (payrollData.totalRecords === 0) {
      throw new Error("No payroll data found in cache");
    }
    
    // Combine both Gross to Nett files (hourly and salary)
    const allGrossToNettRecords = [
      ...(payrollData.hourlyGrossToNett || []),
      ...(payrollData.salaryGrossToNett || [])
    ];
    
    console.log(`🔍 Checking ${allGrossToNettRecords.length} gross to nett records for negative/excessive deductions`);
    result.recordsChecked = allGrossToNettRecords.length;
    
    // Track statistics for reporting
    let negativeCount = 0;
    let excessiveDeductionCount = 0;
    let processedCount = 0;
    
    // Process records in smaller batches to avoid memory issues
    const batchSize = 1000;
    
    for (let i = 0; i < allGrossToNettRecords.length; i += batchSize) {
      const batch = allGrossToNettRecords.slice(i, i + batchSize);
      
      console.log(`📦 Processing batch ${Math.floor(i/batchSize) + 1}/${Math.ceil(allGrossToNettRecords.length/batchSize)} (${batch.length} records)`);
      
      // Check each record in the batch
      batch.forEach(record => {
        try {
          processedCount++;
          
          // Extract key fields using flexible column name matching
          const empNumber = normalizeEmployeeNumber(
            record['Emp No.'] || record['EmpNo'] || record['EmployeeNumber'] || 
            record['Employee Number'] || record.EmployeeNumber || ''
          );
          const employeeName = (
            record['Employee Name'] || record['EmployeeName'] || 
            record.EmployeeName || record.Name || ''
          ).trim();
          const grossPay = parseFloat(
            record['GP'] || record['Gross'] || record['GrossPay'] || 
            record.GrossAmount || record.GrossAmount || 0
          );
          const netPay = parseFloat(
            record['NP'] || record['Net'] || record['NetPay'] || 
            record.NetAmount || record.NetAmount || 0
          );
          const location = (record['Location'] || record.Location || '').trim();
          const division = (record['Division'] || record.Division || '').trim();
          
          // Skip records with missing essential data
          if (!empNumber || empNumber === 'unknown') {
            return;
          }
          
          // Find employee in master data for additional context
          const employee = employees.find(emp => 
            normalizeEmployeeNumber(emp.EmployeeNumber) === empNumber
          );
          
          // CRITICAL CHECK: Any negative net pay
          if (netPay < 0) {
            negativeCount++;
            
            result.exceptions.push(createStandardException(
              'Net Negative Payments',
              empNumber,
              employeeName || employee?.Name || employee?.EmployeeName || 'Unknown',
              `CRITICAL: Negative net pay of £${netPay.toFixed(2)}`,
              'critical',
              {
                netPay: netPay,
                grossPay: grossPay,
                totalDeductions: grossPay - netPay,
                deductionAmount: (grossPay - netPay).toFixed(2),
                deductionPercentage: grossPay > 0 ? ((grossPay - netPay) / grossPay * 100).toFixed(1) + '%' : 'N/A',
                location: location,
                division: division,
                sourceFile: record._sourceFile,
                payType: record._sourceType?.includes('hourly') ? 'Hourly' : 'Salary'
              }
            ));
          }
          // WARNING CHECK: Net pay < 50% of gross (but not negative)
          else if (grossPay > 0 && netPay >= 0 && (netPay / grossPay) < 0.5) {
            excessiveDeductionCount++;
            
            const deductionPercentage = ((grossPay - netPay) / grossPay * 100).toFixed(1);
            
            result.exceptions.push(createStandardException(
              'Net Negative Payments',
              empNumber,
              employeeName || employee?.Name || employee?.EmployeeName || 'Unknown',
              `WARNING: Net pay £${netPay.toFixed(2)} is only ${(netPay / grossPay * 100).toFixed(1)}% of gross pay (${deductionPercentage}% deducted)`,
              'warning',
              {
                netPay: netPay,
                grossPay: grossPay,
                netPercentage: (netPay / grossPay * 100).toFixed(1) + '%',
                totalDeductions: grossPay - netPay,
                deductionAmount: (grossPay - netPay).toFixed(2),
                deductionPercentage: deductionPercentage + '%',
                location: location,
                division: division,
                sourceFile: record._sourceFile,
                payType: record._sourceType?.includes('hourly') ? 'Hourly' : 'Salary'
              }
            ));
          }
          
          // VALIDATION CHECK: Gross pay but no net calculation
          if (grossPay > 0 && netPay === 0) {
            result.exceptions.push(createStandardException(
              'Net Negative Payments',
              empNumber,
              employeeName || employee?.Name || employee?.EmployeeName || 'Unknown',
              'WARNING: Gross pay exists but net pay is zero - possible calculation error',
              'warning',
              {
                netPay: netPay,
                grossPay: grossPay,
                issue: 'Zero net with positive gross',
                location: location,
                division: division,
                sourceFile: record._sourceFile,
                payType: record._sourceType?.includes('hourly') ? 'Hourly' : 'Salary'
              }
            ));
          }
          
        } catch (recordError) {
          console.warn(`⚠️ Error processing record for employee ${record['Emp No.'] || 'unknown'}: ${recordError.message}`);
          // Continue processing other records
        }
      });
    }
    
    // Set result status based on findings
    if (result.exceptions.length === 0) {
      result.status = "pass";
      result.summary = `All ${allGrossToNettRecords.length} payment calculations have acceptable net amounts`;
    } else {
      const criticalCount = result.exceptions.filter(e => e.severity === 'critical').length;
      const warningCount = result.exceptions.filter(e => e.severity === 'warning').length;
      
      if (criticalCount > 0) {
        result.status = "fail";
        result.summary = `Found ${criticalCount} CRITICAL negative payments and ${warningCount} excessive deduction warnings`;
      } else {
        result.status = "warning";
        result.summary = `Found ${warningCount} payments with excessive deductions (>50% of gross)`;
      }
    }
    
    // Add detailed statistics to the result
    result.details = {
      recordsProcessed: processedCount,
      negativePayments: negativeCount,
      excessiveDeductions: excessiveDeductionCount,
      hourlyRecords: payrollData.hourlyGrossToNett?.length || 0,
      salaryRecords: payrollData.salaryGrossToNett?.length || 0,
      checkCriteria: {
        critical: 'Net pay < £0.00',
        warning: 'Net pay < 50% of gross pay'
      }
    };
    
    console.log(`✅ Net Negative Payments check completed: ${result.exceptions.length} exceptions found`);
    console.log(`📊 Statistics: ${negativeCount} negative, ${excessiveDeductionCount} excessive deductions`);
    
  } catch (error) {
    result.status = "error";
    result.summary = `Net negative payments check failed: ${error.message}`;
    result.error = error.message;
    console.error("❌ Net negative payments check error:", error);
    
    // Add error details for debugging
    result.details = {
      error: error.message,
      stack: error.stack
    };
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

/**
 * 6-13. PLACEHOLDER CHECKS (Updated with better error handling)
 */
function checkEVSchemePayments() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "EV Scheme Payments";
  
  try {
    console.log("Running EV scheme payments check...");
    result.status = "pass";
    result.summary = "Placeholder - EV scheme check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `EV scheme check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

function checkCarAllowance() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Car Allowance";
  
  try {
    console.log("Running car allowance check...");
    result.status = "pass";
    result.summary = "Placeholder - car allowance check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `Car allowance check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

function checkPreviewVsUpload() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Preview vs Upload Files";
  
  try {
    console.log("Running preview vs upload files check...");
    result.status = "pass";
    result.summary = "Placeholder - file comparison check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `File comparison check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

function checkAdditionalHolidayPay() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Additional Holiday Pay";
  
  try {
    console.log("Running additional holiday pay check...");
    result.status = "pass";
    result.summary = "Placeholder - additional holiday pay check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `Additional holiday pay check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

function checkMultiplePensionSchemes() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Multiple Pension Schemes";
  
  try {
    console.log("Running multiple pension schemes check...");
    result.status = "pass";
    result.summary = "Placeholder - multiple pension schemes check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `Multiple pension schemes check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

function checkSalarySMP() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Salary & SMP";
  
  try {
    console.log("Running salary & SMP check...");
    result.status = "pass";
    result.summary = "Placeholder - salary & SMP check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `Salary & SMP check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

function checkMaternityLeave() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Maternity Leave";
  
  try {
    console.log("Running maternity leave check...");
    result.status = "pass";
    result.summary = "Placeholder - maternity leave check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `Maternity leave check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

function checkAbsenceDaysDeduction() {
  const startTime = Date.now();
  const result = { ...CHECK_RESULT_TEMPLATE };
  result.checkName = "Absence Days Deduction";
  
  try {
    console.log("Running absence days deduction check...");
    result.status = "pass";
    result.summary = "Placeholder - absence days deduction check logic to be implemented";
    result.recordsChecked = 0;
  } catch (error) {
    result.status = "error";
    result.summary = `Absence days deduction check failed: ${error.message}`;
    result.error = error.message;
  }
  
  result.executionTime = Date.now() - startTime;
  return result;
}

/**
 * UTILITY: Get all available checks (CACHE-BASED APPROACH)
 * Returns array of all check functions that read from cache instead of parameters
 */
function getAllPrePayrollChecks() {
  return [
    { name: "Salary Validation", function: checkSalaryValidation },
    { name: "NMW Compliance", function: checkNMWCompliance },
    { name: "Missing/Blank Payments", function: checkMissingBlankPayments },
    { name: "Duplicate Backpays", function: checkDuplicateBackpays },
    { name: "Net Negative Payments", function: checkNetNegativePayments },
    { name: "EV Scheme Payments", function: checkEVSchemePayments },
    { name: "Car Allowance", function: checkCarAllowance },
    { name: "Preview vs Upload", function: checkPreviewVsUpload },
    { name: "Additional Holiday Pay", function: checkAdditionalHolidayPay },
    { name: "Multiple Pension Schemes", function: checkMultiplePensionSchemes },
    { name: "Salary & SMP", function: checkSalarySMP },
    { name: "Maternity Leave", function: checkMaternityLeave },
    { name: "Absence Days Deduction", function: checkAbsenceDaysDeduction }
  ];
}

/**
 * HELPER FUNCTIONS FOR VALIDATION CHECKS
 */

/**
 * Calculate age from date of birth string
 */
function calculateAge(dateOfBirth) {
  if (!dateOfBirth) return null;
  
  try {
    const birthDate = typeof dateOfBirth === 'string' ? new Date(dateOfBirth) : dateOfBirth;
    const today = new Date();
    let age = today.getFullYear() - birthDate.getFullYear();
    const monthDiff = today.getMonth() - birthDate.getMonth();
    
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
      age--;
    }
    
    return age;
  } catch (error) {
    console.warn(`Could not calculate age from: ${dateOfBirth}`);
    return null;
  }
}

/**
 * Calculate string similarity for duplicate detection
 */
function calculateStringSimilarity(str1, str2) {
  if (!str1 || !str2) return 0;
  
  // Simple similarity based on common words
  const words1 = str1.toLowerCase().split(/\s+/);
  const words2 = str2.toLowerCase().split(/\s+/);
  
  const commonWords = words1.filter(word => words2.includes(word));
  const totalWords = new Set([...words1, ...words2]).size;
  
  return totalWords > 0 ? commonWords.length / totalWords : 0;
}

/**
 * Get current NMW rates (implement based on your rate card system)
 */
function getCurrentNMWRates() {
  try {
    // Load from your NMW rate card file
    const rateCardFile = DriveApp.getFileById(CONFIG.NMW_RATE_CARD_FILE_ID);
    const rateData = parseCsvToObjects(rateCardFile);
    
    // Find current rates (adapt this logic to your rate card structure)
    const currentDate = new Date();
    const currentRates = rateData.find(rate => {
      const effectiveDate = new Date(rate.EffectiveDate || rate.effectiveDate);
      return effectiveDate <= currentDate;
    });
    
    if (!currentRates) {
      // Fallback to current UK rates (as of 2024)
      console.warn("Could not load NMW rates from file, using fallback rates");
      return {
        adult: 10.42,
        development: 10.18,
        under18: 7.55,
        apprentice: 7.55
      };
    }
    
    return {
      adult: parseFloat(currentRates.AdultRate || currentRates.adultRate || 10.42),
      development: parseFloat(currentRates.DevelopmentRate || currentRates.developmentRate || 10.18),
      under18: parseFloat(currentRates.Under18Rate || currentRates.under18Rate || 7.55),
      apprentice: parseFloat(currentRates.ApprenticeRate || currentRates.apprenticeRate || 7.55)
    };
    
  } catch (error) {
    console.warn(`Could not load NMW rates: ${error.message}, using fallback rates`);
    return {
      adult: 10.42,
      development: 10.18,
      under18: 7.55,
      apprentice: 7.55
    };
  }
}

/**
 * IMPROVED: Create standardized exception object with unique ID generation
 */
function createStandardException(checkName, employeeNumber, employeeName, issue, severity, details = {}) {
  const exception = {
    checkName: checkName,
    employeeNumber: employeeNumber || "Unknown",
    employeeName: employeeName || "Name Unknown", 
    issue: issue,
    severity: severity || 'warning', // critical | warning | info
    details: details,
    reviewStatus: null,
    reviewedBy: null,
    reviewedAt: null,
    createdAt: new Date().toISOString(),
    exceptionId: null // Will be generated when written to sheet
  };
  
  // Generate unique exception ID
  exception.exceptionId = generateExceptionId(exception);
  
  return exception;
}

/**
 * Generate unique exception ID
 */
function generateExceptionId(exception) {
  const checkPrefix = (exception.checkName || 'unknown').toLowerCase().replace(/[^a-z0-9]/g, '');
  const empNumber = (exception.employeeNumber || 'unknown').toString().replace(/[^a-z0-9]/g, '');
  const timestamp = Date.now();
  const random = Math.floor(Math.random() * 1000);
  
  return `${checkPrefix}_${empNumber}_${timestamp}_${random}`;
}
