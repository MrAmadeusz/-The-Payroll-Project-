/**
 * 💼 06_MaternityTopUp.gs
 * Company Maternity Pay (CMP) Top-Up Calculation Logic
 * Calculates CMP amounts based on contracted vs average earnings
 */

/**
 * Calculates contracted weekly earnings for an employee
 * @param {Object} employee - Employee object from directory
 * @returns {number} Contracted weekly earnings
 */
function calculateContractedWeeklyEarnings(employee) {
  if (!employee) {
    throw new Error('Employee data is required for CMP calculation');
  }
  
  console.log(`Calculating contracted earnings for ${employee.FullName || employee.Firstnames + ' ' + employee.Surname}`);
  
  if (employee.PayType === 'Salary') {
    // Salaried: Annual salary ÷ 52
    const annualSalary = parseFloat(employee.Salary) || 0;
    const weeklyEarnings = annualSalary / 52;
    
    console.log(`  Salaried: £${annualSalary}/year = £${weeklyEarnings.toFixed(2)}/week`);
    return Math.round(weeklyEarnings * 100) / 100;
    
  } else if (employee.PayType === 'Hourly') {
    // Hourly: Rate × Contracted Hours
    const hourlyRate = parseFloat(employee.HourlyShiftAmount) || 0;
    const contractedHours = parseFloat(employee.ContractHours) || 0;
    const weeklyEarnings = hourlyRate * contractedHours;
    
    console.log(`  Hourly: £${hourlyRate}/hour × ${contractedHours} hours = £${weeklyEarnings.toFixed(2)}/week`);
    
    // Handle zero hours contracts
    if (contractedHours === 0) {
      console.log('  ⚠️ Zero hours contract - weekly earnings = £0');
    }
    
    return Math.round(weeklyEarnings * 100) / 100;
    
  } else {
    console.log(`  ⚠️ Unknown pay type: ${employee.PayType}`);
    return 0;
  }
}

/**
 * Calculates target weekly amount (higher of average or contracted earnings)
 * @param {number} averageWeeklyEarnings - Manual input from case form
 * @param {number} contractedWeeklyEarnings - Calculated from employee data
 * @returns {number} Target weekly amount for CMP calculation
 */
function calculateTargetWeeklyAmount(averageWeeklyEarnings, contractedWeeklyEarnings) {
  const average = parseFloat(averageWeeklyEarnings) || 0;
  const contracted = parseFloat(contractedWeeklyEarnings) || 0;
  const target = Math.max(average, contracted);
  
  console.log(`Target weekly calculation:`);
  console.log(`  Average weekly earnings: £${average.toFixed(2)}`);
  console.log(`  Contracted weekly earnings: £${contracted.toFixed(2)}`);
  console.log(`  Target (higher amount): £${target.toFixed(2)}`);
  
  return Math.round(target * 100) / 100;
}

/**
 * Calculates the number of weeks in a period based on start and end dates
 * @param {Date} periodStart - Period start date
 * @param {Date} periodEnd - Period end date
 * @returns {number} Number of weeks (rounded to nearest 0.1)
 */
function calculateWeeksInPeriod(periodStart, periodEnd) {
  const startDate = new Date(periodStart);
  const endDate = new Date(periodEnd);
  const diffTime = Math.abs(endDate - startDate);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 to include end date
  const weeks = diffDays / 7;
  
  console.log(`  Period ${startDate.toDateString()} to ${endDate.toDateString()}: ${diffDays} days = ${weeks.toFixed(1)} weeks`);
  
  return Math.round(weeks * 10) / 10; // Round to 1 decimal place
}

/**
 * Calculates CMP amount for a specific period
 * @param {Object} period - Period object with start/end dates and SMP amount
 * @param {number} targetWeeklyAmount - Target weekly amount
 * @param {number} weeksInPeriod - Number of weeks in this period
 * @returns {number} CMP amount for this period
 */
function calculateCMPForPeriod(period, targetWeeklyAmount, weeksInPeriod) {
  const targetAmount = targetWeeklyAmount * weeksInPeriod;
  const smpAmount = parseFloat(period.smpAmount) || 0;
  const cmpAmount = Math.max(0, targetAmount - smpAmount); // CMP can't be negative
  
  console.log(`  Period ${period.periodNumber}:`);
  console.log(`    Target: £${targetWeeklyAmount.toFixed(2)}/week × ${weeksInPeriod} weeks = £${targetAmount.toFixed(2)}`);
  console.log(`    SMP: £${smpAmount.toFixed(2)}`);
  console.log(`    CMP: £${targetAmount.toFixed(2)} - £${smpAmount.toFixed(2)} = £${cmpAmount.toFixed(2)}`);
  
  return Math.round(cmpAmount * 100) / 100;
}

/**
 * Main function to calculate and populate CMP amounts for a maternity case
 * @param {Object} maternityCase - Maternity case object with periods
 * @param {number} averageWeeklyEarnings - Manual input from case form
 * @param {number} cmpWeeks - Number of weeks to provide CMP (default 8)
 * @returns {Object} Updated maternity case with CMP amounts populated
 */
function calculateAndPopulateCMP(maternityCase, averageWeeklyEarnings, cmpWeeks = 8) {
  console.log(`\n=== CALCULATING CMP FOR CASE ${maternityCase.caseId} ===`);
  console.log(`Employee: ${maternityCase.employeeName}`);
  console.log(`CMP Duration: ${cmpWeeks} weeks`);
  
  try {
    // 1. Get employee data
    const employee = getEmployeeById(maternityCase.employeeId);
    if (!employee) {
      throw new Error(`Employee ${maternityCase.employeeId} not found for CMP calculation`);
    }
    
    // 2. Calculate contracted weekly earnings
    const contractedWeeklyEarnings = calculateContractedWeeklyEarnings(employee);
    
    // 3. Calculate target weekly amount
    const targetWeeklyAmount = calculateTargetWeeklyAmount(averageWeeklyEarnings, contractedWeeklyEarnings);
    
    // 4. Store calculations at case level
    maternityCase.averageWeeklyEarnings = parseFloat(averageWeeklyEarnings) || 0;
    maternityCase.contractedWeeklyEarnings = contractedWeeklyEarnings;
    maternityCase.targetWeeklyAmount = targetWeeklyAmount;
    maternityCase.cmpWeeks = parseInt(cmpWeeks) || 8;
    
    // 5. Calculate CMP for each period within the CMP duration
    let totalCMPCalculated = 0;
    let weeksProcessed = 0;
    
    console.log(`\nProcessing periods for CMP calculation:`);
    
    // Sort periods by period number to ensure correct order
    const sortedPeriods = [...maternityCase.periods].sort((a, b) => a.periodNumber - b.periodNumber);
    
    sortedPeriods.forEach(period => {
      // Calculate weeks in this period
      const weeksInPeriod = calculateWeeksInPeriod(period.periodStart, period.periodEnd);
      
      // Check if we're still within CMP weeks
      if (weeksProcessed < cmpWeeks) {
        // Calculate how many weeks of this period are eligible for CMP
        const eligibleWeeks = Math.min(weeksInPeriod, cmpWeeks - weeksProcessed);
        
        if (eligibleWeeks > 0) {
          // Calculate CMP for eligible weeks only
          const adjustedTargetAmount = targetWeeklyAmount * eligibleWeeks;
          const smpAmount = parseFloat(period.smpAmount) || 0;
          
          // If period extends beyond CMP weeks, prorate the SMP
          let adjustedSMPAmount = smpAmount;
          if (eligibleWeeks < weeksInPeriod) {
            adjustedSMPAmount = smpAmount * (eligibleWeeks / weeksInPeriod);
            console.log(`    Prorating SMP: £${smpAmount.toFixed(2)} × (${eligibleWeeks}/${weeksInPeriod}) = £${adjustedSMPAmount.toFixed(2)}`);
          }
          
          const cmpAmount = Math.max(0, adjustedTargetAmount - adjustedSMPAmount);
          
          period.companyAmount = cmpAmount;
          period.cmpNotes = `CMP top-up: ${eligibleWeeks} weeks @ £${targetWeeklyAmount.toFixed(2)}/week`;
          
          totalCMPCalculated += cmpAmount;
          
          console.log(`  ✓ Period ${period.periodNumber}: £${cmpAmount.toFixed(2)} CMP (${eligibleWeeks} weeks)`);
        } else {
          period.companyAmount = 0;
          console.log(`  - Period ${period.periodNumber}: £0 CMP (no eligible weeks)`);
        }
        
        weeksProcessed += weeksInPeriod;
      } else {
        // Beyond CMP weeks - no company top-up
        period.companyAmount = 0;
        period.cmpNotes = 'Beyond CMP entitlement period';
        console.log(`  - Period ${period.periodNumber}: £0 CMP (beyond ${cmpWeeks} week limit)`);
      }
    });
    
    // 6. Update case total CMP
    maternityCase.totalCMP = Math.round(totalCMPCalculated * 100) / 100;
    
    console.log(`\n=== CMP CALCULATION COMPLETE ===`);
    console.log(`Total CMP calculated: £${maternityCase.totalCMP.toFixed(2)}`);
    console.log(`Weeks processed: ${Math.min(weeksProcessed, cmpWeeks)} of ${cmpWeeks} eligible weeks`);
    
    return maternityCase;
    
  } catch (error) {
    console.error(`CMP calculation failed for case ${maternityCase.caseId}: ${error.message}`);
    
    // Set defaults on error
    maternityCase.totalCMP = 0;
    maternityCase.averageWeeklyEarnings = parseFloat(averageWeeklyEarnings) || 0;
    maternityCase.contractedWeeklyEarnings = 0;
    maternityCase.targetWeeklyAmount = 0;
    maternityCase.cmpWeeks = parseInt(cmpWeeks) || 8;
    
    // Zero out all period CMP amounts
    maternityCase.periods.forEach(period => {
      period.companyAmount = 0;
      period.cmpNotes = `CMP calculation error: ${error.message}`;
    });
    
    throw new Error(`CMP calculation failed: ${error.message}`);
  }
}

/**
 * Recalculates CMP amounts when SMP amounts are updated
 * @param {string} caseId - Maternity case ID
 * @returns {Object} Updated case with recalculated CMP
 */
function recalculateCMPAmounts(caseId) {
  console.log(`\n=== RECALCULATING CMP FOR CASE ${caseId} ===`);
  
  const maternityCase = getMaternityCase(caseId);
  if (!maternityCase) {
    throw new Error(`Case ${caseId} not found for CMP recalculation`);
  }
  
  // Use stored values from original calculation
  const averageWeeklyEarnings = maternityCase.averageWeeklyEarnings || 0;
  const cmpWeeks = maternityCase.cmpWeeks || 8;
  
  if (!averageWeeklyEarnings) {
    throw new Error('Average weekly earnings not found - cannot recalculate CMP');
  }
  
  // Recalculate with current SMP amounts
  const updatedCase = calculateAndPopulateCMP(maternityCase, averageWeeklyEarnings, cmpWeeks);
  
  // Save the updated case
  const allCases = getAllMaternityCases();
  const caseIndex = allCases.findIndex(c => c.caseId === caseId);
  if (caseIndex !== -1) {
    allCases[caseIndex] = updatedCase;
    const saveResult = saveMaternityCases(allCases);
    if (!saveResult.success) {
      throw new Error('Failed to save recalculated CMP amounts');
    }
  }
  
  console.log('✅ CMP recalculation complete');
  return updatedCase;
}

/**
 * Gets CMP summary information for a case
 * @param {string} caseId - Maternity case ID
 * @returns {Object} CMP summary with totals and breakdown
 */
function getCMPSummary(caseId) {
  const maternityCase = getMaternityCase(caseId);
  if (!maternityCase) {
    return null;
  }
  
  const cmpPeriods = maternityCase.periods.filter(p => (p.companyAmount || 0) > 0);
  const totalCMP = maternityCase.totalCMP || 0;
  const weeksEntitled = maternityCase.cmpWeeks || 8;
  const targetWeekly = maternityCase.targetWeeklyAmount || 0;
  
  return {
    caseId: caseId,
    employeeName: maternityCase.employeeName,
    totalCMP: totalCMP,
    cmpWeeks: weeksEntitled,
    targetWeeklyAmount: targetWeekly,
    averageWeeklyEarnings: maternityCase.averageWeeklyEarnings || 0,
    contractedWeeklyEarnings: maternityCase.contractedWeeklyEarnings || 0,
    periodsWithCMP: cmpPeriods.length,
    totalPeriods: maternityCase.periods.length,
    isCalculated: totalCMP > 0 || (targetWeekly > 0)
  };
}

// === INTEGRATION FUNCTIONS ===

/**
 * Integration function for case creation process
 * Call this from createMaternityCase() after periods are generated
 * @param {Object} maternityCase - Newly created maternity case
 * @param {Object} formData - Form data including averageWeeklyEarnings and cmpWeeks
 * @returns {Object} Updated case with CMP calculated
 */
function integrateCMPIntoNewCase(maternityCase, formData) {
  console.log(`\n🔗 INTEGRATING CMP INTO NEW CASE ${maternityCase.caseId}`);
  
  try {
    // Extract CMP parameters from form data
    const averageWeeklyEarnings = parseFloat(formData.averageWeeklyEarnings) || 0;
    const cmpWeeks = parseInt(formData.cmpWeeks) || 8;
    
    console.log(`Form data: Average £${averageWeeklyEarnings}/week, ${cmpWeeks} weeks CMP`);
    
    if (averageWeeklyEarnings <= 0) {
      console.log('⚠️ No average weekly earnings provided - skipping CMP calculation');
      maternityCase.totalCMP = 0;
      return maternityCase;
    }
    
    // Calculate and populate CMP amounts
    return calculateAndPopulateCMP(maternityCase, averageWeeklyEarnings, cmpWeeks);
    
  } catch (error) {
    console.error(`CMP integration failed: ${error.message}`);
    
    // Ensure case has CMP defaults
    maternityCase.totalCMP = 0;
    maternityCase.averageWeeklyEarnings = parseFloat(formData.averageWeeklyEarnings) || 0;
    maternityCase.cmpWeeks = parseInt(formData.cmpWeeks) || 8;
    
    // Don't fail case creation - just log the error
    console.log('⚠️ CMP calculation failed but case creation will continue');
    return maternityCase;
  }
}

/**
 * Test function to validate CMP calculation logic
 * @param {string} employeeId - Test employee ID
 * @param {number} averageWeeklyEarnings - Test average earnings
 * @returns {Object} Test results
 */
function testCMPCalculation(employeeId = null, averageWeeklyEarnings = 300) {
  console.log('🧪 TESTING CMP CALCULATION LOGIC');
  
  try {
    // Use first active employee if none specified
    if (!employeeId) {
      const activeEmployees = getActivePayrollEmployees();
      if (activeEmployees.length === 0) {
        throw new Error('No active employees found for testing');
      }
      employeeId = activeEmployees[0].EmployeeNumber;
    }
    
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      throw new Error(`Employee ${employeeId} not found`);
    }
    
    console.log(`Testing with employee: ${employee.FullName || employee.Firstnames + ' ' + employee.Surname}`);
    console.log(`Pay type: ${employee.PayType}`);
    
    // Test contracted earnings calculation
    const contractedWeekly = calculateContractedWeeklyEarnings(employee);
    const targetWeekly = calculateTargetWeeklyAmount(averageWeeklyEarnings, contractedWeekly);
    
    // Create test periods
    const testPeriods = [
      { periodNumber: 1, periodStart: new Date('2025-06-01'), periodEnd: new Date('2025-06-28'), smpAmount: 500 },
      { periodNumber: 2, periodStart: new Date('2025-06-29'), periodEnd: new Date('2025-07-26'), smpAmount: 600 },
      { periodNumber: 3, periodStart: new Date('2025-07-27'), periodEnd: new Date('2025-08-30'), smpAmount: 550 }
    ];
    
    console.log('\nTest period calculations:');
    testPeriods.forEach(period => {
      const weeks = calculateWeeksInPeriod(period.periodStart, period.periodEnd);
      const cmp = calculateCMPForPeriod(period, targetWeekly, weeks);
      console.log(`Period ${period.periodNumber}: ${weeks} weeks, SMP £${period.smpAmount}, CMP £${cmp}`);
    });
    
    return {
      success: true,
      employee: employee,
      contractedWeekly: contractedWeekly,
      averageWeekly: averageWeeklyEarnings,
      targetWeekly: targetWeekly,
      testPeriods: testPeriods
    };
    
  } catch (error) {
    console.error(`CMP test failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}
