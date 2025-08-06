/**
 * Loads the complete payroll calendar with parsed dates.
 * Cached for performance across multiple calls.
 * 
 * @returns {Array<Object>} Array of payroll period objects
 */
function getPayrollCalendar() {
  // Use the generic utility to load JSON
  const rawData = loadJsonFromDrive(CONFIG.PAYROLL_CALENDAR_FILE_ID);
  
  // Parse dates for business use
  return rawData.map(entry => ({
    payDate: new Date(entry.payDate),
    staffType: entry.staffType,
    periodStart: new Date(entry.periodStart),
    periodEnd: entry.periodEnd ? new Date(entry.periodEnd) : null,
    periodName: entry.periodName,
    cutoffDate: entry.cutoffDate ? new Date(entry.cutoffDate) : null,
    bacsRun: entry.bacsRun ? new Date(entry.bacsRun) : null
  }));
}

// === Core Business Functions (Used by Other Modules) ===

/**
 * Gets the payroll period that contains a specific date.
 * 
 * @param {Date|string} targetDate - Date to find period for
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Object|null} Matching period or null
 */
function getPayPeriodForDate(targetDate, staffType) {
  const date = new Date(targetDate);
  const calendar = getPayrollCalendar();
  
  return calendar.find(period =>
    period.staffType === staffType &&
    period.periodStart &&
    period.periodEnd &&
    date >= period.periodStart &&
    date <= period.periodEnd
  ) || null;
}

/**
 * Gets the current payroll period for a staff type.
 * 
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Object|null} Current period or null
 */
function getCurrentPayPeriod(staffType) {
  return getPayPeriodForDate(new Date(), staffType);
}

/**
 * Gets all periods for a specific staff type, sorted chronologically.
 * 
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Array<Object>} All periods for staff type
 */
function getAllPeriodsForStaffType(staffType) {
  return getPayrollCalendar()
    .filter(period => period.staffType === staffType)
    .sort((a, b) => a.periodStart - b.periodStart);
}

/**
 * Gets the previous payroll period relative to a given period.
 * 
 * @param {Object} currentPeriod - Current period object
 * @returns {Object|null} Previous period or null
 */
function getPreviousPayPeriod(currentPeriod) {
  const periods = getAllPeriodsForStaffType(currentPeriod.staffType);
  return getAdjacentRecord(periods, currentPeriod, 'periodName', -1);
}

/**
 * Gets the next payroll period relative to a given period.
 * 
 * @param {Object} currentPeriod - Current period object
 * @returns {Object|null} Next period or null
 */
function getNextPayPeriod(currentPeriod) {
  const periods = getAllPeriodsForStaffType(currentPeriod.staffType);
  return getAdjacentRecord(periods, currentPeriod, 'periodName', 1);
}

// === Validation & Helper Functions ===

/**
 * Validates if a date falls within any defined payroll period.
 * 
 * @param {Date|string} date - Date to validate
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {boolean} True if date is in a defined period
 */
function isDateInPayrollPeriod(date, staffType) {
  return getPayPeriodForDate(date, staffType) !== null;
}

/**
 * Gets all staff types available in the payroll calendar.
 * 
 * @returns {Array<string>} Unique staff types
 */
function getAvailableStaffTypes() {
  const calendar = getPayrollCalendar();
  return [...new Set(calendar.map(period => period.staffType))];
}

/**
 * Checks if a payroll period is currently active.
 * 
 * @param {Object} period - Period object to check
 * @returns {boolean} True if period contains today's date
 */
function isPeriodCurrent(period) {
  const today = new Date();
  return period.periodStart <= today && today <= period.periodEnd;
}

// === Module Integration Helpers ===

/**
 * Gets payroll context for multiple staff types.
 * Used by reporting and dashboard modules.
 * 
 * @param {Array<string>} staffTypes - Array of staff types to get context for
 * @returns {Array<Object>} Context objects for each staff type
 */
function getPayrollContextForStaffTypes(staffTypes = null) {
  const types = staffTypes || getAvailableStaffTypes();
  
  return types.map(staffType => {
    const current = getCurrentPayPeriod(staffType);
    const previous = current ? getPreviousPayPeriod(current) : null;
    const next = current ? getNextPayPeriod(current) : null;
    
    return {
      staffType,
      current: current || null,
      previous: previous || null,
      next: next || null,
      inValidPeriod: current !== null
    };
  });
}

/**
 * Gets upcoming deadlines across all staff types.
 * Used by notification and reminder modules.
 * 
 * @param {number} daysAhead - How many days to look ahead (default: 30)
 * @returns {Array<Object>} Upcoming deadlines with context
 */
function getUpcomingPayrollDeadlines(daysAhead = 30) {
  const today = new Date();
  const futureDate = new Date(today.getTime() + (daysAhead * 24 * 60 * 60 * 1000));
  const calendar = getPayrollCalendar();
  
  const deadlines = [];
  
  calendar.forEach(period => {
    // Check cutoff dates
    if (period.cutoffDate && period.cutoffDate >= today && period.cutoffDate <= futureDate) {
      deadlines.push({
        type: 'cutoff',
        date: period.cutoffDate,
        period: period,
        daysUntil: Math.ceil((period.cutoffDate - today) / (1000 * 60 * 60 * 24))
      });
    }
    
    // Check pay dates
    if (period.payDate && period.payDate >= today && period.payDate <= futureDate) {
      deadlines.push({
        type: 'payDate',
        date: period.payDate,
        period: period,
        daysUntil: Math.ceil((period.payDate - today) / (1000 * 60 * 60 * 24))
      });
    }
  });
  
  return deadlines.sort((a, b) => a.date - b.date);
}

/**
 * Gets all pay periods that fall within a date range.
 * Useful for quarterly/yearly reporting.
 * 
 * @param {Date|string} startDate - Range start date
 * @param {Date|string} endDate - Range end date  
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Array<Object>} Pay periods in the date range
 */
function getPayPeriodsInDateRange(startDate, endDate, staffType) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  
  return getAllPeriodsForStaffType(staffType).filter(period => {
    // Period overlaps with the date range
    return period.periodStart <= end && period.periodEnd >= start;
  });
}

/**
 * Gets all pay periods for a specific calendar year.
 * 
 * @param {number} year - Calendar year (e.g., 2024)
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Array<Object>} Pay periods in the year
 */
function getPayPeriodsForYear(year, staffType) {
  const startDate = new Date(year, 0, 1); // Jan 1
  const endDate = new Date(year, 11, 31); // Dec 31
  return getPayPeriodsInDateRange(startDate, endDate, staffType);
}

/**
 * Validates that employee pay types are consistent with calendar staff types.
 * Useful for data integrity checks.
 * 
 * @returns {Object} Validation results with any inconsistencies
 */
function validateEmployeePayTypeConsistency() {
  const employees = getAllEmployees();
  const calendarStaffTypes = getAvailableStaffTypes();
  
  const issues = [];
  
  employees.forEach(emp => {
    if (emp.IsPayrollEmployee?.toLowerCase() === "yes") {
      let expectedStaffType;
      
      if (emp.PayType === "Salary") {
        expectedStaffType = "Salaried";
      } else if (emp.PayType === "Hourly") {
        expectedStaffType = "Hourly";
      }
      
      if (expectedStaffType && !calendarStaffTypes.includes(expectedStaffType)) {
        issues.push({
          employeeNumber: emp.EmployeeNumber,
          employeeName: emp.FirstName + " " + emp.LastName,
          payType: emp.PayType,
          expectedStaffType: expectedStaffType,
          issue: `No calendar found for staff type "${expectedStaffType}"`
        });
      }
    }
  });
  
  return {
    isValid: issues.length === 0,
    issues: issues,
    summary: `${issues.length} employees with pay type mismatches`
  };
}

/**
 * 📅 Enhanced Payroll Calendar Module
 * NOW INCLUDES: Enhanced maternity leave period matching with cutoff logic
 */

/**
 * Loads the complete payroll calendar with parsed dates.
 * Cached for performance across multiple calls.
 * 
 * @returns {Array<Object>} Array of payroll period objects
 */
function getPayrollCalendar() {
  // Use the generic utility to load JSON
  const rawData = loadJsonFromDrive(CONFIG.PAYROLL_CALENDAR_FILE_ID);
  
  // Parse dates for business use
  return rawData.map(entry => ({
    payDate: new Date(entry.payDate),
    staffType: entry.staffType,
    periodStart: new Date(entry.periodStart),
    periodEnd: entry.periodEnd ? new Date(entry.periodEnd) : null,
    periodName: entry.periodName,
    cutoffDate: entry.cutoffDate ? new Date(entry.cutoffDate) : null,
    bacsRun: entry.bacsRun ? new Date(entry.bacsRun) : null
  }));
}

// === ENHANCED MATERNITY SUPPORT FUNCTIONS ===

/**
 * ENHANCED: Finds the correct payroll period for an SMP start date, considering cutoffs.
 * 
 * For Salaried: Simple date range matching
 * For Hourly: Must be after cutoff date to be paid in that period
 * 
 * @param {Date|string} smpStartDate - SMP start date
 * @param {string} staffType - "Salaried" or "Hourly" 
 * @returns {Object|null} Matching payroll period
 */
function getPayrollPeriodForSMPStart(smpStartDate, staffType) {
  const smpDate = new Date(smpStartDate);
  const calendar = getPayrollCalendar();
  
  const periods = calendar.filter(period => period.staffType === staffType)
                          .sort((a, b) => a.periodStart - b.periodStart);
  
  if (staffType === "Salaried") {
    // Simple date range matching for salaried employees
    return periods.find(period =>
      smpDate >= period.periodStart && smpDate <= period.periodEnd
    ) || null;
  }
  
  if (staffType === "Hourly") {
    // Cutoff-based matching for hourly employees
    // SMP date must be: cutoffDate < smpDate <= periodEnd
    return periods.find(period =>
      period.cutoffDate && 
      smpDate > period.cutoffDate && 
      smpDate <= period.periodEnd
    ) || null;
  }
  
  return null;
}

/**
 * ENHANCED: Gets all payroll periods affected by maternity leave.
 * Uses proper cutoff logic for period assignment.
 * 
 * @param {Date|string} smpStartDate - SMP start date
 * @param {Date|string} maternityEndDate - End of maternity leave
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Array<Object>} Chronological array of affected periods
 */
function getMaternityPayrollPeriods(smpStartDate, maternityEndDate, staffType) {
  const smpDate = new Date(smpStartDate);
  const endDate = new Date(maternityEndDate);
  const calendar = getPayrollCalendar();
  
  const periods = calendar.filter(period => period.staffType === staffType)
                          .sort((a, b) => a.periodStart - b.periodStart);
  
  const affectedPeriods = [];
  
  // Find the first period containing the SMP start date
  const firstPeriod = getPayrollPeriodForSMPStart(smpDate, staffType);
  if (!firstPeriod) {
    console.log('⚠️ No payroll period found for SMP start date: ' + smpDate);
    return [];
  }
  
  // Add all periods from SMP start until maternity end
  let foundFirst = false;
  for (const period of periods) {
    // Start adding periods from the first period we found
    if (!foundFirst) {
      if (period.periodName === firstPeriod.periodName) {
        foundFirst = true;
        affectedPeriods.push(period);
      }
      continue;
    }
    
    // Add subsequent periods until we pass the maternity end date
    if (period.periodStart <= endDate) {
      affectedPeriods.push(period);
    } else {
      break; // We've gone past the maternity leave period
    }
  }
  
  console.log(`Found ${affectedPeriods.length} payroll periods for maternity leave`);
  return affectedPeriods;
}

/**
 * ENHANCED: Validates if an SMP start date can be correctly assigned to a payroll period.
 * Useful for form validation and error checking.
 * 
 * @param {Date|string} smpStartDate - SMP start date to validate
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Object} Validation result with period assignment details
 */
function validateSMPStartDate(smpStartDate, staffType) {
  const smpDate = new Date(smpStartDate);
  
  try {
    const assignedPeriod = getPayrollPeriodForSMPStart(smpDate, staffType);
    
    if (!assignedPeriod) {
      return {
        isValid: false,
        error: `No ${staffType} payroll period found for SMP start date ${smpDate.toDateString()}`,
        suggestion: 'Please check the SMP start date and ensure it falls within a defined payroll period'
      };
    }
    
    // Additional validation for hourly employees
    if (staffType === "Hourly" && assignedPeriod.cutoffDate) {
      const daysSinceCutoff = Math.ceil((smpDate - assignedPeriod.cutoffDate) / (1000 * 60 * 60 * 24));
      
      return {
        isValid: true,
        assignedPeriod: assignedPeriod,
        payrollDetails: {
          periodName: assignedPeriod.periodName,
          payDate: assignedPeriod.payDate,
          daysSinceCutoff: daysSinceCutoff,
          cutoffDate: assignedPeriod.cutoffDate
        },
        message: `SMP will be paid in ${assignedPeriod.periodName} payroll (${daysSinceCutoff} days after cutoff)`
      };
    }
    
    // Validation for salaried employees
    return {
      isValid: true,
      assignedPeriod: assignedPeriod,
      payrollDetails: {
        periodName: assignedPeriod.periodName,
        payDate: assignedPeriod.payDate
      },
      message: `SMP will be paid in ${assignedPeriod.periodName} payroll`
    };
    
  } catch (error) {
    return {
      isValid: false,
      error: `Validation failed: ${error.message}`,
      suggestion: 'Please contact system administrator'
    };
  }
}

// === CORE BUSINESS FUNCTIONS (EXISTING + ENHANCED) ===

/**
 * Gets the payroll period that contains a specific date.
 * 
 * @param {Date|string} targetDate - Date to find period for
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Object|null} Matching period or null
 */
function getPayPeriodForDate(targetDate, staffType) {
  const date = new Date(targetDate);
  const calendar = getPayrollCalendar();
  
  return calendar.find(period =>
    period.staffType === staffType &&
    period.periodStart &&
    period.periodEnd &&
    date >= period.periodStart &&
    date <= period.periodEnd
  ) || null;
}

/**
 * Gets the current payroll period for a staff type.
 * 
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Object|null} Current period or null
 */
function getCurrentPayPeriod(staffType) {
  return getPayPeriodForDate(new Date(), staffType);
}

/**
 * Gets all periods for a specific staff type, sorted chronologically.
 * 
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Array<Object>} All periods for staff type
 */
function getAllPeriodsForStaffType(staffType) {
  return getPayrollCalendar()
    .filter(period => period.staffType === staffType)
    .sort((a, b) => a.periodStart - b.periodStart);
}

/**
 * Gets the previous payroll period relative to a given period.
 * 
 * @param {Object} currentPeriod - Current period object
 * @returns {Object|null} Previous period or null
 */
function getPreviousPayPeriod(currentPeriod) {
  const periods = getAllPeriodsForStaffType(currentPeriod.staffType);
  return getAdjacentRecord(periods, currentPeriod, 'periodName', -1);
}

/**
 * Gets the next payroll period relative to a given period.
 * 
 * @param {Object} currentPeriod - Current period object
 * @returns {Object|null} Next period or null
 */
function getNextPayPeriod(currentPeriod) {
  const periods = getAllPeriodsForStaffType(currentPeriod.staffType);
  return getAdjacentRecord(periods, currentPeriod, 'periodName', 1);
}

// === VALIDATION & HELPER FUNCTIONS ===

/**
 * Validates if a date falls within any defined payroll period.
 * 
 * @param {Date|string} date - Date to validate
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {boolean} True if date is in a defined period
 */
function isDateInPayrollPeriod(date, staffType) {
  return getPayPeriodForDate(date, staffType) !== null;
}

/**
 * Gets all staff types available in the payroll calendar.
 * 
 * @returns {Array<string>} Unique staff types
 */
function getAvailableStaffTypes() {
  const calendar = getPayrollCalendar();
  return [...new Set(calendar.map(period => period.staffType))];
}

/**
 * Checks if a payroll period is currently active.
 * 
 * @param {Object} period - Period object to check
 * @returns {boolean} True if period contains today's date
 */
function isPeriodCurrent(period) {
  const today = new Date();
  return period.periodStart <= today && today <= period.periodEnd;
}

/**
 * ENHANCED: Gets cutoff information for hourly employees.
 * Useful for maternity management and payroll processing.
 * 
 * @param {Date|string} targetDate - Date to find cutoff context for
 * @returns {Object} Cutoff context with relevant periods
 */
function getCutoffContext(targetDate) {
  const date = new Date(targetDate);
  const hourlyPeriods = getAllPeriodsForStaffType("Hourly");
  
  // Find periods around the target date
  const currentPeriod = hourlyPeriods.find(period =>
    date >= period.periodStart && date <= period.periodEnd
  );
  
  const previousPeriod = currentPeriod ? 
    hourlyPeriods.find(period => period.periodEnd < currentPeriod.periodStart) : null;
  
  const nextPeriod = currentPeriod ?
    hourlyPeriods.find(period => period.periodStart > currentPeriod.periodEnd) : null;
  
  return {
    targetDate: date,
    currentPeriod: currentPeriod,
    previousPeriod: previousPeriod,
    nextPeriod: nextPeriod,
    isAfterCutoff: currentPeriod && currentPeriod.cutoffDate ? 
      date > currentPeriod.cutoffDate : false,
    daysUntilNextCutoff: nextPeriod && nextPeriod.cutoffDate ?
      Math.ceil((nextPeriod.cutoffDate - date) / (1000 * 60 * 60 * 24)) : null
  };
}

// === MODULE INTEGRATION HELPERS ===

/**
 * Gets payroll context for multiple staff types.
 * Used by reporting and dashboard modules.
 * 
 * @param {Array<string>} staffTypes - Array of staff types to get context for
 * @returns {Array<Object>} Context objects for each staff type
 */
function getPayrollContextForStaffTypes(staffTypes = null) {
  const types = staffTypes || getAvailableStaffTypes();
  
  return types.map(staffType => {
    const current = getCurrentPayPeriod(staffType);
    const previous = current ? getPreviousPayPeriod(current) : null;
    const next = current ? getNextPayPeriod(current) : null;
    
    return {
      staffType,
      current: current || null,
      previous: previous || null,
      next: next || null,
      inValidPeriod: current !== null
    };
  });
}

/**
 * Gets upcoming deadlines across all staff types.
 * Used by notification and reminder modules.
 * 
 * @param {number} daysAhead - How many days to look ahead (default: 30)
 * @returns {Array<Object>} Upcoming deadlines with context
 */
function getUpcomingPayrollDeadlines(daysAhead = 30) {
  const today = new Date();
  const futureDate = new Date(today.getTime() + (daysAhead * 24 * 60 * 60 * 1000));
  const calendar = getPayrollCalendar();
  
  const deadlines = [];
  
  calendar.forEach(period => {
    // Check cutoff dates
    if (period.cutoffDate && period.cutoffDate >= today && period.cutoffDate <= futureDate) {
      deadlines.push({
        type: 'cutoff',
        date: period.cutoffDate,
        period: period,
        daysUntil: Math.ceil((period.cutoffDate - today) / (1000 * 60 * 60 * 24))
      });
    }
    
    // Check pay dates
    if (period.payDate && period.payDate >= today && period.payDate <= futureDate) {
      deadlines.push({
        type: 'payDate',
        date: period.payDate,
        period: period,
        daysUntil: Math.ceil((period.payDate - today) / (1000 * 60 * 60 * 24))
      });
    }
  });
  
  return deadlines.sort((a, b) => a.date - b.date);
}

/**
 * Gets all pay periods that fall within a date range.
 * Useful for quarterly/yearly reporting.
 * 
 * @param {Date|string} startDate - Range start date
 * @param {Date|string} endDate - Range end date  
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Array<Object>} Pay periods in the date range
 */
function getPayPeriodsInDateRange(startDate, endDate, staffType) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  
  return getAllPeriodsForStaffType(staffType).filter(period => {
    // Period overlaps with the date range
    return period.periodStart <= end && period.periodEnd >= start;
  });
}

/**
 * Gets all pay periods for a specific calendar year.
 * 
 * @param {number} year - Calendar year (e.g., 2024)
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Array<Object>} Pay periods in the year
 */
function getPayPeriodsForYear(year, staffType) {
  const startDate = new Date(year, 0, 1); // Jan 1
  const endDate = new Date(year, 11, 31); // Dec 31
  return getPayPeriodsInDateRange(startDate, endDate, staffType);
}

/**
 * Validates that employee pay types are consistent with calendar staff types.
 * Useful for data integrity checks.
 * 
 * @returns {Object} Validation results with any inconsistencies
 */
function validateEmployeePayTypeConsistency() {
  const employees = getAllEmployees();
  const calendarStaffTypes = getAvailableStaffTypes();
  
  const issues = [];
  
  employees.forEach(emp => {
    if (emp.IsPayrollEmployee?.toLowerCase() === "yes") {
      let expectedStaffType;
      
      if (emp.PayType === "Salary") {
        expectedStaffType = "Salaried";
      } else if (emp.PayType === "Hourly") {
        expectedStaffType = "Hourly";
      }
      
      if (expectedStaffType && !calendarStaffTypes.includes(expectedStaffType)) {
        issues.push({
          employeeNumber: emp.EmployeeNumber,
          employeeName: emp.FirstName + " " + emp.LastName,
          payType: emp.PayType,
          expectedStaffType: expectedStaffType,
          issue: `No calendar found for staff type "${expectedStaffType}"`
        });
      }
    }
  });
  
  return {
    isValid: issues.length === 0,
    issues: issues,
    summary: `${issues.length} employees with pay type mismatches`
  };
}

// === ENHANCED TESTING & DEBUGGING ===

/**
 * ENHANCED: Tests maternity period assignment logic.
 * Useful for debugging cutoff-based period matching.
 * 
 * @param {Date|string} testSMPStart - Test SMP start date
 * @param {string} staffType - "Salaried" or "Hourly"
 * @returns {Object} Test results with detailed period assignment
 */
function testMaternityPeriodAssignment(testSMPStart, staffType) {
  const smpDate = new Date(testSMPStart);
  
  console.log(`🧪 Testing maternity period assignment for ${staffType} employee`);
  console.log(`SMP Start Date: ${smpDate.toDateString()}`);
  
  try {
    // Test period assignment
    const assignedPeriod = getPayrollPeriodForSMPStart(smpDate, staffType);
    
    if (!assignedPeriod) {
      return {
        success: false,
        error: `No ${staffType} period found for SMP start date`,
        suggestion: 'Check if the date falls within a defined payroll period'
      };
    }
    
    // Test maternity period range
    const testEndDate = new Date(smpDate);
    testEndDate.setDate(testEndDate.getDate() + (52 * 7)); // 52 weeks later
    
    const maternityPeriods = getMaternityPayrollPeriods(smpDate, testEndDate, staffType);
    
    // Validation details
    const validation = validateSMPStartDate(smpDate, staffType);
    
    return {
      success: true,
      testDate: smpDate,
      staffType: staffType,
      assignedPeriod: {
        periodName: assignedPeriod.periodName,
        periodStart: assignedPeriod.periodStart,
        periodEnd: assignedPeriod.periodEnd,
        payDate: assignedPeriod.payDate,
        cutoffDate: assignedPeriod.cutoffDate
      },
      maternityPeriodCount: maternityPeriods.length,
      validation: validation,
      firstThreePeriods: maternityPeriods.slice(0, 3).map(p => ({
        periodName: p.periodName,
        periodStart: p.periodStart,
        periodEnd: p.periodEnd,
        payDate: p.payDate
      }))
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      testDate: smpDate,
      staffType: staffType
    };
  }
}

/**
 * ENHANCED: Comprehensive system health check for maternity support.
 * 
 * @returns {Object} Health check results with recommendations
 */
function checkMaternityPayrollHealth() {
  const healthResults = {
    overall: 'healthy',
    components: {},
    recommendations: []
  };
  
  try {
    // Check calendar data availability
    const calendar = getPayrollCalendar();
    const staffTypes = getAvailableStaffTypes();
    
    healthResults.components.calendarData = {
      status: calendar.length > 0 ? 'healthy' : 'error',
      message: `${calendar.length} payroll periods loaded`,
      data: { totalPeriods: calendar.length, staffTypes: staffTypes }
    };
    
    // Check cutoff data for hourly employees
    const hourlyPeriods = calendar.filter(p => p.staffType === 'Hourly');
    const hourlyWithCutoffs = hourlyPeriods.filter(p => p.cutoffDate);
    
    healthResults.components.cutoffData = {
      status: hourlyWithCutoffs.length === hourlyPeriods.length ? 'healthy' : 'warning',
      message: `${hourlyWithCutoffs.length}/${hourlyPeriods.length} hourly periods have cutoff dates`,
      data: { hourlyPeriods: hourlyPeriods.length, withCutoffs: hourlyWithCutoffs.length }
    };
    
    // Test SMP assignment for both staff types
    const testDate = new Date(); // Today
    const salariedTest = getPayrollPeriodForSMPStart(testDate, 'Salaried');
    const hourlyTest = getPayrollPeriodForSMPStart(testDate, 'Hourly');
    
    healthResults.components.smpAssignment = {
      status: (salariedTest && hourlyTest) ? 'healthy' : 'warning',
      message: `Period assignment working for ${salariedTest ? 'Salaried' : 'None'} and ${hourlyTest ? 'Hourly' : 'None'}`,
      data: { 
        salariedWorking: !!salariedTest, 
        hourlyWorking: !!hourlyTest,
        testDate: testDate
      }
    };
    
    // Check for missing cutoff dates
    if (hourlyWithCutoffs.length < hourlyPeriods.length) {
      healthResults.recommendations.push('Some hourly payroll periods are missing cutoff dates - this may affect maternity CMP calculations');
    }
    
    // Check for future periods
    const today = new Date();
    const futurePeriods = calendar.filter(p => p.periodEnd > today);
    
    healthResults.components.futurePeriods = {
      status: futurePeriods.length > 10 ? 'healthy' : 'warning',
      message: `${futurePeriods.length} future periods available`,
      data: { futurePeriods: futurePeriods.length }
    };
    
    if (futurePeriods.length < 10) {
      healthResults.recommendations.push('Consider adding more future payroll periods to support long-term maternity leave planning');
    }
    
    // Overall health determination
    const componentStatuses = Object.values(healthResults.components).map(c => c.status);
    if (componentStatuses.includes('error')) {
      healthResults.overall = 'error';
    } else if (componentStatuses.includes('warning')) {
      healthResults.overall = 'warning';
    }
    
  } catch (error) {
    healthResults.overall = 'error';
    healthResults.components.systemError = {
      status: 'error',
      message: `System error: ${error.message}`,
      data: { error: error.toString() }
    };
  }
  
  return healthResults;
}
