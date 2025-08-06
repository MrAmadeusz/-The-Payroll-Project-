/**
 * 📱 06_maternityUI.gs
 * User Interface Functions for Maternity Management
 * ENHANCED: Better validation, payroll integration, and CMP calculations
 */

// === ENHANCED FORM SUBMISSION HANDLERS ===

/**
 * ENHANCED: Server-side function to create a new maternity case from form data.
 * NOW INCLUDES: SMP start date validation and payroll period checking
 * @param {Object} formData - Form data from the UI
 * @returns {Object} Result object with success/error information
 */
function submitNewMaternityCase(formData) {
  return safeExecute(() => {
    // Enhanced validation with payroll period checking
    const validationResult = validateMaternityFormData(formData);
    if (!validationResult.isValid) {
      throw new Error(validationResult.errors.join('; '));
    }
    
    // Additional payroll-specific validation
    const employee = getEmployeeById(formData.employeeId);
    if (!employee) {
      throw new Error('Employee not found in directory');
    }
    
    const staffType = employee.PayType === 'Salary' ? 'Salaried' : 'Hourly';
    const smpValidation = validateSMPStartDate(formData.smpStartDate, staffType);
    
    if (!smpValidation.isValid) {
      throw new Error(`SMP start date validation failed: ${smpValidation.error}`);
    }
    
    // Create the case with enhanced data
    const caseDetails = {
      ...formData,
      babyDueDate: formData.babyDueDate,
      smpStartDate: formData.smpStartDate,
      totalSMP: parseFloat(formData.totalSMP) || 0,
      // Add payroll validation results for reference
      payrollValidation: smpValidation
    };
    
    const newCase = createMaternityCase(caseDetails);
    
    return {
      success: true,
      message: `Maternity case created successfully for ${newCase.employeeName}. ${smpValidation.message}`,
      caseId: newCase.caseId,
      case: newCase,
      payrollInfo: smpValidation.payrollDetails
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      message: error.message,
      error: error.toString()
    })
  }).result;
}

/**
 * ENHANCED: Validates maternity form data with comprehensive checks.
 * @param {Object} formData - Form data to validate
 * @returns {Object} Validation result with detailed errors
 */
function validateMaternityFormData(formData) {
  const errors = [];
  const warnings = [];
  
  // Required field validation
  if (!formData.employeeId) errors.push('Employee selection is required');
  if (!formData.maternityStartDate) errors.push('Maternity start date is required');
  if (!formData.expectedReturnDate) errors.push('Expected return date is required');
  if (!formData.babyDueDate) errors.push('Baby due date is required');
  if (!formData.smpStartDate) errors.push('SMP start date is required');
  if (!formData.totalSMP || parseFloat(formData.totalSMP) <= 0) {
    errors.push('Total SMP amount must be greater than £0.00');
  }
  if (!formData.averageWeeklyEarnings || parseFloat(formData.averageWeeklyEarnings) <= 0) {
    errors.push('Average weekly earnings must be greater than £0.00');
  }
  
  // Date logic validation
  if (formData.maternityStartDate && formData.expectedReturnDate) {
    const startDate = new Date(formData.maternityStartDate);
    const returnDate = new Date(formData.expectedReturnDate);
    
    if (startDate >= returnDate) {
      errors.push('Expected return date must be after maternity start date');
    }
    
    // Check for reasonable maternity leave duration (between 2 weeks and 18 months)
    const duration = (returnDate - startDate) / (1000 * 60 * 60 * 24 * 7); // weeks
    if (duration < 2) {
      errors.push('Maternity leave duration seems too short (minimum 2 weeks)');
    } else if (duration > 78) { // 18 months
      warnings.push('Maternity leave duration is longer than 18 months');
    }
  }
  
  if (formData.smpStartDate && formData.maternityStartDate) {
    const smpDate = new Date(formData.smpStartDate);
    const maternityDate = new Date(formData.maternityStartDate);
    
    // SMP usually starts on or around maternity start date
    const daysDifference = Math.abs((smpDate - maternityDate) / (1000 * 60 * 60 * 24));
    if (daysDifference > 14) {
      warnings.push('SMP start date is more than 2 weeks from maternity start date');
    }
  }
  
  // CMP validation
  if (formData.cmpWeeks) {
    const cmpWeeks = parseInt(formData.cmpWeeks);
    if (cmpWeeks < 0 || cmpWeeks > 52) {
      errors.push('CMP weeks must be between 0 and 52');
    }
  }
  
  // Monthly SMP validation
  if (formData.monthlySMP && typeof formData.monthlySMP === 'object') {
    const monthlyTotal = Object.values(formData.monthlySMP)
      .reduce((sum, amount) => sum + (parseFloat(amount) || 0), 0);
    
    const expectedTotal = parseFloat(formData.totalSMP) || 0;
    const difference = Math.abs(monthlyTotal - expectedTotal);
    
    if (difference > 0.01) { // Allow for small rounding differences
      errors.push(`Monthly SMP amounts (£${monthlyTotal.toFixed(2)}) don't match total SMP (£${expectedTotal.toFixed(2)})`);
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings,
    hasWarnings: warnings.length > 0
  };
}

/**
 * ENHANCED: Server-side function to update period amounts.
 * NOW INCLUDES: Better CMP recalculation and validation
 * @param {string} caseId - Maternity case ID
 * @param {string} periodId - Period ID
 * @param {Object} amounts - Amount data
 * @returns {Object} Result object
 */
function submitPeriodAmounts(caseId, periodId, amounts) {
  return safeExecute(() => {
    // Validate amounts before processing
    const amountValidation = validatePeriodAmounts(amounts);
    if (!amountValidation.isValid) {
      throw new Error(amountValidation.errors.join('; '));
    }
    
    // Map UI terminology to internal field names
    const periodAmounts = {
      smpAmount: parseFloat(amounts.smpAmount) || 0,
      companyAmount: parseFloat(amounts.cmpAmount) || 0, // CMP maps to companyAmount
      holidayAccrued: parseFloat(amounts.holidayAccrued) || 0,
      smpNotes: amounts.smpNotes || '',
      companyNotes: amounts.cmpNotes || '', // CMP notes
      holidayNotes: amounts.holidayNotes || ''
    };
    
    const updatedPeriod = updatePeriodAmounts(caseId, periodId, periodAmounts);
    const summary = calculateMaternityDisplayData(caseId);
    
    return {
      success: true,
      message: 'Period amounts updated successfully. CMP has been recalculated automatically.',
      period: updatedPeriod,
      summary: summary,
      warnings: amountValidation.warnings
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      message: error.message,
      error: error.toString()
    })
  }).result;
}

/**
 * Validates period amount data.
 * @param {Object} amounts - Amount data to validate
 * @returns {Object} Validation result
 */
function validatePeriodAmounts(amounts) {
  const errors = [];
  const warnings = [];
  
  // Validate numeric amounts
  if (amounts.smpAmount !== undefined) {
    const smpAmount = parseFloat(amounts.smpAmount);
    if (isNaN(smpAmount) || smpAmount < 0) {
      errors.push('SMP amount must be a positive number');
    } else if (smpAmount > 10000) {
      warnings.push('SMP amount seems unusually high (over £10,000)');
    }
  }
  
  if (amounts.cmpAmount !== undefined) {
    const cmpAmount = parseFloat(amounts.cmpAmount);
    if (isNaN(cmpAmount) || cmpAmount < 0) {
      errors.push('CMP amount must be a positive number');
    } else if (cmpAmount > 5000) {
      warnings.push('CMP amount seems unusually high (over £5,000)');
    }
  }
  
  if (amounts.holidayAccrued !== undefined) {
    const holidayDays = parseFloat(amounts.holidayAccrued);
    if (isNaN(holidayDays) || holidayDays < 0) {
      errors.push('Holiday accrued must be a positive number');
    } else if (holidayDays > 30) {
      warnings.push('Holiday accrued seems high (over 30 days per period)');
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings
  };
}

/**
 * Server-side function to update case status or other metadata.
 * @param {string} caseId - Maternity case ID
 * @param {Object} updates - Fields to update
 * @returns {Object} Result object
 */
function submitCaseUpdates(caseId, updates) {
  return safeExecute(() => {
    const updatedCase = updateMaternityCase(caseId, updates);
    const summary = calculateMaternityDisplayData(caseId);
    
    return {
      success: true,
      message: 'Case updated successfully',
      case: updatedCase,
      summary: summary
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      message: error.message,
      error: error.toString()
    })
  }).result;
}

// === ENHANCED DATA PREPARATION FUNCTIONS ===

/**
 * ENHANCED: Server-side function to get dashboard data.
 * NOW INCLUDES: Better error handling and payroll context
 * @returns {Object} Dashboard data with summaries
 */
function getMaternityDashboardData() {
  return safeExecute(() => {
    const allCases = getAllMaternityCases();
    const activeCases = getActiveMaternityCases(); // All non-archived cases
    
    // Get display data for active cases with enhanced error handling
    const activeSummaries = activeCases.map(c => {
      try {
        const summary = calculateMaternityDisplayData(c.caseId);
        
        // Enhanced status categorization with payroll context
        const today = new Date();
        const maternityStart = new Date(c.maternityStartDate);
        const actualReturn = c.actualReturnDate ? new Date(c.actualReturnDate) : null;
        const expectedReturn = new Date(c.expectedReturnDate);
        
        let statusCategory = 'onLeave';
        let statusDescription = 'On Maternity Leave';
        
        if (maternityStart > today) {
          statusCategory = 'upcoming';
          const daysUntilStart = Math.ceil((maternityStart - today) / (1000 * 60 * 60 * 24));
          statusDescription = `Starts in ${daysUntilStart} days`;
        } else if (actualReturn && actualReturn <= today) {
          statusCategory = 'returned';
          statusDescription = 'Returned';
        } else if (today > expectedReturn && !actualReturn) {
          statusCategory = 'overdue';
          const daysOverdue = Math.ceil((today - expectedReturn) / (1000 * 60 * 60 * 24));
          statusDescription = `Expected return ${daysOverdue} days ago`;
        } else if (today >= expectedReturn) {
          statusCategory = 'returning';
          statusDescription = 'Expected to return soon';
        }
        
        // Add payroll period context
        let currentPayrollPeriod = null;
        try {
          const staffType = c.employeePayType === 'Salary' ? 'Salaried' : 'Hourly';
          currentPayrollPeriod = getPayPeriodForDate(today, staffType);
        } catch (error) {
          console.log(`Could not determine payroll period for case ${c.caseId}: ${error.message}`);
        }
        
        return {
          ...summary,
          caseId: c.caseId,
          employeeName: c.employeeName,
          employeeLocation: c.employeeLocation,
          employeePayType: c.employeePayType,
          maternityStartDate: c.maternityStartDate,
          statusCategory: statusCategory,
          statusDescription: statusDescription,
          currentPayrollPeriod: currentPayrollPeriod ? currentPayrollPeriod.periodName : null
        };
      } catch (error) {
        console.log(`Error calculating display data for case ${c.caseId}: ${error.message}`);
        return {
          caseId: c.caseId,
          employeeName: c.employeeName,
          employeeLocation: c.employeeLocation,
          employeePayType: c.employeePayType,
          maternityStartDate: c.maternityStartDate,
          statusCategory: 'error',
          statusDescription: 'Data Error',
          error: error.message
        };
      }
    });
    
    // Enhanced statistics calculation
    const totalActiveCases = activeCases.length;
    const casesNeedingAttention = activeCases.filter(c => {
      return c.periods && c.periods.some(p => !p.dataComplete);
    }).length;
    
    const validSummaries = activeSummaries.filter(s => !s.error);
    const totalRemainingPay = validSummaries.reduce((sum, s) => {
      return sum + (s.smpRemaining || 0) + (s.cmpRemaining || 0);
    }, 0);
    
    // Payroll system health check
    const payrollHealth = checkMaternityPayrollHealth();
    
    return {
      success: true,
      allCases: allCases,
      activeCases: activeCases,
      activeSummaries: activeSummaries,
      statistics: {
        totalCases: allCases.length,
        activeCases: totalActiveCases,
        casesNeedingAttention: casesNeedingAttention,
        totalRemainingPay: Math.round(totalRemainingPay * 100) / 100,
        casesWithErrors: activeSummaries.filter(s => s.error).length
      },
      systemHealth: {
        payrollIntegration: payrollHealth.overall,
        lastChecked: new Date().toISOString(),
        issues: payrollHealth.recommendations
      }
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      message: error.message,
      error: error.toString()
    })
  }).result;
}

/**
 * ENHANCED: Server-side function to get case data for the period view modal.
 * NOW INCLUDES: Enhanced period organization and CMP tracking
 * @param {string} caseId - Maternity case ID
 * @returns {Object} Case data with enhanced period information
 */
function getCaseDataForPeriodView(caseId) {
  return safeExecute(() => {
    const maternityCase = getMaternityCase(caseId);
    if (!maternityCase) {
      throw new Error(`Case ${caseId} not found`);
    }
    
    // Calculate all display data with enhanced tracking
    const displayData = calculateEnhancedMaternityDisplayData(caseId);
    
    // Organize periods into month-based tabs with CMP information
    const periodTabs = organizePeriodsByMonth(maternityCase.periods);
    
    // Add CMP tracking information
    const cmpSummary = getCMPTrackingInfo(maternityCase);
    
    return {
      success: true,
      case: maternityCase,
      displayData: displayData,
      periodTabs: periodTabs,
      cmpSummary: cmpSummary,
      constants: {
        babyDueDate: maternityCase.babyDueDate,
        smpStartDate: maternityCase.smpStartDate,
        week39EndDate: displayData.week39EndDate,
        week52EndDate: displayData.week52EndDate,
        totalSMP: maternityCase.totalSMP,
        totalCMP: maternityCase.totalCMP
      }
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      message: error.message,
      error: error.toString()
    })
  }).result;
}

/**
 * ENHANCED: Calculates comprehensive display data with CMP tracking.
 * @param {string} caseId - Maternity case ID
 * @returns {Object} Enhanced display data object
 */

function calculateEnhancedMaternityDisplayData(caseId) {
  const baseData = calculateMaternityDisplayData(caseId);
  const maternityCase = getMaternityCase(caseId);
  
  if (!maternityCase) {
    return baseData;
  }
  
  // FIXED: Use consistent week-based tracking
  const cmpWeeksTotal = maternityCase.cmpWeeks || 8;
  const cmpDaysTotal = cmpWeeksTotal * 7;
  const smpStartDate = new Date(maternityCase.smpStartDate);
  
  let cmpWeeksUsed = 0;
  let cmpWeeksEligible = 0;
  
  // Sort periods to ensure correct processing order
  const sortedPeriods = [...maternityCase.periods].sort((a, b) => a.periodNumber - b.periodNumber);
  
  // FIXED: Use the same week calculation logic as main CMP calculation
  sortedPeriods.forEach(period => {
    const weeksInPeriod = calculateWeeksInMaternityPeriod(period, smpStartDate);
    
    if (weeksInPeriod > 0) {
      cmpWeeksEligible += weeksInPeriod;
      
      // Only count weeks used if there's CMP and we haven't exceeded the limit
      if (period.companyAmount > 0 && cmpWeeksUsed < cmpWeeksTotal) {
        const cmpWeeksAvailable = cmpWeeksTotal - cmpWeeksUsed;
        const eligibleWeeks = Math.min(weeksInPeriod, cmpWeeksAvailable);
        cmpWeeksUsed += eligibleWeeks;
      }
    }
  });
  
  // Convert to days for display consistency
  const cmpDaysUsed = Math.round(cmpWeeksUsed * 7);
  const cmpDaysEligible = Math.round(cmpWeeksEligible * 7);
  
  return {
    ...baseData,
    // FIXED: Enhanced CMP tracking using week-based calculations
    cmpDaysTotal: cmpDaysTotal,
    cmpDaysUsed: cmpDaysUsed,
    cmpDaysRemaining: Math.max(0, cmpDaysTotal - cmpDaysUsed),
    cmpDaysEligible: cmpDaysEligible,
    cmpCompletionPercentage: cmpDaysTotal > 0 ? Math.round((cmpDaysUsed / cmpDaysTotal) * 100) : 0,
    
    // Target weekly amount for reference
    targetWeeklyAmount: maternityCase.targetWeeklyAmount || 0,
    averageWeeklyEarnings: maternityCase.averageWeeklyEarnings || 0,
    contractedWeeklyEarnings: maternityCase.contractedWeeklyEarnings || 0,
    
    // Add debug info for troubleshooting
    debug: {
      cmpWeeksTotal: cmpWeeksTotal,
      cmpWeeksUsed: cmpWeeksUsed,
      cmpWeeksEligible: cmpWeeksEligible,
      calculationMethod: 'week-based-fixed'
    }
  };
}

/**
 * Gets CMP tracking information for a case.
 * @param {Object} maternityCase - Maternity case object
 * @returns {Object} CMP tracking summary
 */

function getCMPTrackingInfo(maternityCase) {
  const cmpWeeksTotal = maternityCase.cmpWeeks || 8;
  const cmpDaysTotal = cmpWeeksTotal * 7; // Convert weeks to days for display
  const smpStartDate = new Date(maternityCase.smpStartDate);
  
  let cmpWeeksUsed = 0; // Track in weeks (source of truth)
  let totalCMPPaid = 0;
  const cmpPeriods = [];
  
  // Sort periods by number to ensure correct order
  const sortedPeriods = [...maternityCase.periods].sort((a, b) => a.periodNumber - b.periodNumber);
  
  sortedPeriods.forEach(period => {
    if (period.companyAmount > 0) {
      // FIXED: Use the same week calculation as the main CMP logic
      const weeksInPeriod = calculateWeeksInMaternityPeriod(period, smpStartDate);
      
      // Calculate how many weeks this period used for CMP (respecting the limit)
      const cmpWeeksAvailable = cmpWeeksTotal - cmpWeeksUsed;
      const eligibleWeeks = Math.min(weeksInPeriod, cmpWeeksAvailable);
      
      // Only count weeks if there's CMP and we haven't exceeded the limit
      if (eligibleWeeks > 0 && cmpWeeksUsed < cmpWeeksTotal) {
        cmpWeeksUsed += eligibleWeeks;
        totalCMPPaid += period.companyAmount;
        
        cmpPeriods.push({
          periodNumber: period.periodNumber,
          periodName: period.periodName,
          cmpAmount: period.companyAmount,
          cmpWeeks: eligibleWeeks,
          cmpDays: Math.round(eligibleWeeks * 7), // Convert to days for display
          cmpNotes: period.companyNotes
        });
      }
    }
  });
  
  // Convert final weeks to days for display consistency
  const cmpDaysUsed = Math.round(cmpWeeksUsed * 7);
  
  return {
    cmpDaysTotal: cmpDaysTotal,
    cmpDaysUsed: cmpDaysUsed,
    cmpDaysRemaining: Math.max(0, cmpDaysTotal - cmpDaysUsed),
    totalCMPPaid: totalCMPPaid,
    cmpPeriods: cmpPeriods,
    isComplete: cmpWeeksUsed >= cmpWeeksTotal,
    targetWeeklyAmount: maternityCase.targetWeeklyAmount || 0,
    
    // Add debug info
    debug: {
      cmpWeeksTotal: cmpWeeksTotal,
      cmpWeeksUsed: cmpWeeksUsed,
      periodsWithCMP: cmpPeriods.length
    }
  };
}

/**
 * Gets employee list for form dropdowns with enhanced validation.
 * @returns {Array<Object>} Enhanced employee list
 */
function getEmployeesForMaternityForm() {
  return safeExecute(() => {
    const employees = getAllEmployees(); // From employeeDirectory.gs
    
    // Return enhanced employee data for form
    return employees
      .filter(emp => emp.IsPayrollEmployee?.toLowerCase() === "yes") // Only payroll employees
      .map(emp => ({
        id: emp.EmployeeNumber || emp.id,
        name: emp.FullName || `${emp.Firstnames} ${emp.Surname}`,
        location: emp.Location,
        payType: emp.PayType,
        salary: emp.Salary,
        hourlyRate: emp.HourlyShiftAmount,
        contractHours: emp.ContractHours,
        holidayEntitlement: emp.HolidayEntitlement || 25,
        // Add payroll context
        staffType: emp.PayType === 'Salary' ? 'Salaried' : 'Hourly'
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
    
  }, {
    fallback: (error) => {
      console.error('Error getting employees for form:', error);
      return [];
    }
  }).result;
}

// === ENHANCED VALIDATION FUNCTIONS ===

/**
 * ENHANCED: Validates SMP start date for form submission.
 * @param {string} employeeId - Employee ID
 * @param {string} smpStartDate - SMP start date
 * @returns {Object} Validation result
 */
function validateSMPStartDateForForm(employeeId, smpStartDate) {
  return safeExecute(() => {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      throw new Error('Employee not found');
    }
    
    const staffType = employee.PayType === 'Salary' ? 'Salaried' : 'Hourly';
    const validation = validateSMPStartDate(smpStartDate, staffType);
    
    return {
      success: true,
      validation: validation,
      employee: {
        name: employee.FullName || `${employee.Firstnames} ${employee.Surname}`,
        payType: employee.PayType,
        staffType: staffType
      }
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      error: error.message
    })
  }).result;
}

// === TAB ORGANIZATION FUNCTIONS ===

/**
 * Organizes periods into month-based tabs for the modal view.
 * @param {Array<Object>} periods - Array of period objects
 * @returns {Array<Object>} Array of tab objects with periods grouped by month
 */
function organizePeriodsByMonth(periods) {
  const tabs = [];
  const monthGroups = {};
  
  // Group periods by month and year
  periods.forEach(period => {
    const periodStart = new Date(period.periodStart);
    const monthKey = `${periodStart.getFullYear()}-${String(periodStart.getMonth() + 1).padStart(2, '0')}`;
    const monthName = periodStart.toLocaleDateString('en-GB', { year: 'numeric', month: 'long' });
    
    if (!monthGroups[monthKey]) {
      monthGroups[monthKey] = {
        monthKey: monthKey,
        monthName: monthName,
        periods: []
      };
    }
    
    monthGroups[monthKey].periods.push(period);
  });
  
  // Convert to sorted array
  Object.keys(monthGroups)
    .sort()
    .forEach((monthKey, index) => {
      tabs.push({
        tabId: `tab-${index + 1}`,
        tabNumber: index + 1,
        monthKey: monthKey,
        monthName: monthGroups[monthKey].monthName,
        periods: monthGroups[monthKey].periods.sort((a, b) => 
          new Date(a.periodStart) - new Date(b.periodStart)
        ),
        periodCount: monthGroups[monthKey].periods.length
      });
    });
  
  return tabs;
}

// === ENHANCED MODAL DISPLAY FUNCTIONS ===

/**
 * Shows the enhanced maternity management dashboard.
 */
function showMaternityDashboard() {
  const dashboardData = getMaternityDashboardData();
  
  if (!dashboardData.success) {
    SpreadsheetApp.getUi().alert('Dashboard Error', dashboardData.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Show system health warnings if needed
  if (dashboardData.systemHealth.payrollIntegration !== 'healthy') {
    const healthMessage = `Payroll system health: ${dashboardData.systemHealth.payrollIntegration}. ${dashboardData.systemHealth.issues.join(' ')}`;
    console.log('⚠️ ' + healthMessage);
  }
  
  const html = HtmlService.createTemplateFromFile('06_maternityDashboard');
  html.dashboardData = dashboardData;
  
  const output = html.evaluate()
    .setWidth(1000)
    .setHeight(700)
    .setTitle('🤱 Enhanced Maternity Management Dashboard');
  
  SpreadsheetApp.getUi().showModalDialog(output, 'Maternity Management');
}

/**
 * Shows enhanced detailed period view for a specific maternity case.
 * @param {string} caseId - Case ID to show
 */
function showMaternityPeriodView(caseId) {
  const periodData = getCaseDataForPeriodView(caseId);
  
  if (!periodData.success) {
    SpreadsheetApp.getUi().alert('Period View Error', periodData.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const html = HtmlService.createTemplateFromFile('06_maternityPeriodView');
  html.periodData = periodData;
  
  const output = html.evaluate()
    .setWidth(1100)
    .setHeight(800)
    .setTitle(`📅 ${periodData.case.employeeName} - Enhanced Period Details`);
  
  SpreadsheetApp.getUi().showModalDialog(output, 'Enhanced Maternity Period View');
}

/**
 * Server-side function to set actual return date.
 * @param {string} caseId - Case ID to update
 * @param {string} returnDate - Actual return date (YYYY-MM-DD format)
 * @returns {Object} Result object
 */
function submitReturnDate(caseId, returnDate) {
  return safeExecute(() => {
    if (!returnDate) {
      throw new Error('Return date is required');
    }
    
    // Validate return date is reasonable
    const returnDateObj = new Date(returnDate);
    const today = new Date();
    
    if (returnDateObj > today) {
      // Future return date - validate it's not too far in the future
      const daysDifference = (returnDateObj - today) / (1000 * 60 * 60 * 24);
      if (daysDifference > 365) {
        throw new Error('Return date cannot be more than 1 year in the future');
      }
    }
    
    const updates = {
      actualReturnDate: returnDateObj
    };
    
    const updatedCase = updateMaternityCase(caseId, updates);
    
    return {
      success: true,
      message: `Return date set for ${updatedCase.employeeName}. Periods and CMP have been recalculated.`,
      case: updatedCase
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      message: error.message,
      error: error.toString()
    })
  }).result;
}

/**
 * Server-side function to archive a case.
 * @param {string} caseId - Case ID to archive
 * @param {string} reason - Reason for archiving
 * @returns {Object} Result object
 */
function submitArchiveCase(caseId, reason) {
  return safeExecute(() => {
    if (!reason || reason.trim() === '') {
      throw new Error('Archive reason is required');
    }
    
    const archivedCase = archiveMaternityCase(caseId, reason.trim());
    
    return {
      success: true,
      message: `Case for ${archivedCase.employeeName} has been archived`,
      case: archivedCase
    };
    
  }, {
    fallback: (error) => ({
      success: false,
      message: error.message,
      error: error.toString()
    })
  }).result;
}

/**
 * Calculates display data for a maternity case 
 * @param {string} caseId - Maternity case ID
 * @returns {Object} Display data object
 */

/**
 * FIXED: Calculates display data for a maternity case 
 * Now uses consistent week-based logic instead of recalculating days
 * @param {string} caseId - Maternity case ID
 * @returns {Object} Display data object
 */
function calculateMaternityDisplayData(caseId) {
  try {
    const maternityCase = getMaternityCase(caseId);
    if (!maternityCase) {
      throw new Error(`Case ${caseId} not found`);
    }
    
    // Calculate basic totals from stored values (these are correct)
    let totalSMPPaid = 0;
    let totalCMPPaid = 0;
    let periodsWithData = 0;
    
    maternityCase.periods.forEach(period => {
      const smpAmount = parseFloat(period.smpAmount) || 0;
      const cmpAmount = parseFloat(period.companyAmount) || 0;
      
      totalSMPPaid += smpAmount;
      totalCMPPaid += cmpAmount;
      
      if (smpAmount > 0 || cmpAmount > 0 || period.dataComplete) {
        periodsWithData++;
      }
    });
    
    // Calculate remaining amounts
    const totalSMP = parseFloat(maternityCase.totalSMP) || 0;
    const totalCMP = parseFloat(maternityCase.totalCMP) || 0;
    const smpRemaining = Math.max(0, totalSMP - totalSMPPaid);
    const cmpRemaining = Math.max(0, totalCMP - totalCMPPaid);
    
    // Calculate important dates
    const smpStartDate = new Date(maternityCase.smpStartDate);
    const week39EndDate = new Date(smpStartDate);
    week39EndDate.setDate(week39EndDate.getDate() + (39 * 7) - 1); // 39 weeks from SMP start
    
    const week52EndDate = new Date(smpStartDate);
    week52EndDate.setDate(week52EndDate.getDate() + (52 * 7) - 1); // 52 weeks from SMP start
    
    // FIXED: Use consistent week-based CMP tracking
    const cmpWeeksTotal = maternityCase.cmpWeeks || 8;
    const cmpDaysTotal = cmpWeeksTotal * 7;
    let cmpWeeksUsed = 0;
    
    // Sort periods to ensure correct processing order
    const sortedPeriods = [...maternityCase.periods].sort((a, b) => a.periodNumber - b.periodNumber);
    
    // FIXED: Use the same week calculation logic as main CMP calculation
    sortedPeriods.forEach(period => {
      if (period.companyAmount > 0 && cmpWeeksUsed < cmpWeeksTotal) {
        const weeksInPeriod = calculateWeeksInMaternityPeriod(period, smpStartDate);
        const cmpWeeksAvailable = cmpWeeksTotal - cmpWeeksUsed;
        const eligibleWeeks = Math.min(weeksInPeriod, cmpWeeksAvailable);
        cmpWeeksUsed += eligibleWeeks;
      }
    });
    
    // Convert to days for display consistency
    const cmpDaysUsed = Math.round(cmpWeeksUsed * 7);
    
    return {
      totalSMP: totalSMP,
      totalSMPPaid: totalSMPPaid,
      smpRemaining: smpRemaining,
      totalCMP: totalCMP,
      totalCMPPaid: totalCMPPaid,
      cmpRemaining: cmpRemaining,
      totalPeriods: maternityCase.periods.length,
      periodsWithData: periodsWithData,
      week39EndDate: week39EndDate,
      week52EndDate: week52EndDate,
      // FIXED: Consistent CMP days tracking
      cmpDaysTotal: cmpDaysTotal,
      cmpDaysUsed: cmpDaysUsed,
      cmpDaysRemaining: Math.max(0, cmpDaysTotal - cmpDaysUsed),
      
      // Add debug info
      debug: {
        cmpWeeksTotal: cmpWeeksTotal,
        cmpWeeksUsed: cmpWeeksUsed,
        calculationMethod: 'week-based-fixed'
      }
    };
    
  } catch (error) {
    console.error(`Error calculating display data for case ${caseId}: ${error.message}`);
    
    // Return safe defaults
    return {
      totalSMP: 0,
      totalSMPPaid: 0,
      smpRemaining: 0,
      totalCMP: 0,
      totalCMPPaid: 0,
      cmpRemaining: 0,
      totalPeriods: 0,
      periodsWithData: 0,
      week39EndDate: new Date(),
      week52EndDate: new Date(),
      cmpDaysTotal: 56,
      cmpDaysUsed: 0,
      cmpDaysRemaining: 56,
      error: error.message
    };
  }
}

// === ENHANCED DEBUGGING AND TESTING ===

/**
 * ENHANCED: Tests the complete maternity system end-to-end.
 * @returns {Object} Comprehensive test results
 */
function testMaternitySystemEndToEnd() {
  console.log('🧪 Running comprehensive maternity system test...');
  
  const testResults = {
    overall: 'success',
    tests: {},
    summary: ''
  };
  
  try {
    // Test 1: Payroll calendar health
    testResults.tests.payrollHealth = checkMaternityPayrollHealth();
    
    // Test 2: Employee directory integration
    try {
      const employees = getEmployeesForMaternityForm();
      testResults.tests.employeeDirectory = {
        status: employees.length > 0 ? 'success' : 'warning',
        message: `${employees.length} payroll employees found`,
        data: { count: employees.length }
      };
    } catch (error) {
      testResults.tests.employeeDirectory = {
        status: 'error',
        message: `Employee directory test failed: ${error.message}`
      };
    }
    
    // Test 3: CMP calculation logic
    try {
      const cmpTest = testEnhancedCMPCalculation();
      testResults.tests.cmpCalculation = {
        status: cmpTest.success ? 'success' : 'error',
        message: cmpTest.success ? 'CMP calculation working' : cmpTest.error,
        data: cmpTest
      };
    } catch (error) {
      testResults.tests.cmpCalculation = {
        status: 'error',
        message: `CMP calculation test failed: ${error.message}`
      };
    }
    
    // Test 4: Dashboard data preparation
    try {
      const dashboardData = getMaternityDashboardData();
      testResults.tests.dashboardData = {
        status: dashboardData.success ? 'success' : 'error',
        message: dashboardData.success ? 'Dashboard data generation working' : dashboardData.message,
        data: { 
          activeCases: dashboardData.activeCases?.length || 0,
          totalCases: dashboardData.allCases?.length || 0
        }
      };
    } catch (error) {
      testResults.tests.dashboardData = {
        status: 'error',
        message: `Dashboard data test failed: ${error.message}`
      };
    }
    
    // Determine overall status
    const testStatuses = Object.values(testResults.tests).map(test => test.status);
    if (testStatuses.includes('error')) {
      testResults.overall = 'error';
    } else if (testStatuses.includes('warning')) {
      testResults.overall = 'warning';
    }
    
    // Generate summary
    const successCount = testStatuses.filter(s => s === 'success').length;
    const totalTests = testStatuses.length;
    testResults.summary = `${successCount}/${totalTests} tests passed. Overall: ${testResults.overall}`;
    
  } catch (error) {
    testResults.overall = 'error';
    testResults.summary = `System test failed: ${error.message}`;
  }
  
  console.log(`✅ Maternity system test complete: ${testResults.summary}`);
  return testResults;
}
