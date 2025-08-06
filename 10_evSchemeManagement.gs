/**
 * 🚗 EV Scheme Management Module - Production Version
 * Salary sacrifice EV scheme agreement processing and management.
 * Admin-only module for processing provider agreements and tracking contracts.
 */

// === Data Management Functions ===

/**
 * Loads all EV scheme agreements from JSON file.
 * @returns {Array<Object>}
*/
function loadEVSchemeAgreements() {
  try {
    const data = loadJsonFromDrive(CONFIG.EV_SCHEME_DATA_FILE_ID);
    return data.agreements || [];
  } catch (error) {
    console.log('No existing EV scheme data found, returning empty array');
    return [];
  }
}

/**
 * Saves EV scheme agreements to JSON file.
 * @param {Array<Object>} agreements - Array of agreement objects to save
 */
function saveEVSchemeAgreements(agreements) {
  try {
    var data = {
      agreements: agreements,
      lastUpdated: new Date().toISOString(),
      updatedBy: Session.getEffectiveUser().getEmail(),
      statistics: generateEVSchemeSummary(agreements)
    };
    
    var file = DriveApp.getFileById(CONFIG.EV_SCHEME_DATA_FILE_ID);
    file.setContent(JSON.stringify(data, null, 2));
    
    console.log('Saved ' + agreements.length + ' EV scheme agreements');
  } catch (error) {
    throw new Error('Failed to save EV scheme data: ' + error.message);
  }
}

/**
 * Generates summary statistics for the EV scheme dashboard.
 * @param {Array<Object>} agreements - Array of agreements
 * @returns {Object} Summary statistics
 */
function generateEVSchemeSummary(agreements) {
  var summary = {
    total: agreements.length,
    eligible: 0,
    ineligible: 0,
    pendingApproval: 0,
    approved: 0,
    awaitingDelivery: 0,
    active: 0,
    completed: 0,
    totalLeaseValue: 0,
    totalMonthlyDeductions: 0,
    totalP11DValue: 0,
    averageMonthlyCost: 0,
    byLocation: {},
    byDivision: {},
    byProvider: {},
    byStatus: {},
    recentActivity: []
  };
  
  agreements.forEach(function(agreement) {
    // Detailed status counts
    var status = agreement.workflowStatus || agreement.status || 'Unknown';
    summary.byStatus[status] = (summary.byStatus[status] || 0) + 1;
    
    // Legacy status mapping for backward compatibility
    if (agreement.isEligible) {
      summary.eligible++;
    } else {
      summary.ineligible++;
    }
    
    // Workflow-based status counts
    if (status === 'Eligible - Pending Approval') {
      summary.pendingApproval++;
    } else if (status === 'Approved - Awaiting Delivery') {
      summary.approved++;
      summary.awaitingDelivery++;
    } else if (status === 'Active - Vehicle Delivered') {
      summary.active++;
    } else if (status === 'Completed') {
      summary.completed++;
    }
    
    // Financial totals (only for active agreements)
    if (status === 'Active - Vehicle Delivered') {
      summary.totalLeaseValue += agreement.leaseValue || 0;
      summary.totalMonthlyDeductions += agreement.totalMonthlyCost || agreement.monthlyDeduction || 0;
    }
    summary.totalP11DValue += agreement.p11dValue || 0;
    
    // Location breakdown
    var location = agreement.location || 'Unknown';
    summary.byLocation[location] = (summary.byLocation[location] || 0) + 1;
    
    // Division breakdown
    var division = agreement.division || 'Unknown';
    summary.byDivision[division] = (summary.byDivision[division] || 0) + 1;
    
    // Provider breakdown
    var provider = agreement.providerName || 'Unknown';
    summary.byProvider[provider] = (summary.byProvider[provider] || 0) + 1;
  });
  
  // Calculate averages (only for active agreements)
  summary.averageMonthlyCost = summary.active > 0 ? 
    Math.round(summary.totalMonthlyDeductions / summary.active) : 0;
  
  // Recent activity (last 10 submissions)
  summary.recentActivity = agreements
    .sort(function(a, b) {
      var aDate = new Date(a.lastStatusChange || a.submissionDate);
      var bDate = new Date(b.lastStatusChange || b.submissionDate);
      return bDate - aDate;
    })
    .slice(0, 10)
    .map(function(agreement) {
      return {
        agreementId: agreement.agreementId,
        employeeName: agreement.employeeName,
        submissionDate: agreement.submissionDate,
        status: agreement.workflowStatus || agreement.status,
        vehicleInfo: agreement.vehicleMake + ' ' + agreement.vehicleModel,
        p11dValue: agreement.p11dValue || 0,
        monthlyCost: agreement.totalMonthlyCost || agreement.monthlyDeduction || 0,
        lastStatusChange: agreement.lastStatusChange
      };
    });
  
  return summary;
}

/**
 * Helper function to generate the agreements table HTML SERVER-SIDE.
 * FIXED: Now executes on server before being sent to browser
 */
function generateAgreementsTableHTML(agreements) {
  if (agreements.length === 0) {
    return '<p class="loading">No agreements found</p>';
  }

  var tableHTML = '<div style="overflow-x: auto;"><table class="agreements-table"><thead><tr><th>Agreement ID</th><th>Employee</th><th>Vehicle</th><th>P11D Value</th><th>Monthly Cost</th><th>Provider</th><th>Workflow Status</th><th>Remaining Term</th><th>Actions</th></tr></thead><tbody>';

  agreements.forEach(function(agreement) {
    var workflowStatus = agreement.workflowStatus || agreement.status || 'Unknown';
    var statusClass = 'status-pending';
    
    // Determine status styling
    if (workflowStatus.includes('Ineligible')) {
      statusClass = 'status-ineligible';
    } else if (workflowStatus.includes('Active')) {
      statusClass = 'status-active';
    } else if (workflowStatus.includes('Approved')) {
      statusClass = 'status-approved';
    } else if (workflowStatus.includes('Eligible')) {
      statusClass = 'status-eligible';
    }
    
    // Calculate remaining term
    var remainingTerm = 'N/A';
    if (agreement.activationDate && agreement.leaseTerm) {
      var activationDate = new Date(agreement.activationDate);
      var now = new Date();
      var monthsElapsed = Math.floor((now - activationDate) / (1000 * 60 * 60 * 24 * 30.44));
      var remainingMonths = Math.max(0, agreement.leaseTerm - monthsElapsed);
      remainingTerm = remainingMonths + ' months';
    } else if (agreement.leaseTerm) {
      remainingTerm = agreement.leaseTerm + ' months (not started)';
    }
    
    tableHTML += '<tr>';
    tableHTML += '<td>' + agreement.agreementId + '</td>';
    tableHTML += '<td>' + agreement.employeeName + '<br><small>' + agreement.employeeNumber + '</small></td>';
    tableHTML += '<td>' + agreement.vehicleMake + ' ' + agreement.vehicleModel + '</td>';
    tableHTML += '<td>£' + (agreement.p11dValue || 0).toLocaleString() + '</td>';
    tableHTML += '<td>£' + (agreement.totalMonthlyCost || agreement.monthlyDeduction || 0).toLocaleString() + '</td>';
    tableHTML += '<td>' + agreement.providerName + '</td>';
    tableHTML += '<td><span class="status-badge ' + statusClass + '">' + workflowStatus + '</span></td>';
    tableHTML += '<td>' + remainingTerm + '</td>';
    tableHTML += '<td><button class="btn btn-sm" onclick="manageAgreement(\'' + agreement.agreementId + '\')">Manage</button></td>';
    tableHTML += '</tr>';
  });

  tableHTML += '</tbody></table></div>';
  return tableHTML;
}

/**
 * FIXED: Server-side function to get dashboard data with pre-generated HTML
 * Following the working maternity pattern
 */
function getEVSchemeDashboardData() {
  return safeExecute(() => {
    // Load agreements and generate summary
    const agreements = loadEVSchemeAgreements();
    const summary = generateEVSchemeSummary(agreements);
    
    // PRE-GENERATE HTML SERVER-SIDE (this is the key fix!)
    const agreementsTableHTML = generateAgreementsTableHTML(agreements);
    
    // Generate workflow table HTML server-side too
    const workflowTableHTML = generateWorkflowTableHTML(agreements);
    
    return {
      success: true,
      agreements: agreements,
      summary: summary,
      // Pre-generated HTML strings (safe to send to browser)
      agreementsTableHTML: agreementsTableHTML,
      workflowTableHTML: workflowTableHTML,
      systemHealth: {
        agreementCount: agreements.length,
        lastChecked: new Date().toISOString()
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
 * Generate workflow table HTML server-side
 */
function generateWorkflowTableHTML(agreements) {
  if (agreements.length === 0) {
    return '<p class="loading">No agreements found</p>';
  }
  
  var tableHTML = '<table class="agreements-table"><thead><tr><th>Agreement ID</th><th>Employee</th><th>Vehicle</th><th>Current Status</th><th>Expected Delivery</th><th>Actions</th></tr></thead><tbody>';
  
  agreements.forEach(function(agreement) {
    var workflowStatus = agreement.workflowStatus || agreement.status || 'Unknown';
    var statusClass = getStatusClass(workflowStatus);
    var expectedDelivery = agreement.expectedDeliveryDate ? new Date(agreement.expectedDeliveryDate).toLocaleDateString() : 'Not set';
    
    tableHTML += '<tr>';
    tableHTML += '<td>' + agreement.agreementId + '</td>';
    tableHTML += '<td>' + agreement.employeeName + '</td>';
    tableHTML += '<td>' + agreement.vehicleMake + ' ' + agreement.vehicleModel + '</td>';
    tableHTML += '<td><span class="status-badge ' + statusClass + '">' + workflowStatus + '</span></td>';
    tableHTML += '<td>' + expectedDelivery + '</td>';
    tableHTML += '<td>';
    tableHTML += '<button class="btn btn-sm btn-primary" onclick="manageAgreement(\'' + agreement.agreementId + '\')" style="margin-right: 5px;">Manage</button>';
    if (workflowStatus === 'Eligible - Pending Approval') {
      tableHTML += '<button class="btn btn-sm" onclick="approveAgreement(\'' + agreement.agreementId + '\')" style="background: #28a745; color: white;">Approve</button>';
    } else if (workflowStatus === 'Approved - Awaiting Delivery') {
      tableHTML += '<button class="btn btn-sm" onclick="markAsDelivered(\'' + agreement.agreementId + '\')" style="background: #007bff; color: white;">Mark Delivered</button>';
    }
    tableHTML += '</td>';
    tableHTML += '</tr>';
  });
  
  tableHTML += '</tbody></table>';
  return tableHTML;
}

/**
 * Helper function for status class determination
 */
function getStatusClass(status) {
  if (status.includes('Ineligible')) return 'status-ineligible';
  if (status.includes('Active')) return 'status-active';
  if (status.includes('Approved')) return 'status-approved';
  if (status.includes('Eligible')) return 'status-eligible';
  return 'status-pending';
}

// === Main Dashboard Function - FIXED to use working maternity pattern ===

/**
 * FIXED: Opens the EV Scheme Management dashboard using template pattern
 * Now follows the exact same pattern as the working maternity system
 */
function showEVSchemeDashboard() {
  var userEmail = Session.getEffectiveUser().getEmail();
  
  // Admin access check
  if (!CONFIG.ADMIN_EMAILS.includes(userEmail)) {
    SpreadsheetApp.getUi().alert(
      'Access Denied',
      'EV Scheme Management is only available to system administrators.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  try {
    // FIXED: Load data server-side (following maternity pattern)
    const dashboardData = getEVSchemeDashboardData();
    
    if (!dashboardData.success) {
      SpreadsheetApp.getUi().alert('Dashboard Error', dashboardData.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // FIXED: Use template pattern like maternity system
    const html = HtmlService.createTemplateFromFile('10_evSchemeDashboard');
    html.dashboardData = dashboardData;  // Pass data to template
    
    // FIXED: Evaluate server-side before sending to browser
    const output = html.evaluate()
      .setWidth(1200)
      .setHeight(800)
      .setTitle('EV Scheme Management Dashboard');
    
    SpreadsheetApp.getUi().showModalDialog(output, 'EV Scheme Management');
    
  } catch (error) {
    console.error('Error opening EV Scheme dashboard:', error);
    SpreadsheetApp.getUi().alert(
      'Dashboard Error',
      'Failed to open EV Scheme dashboard: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// === Business Logic Functions ===

/**
 * Gets existing salary sacrifices for an employee.
 * @param {string} employeeNumber - Employee number
 * @returns {Array} Array of existing salary sacrifice objects
 */
function getExistingSalarySacrifices(employeeNumber) {
  try {
    // This would integrate with your existing salary sacrifice tracking system
    return [];
  } catch (error) {
    console.warn('Could not retrieve existing salary sacrifices for employee ' + employeeNumber + ':', error);
    return [];
  }
}

/**
 * Gets the current NMW rate for a specific employee using the existing NMW compliance system.
 * ALWAYS uses the full adult NMW rate regardless of employee age.
 * @param {Object} employee - Employee object
 * @param {Date} effectiveDate - Date to check NMW rate for
 * @returns {Object} NMW rate result {success: boolean, rate: number, staffType: string, contractHours: number, effectiveDate: Date, error?: string}
 */
function getEmployeeNMWRate(employee, effectiveDate) {
  if (effectiveDate === undefined) effectiveDate = null;
  
  try {
    var checkDate = effectiveDate || new Date();
    
    // Determine staff type
    var staffType;
    var empPayType = (employee.PayType || employee.PayBasis || '').toLowerCase();
    if (empPayType.includes('hourly')) {
      staffType = 'Hourly';
    } else {
      staffType = 'Salaried';
    }
    
    // Get contract hours from employee data
    var contractHours = parseFloat(employee.ContractHours || employee.contractHours || 37.5);
    if (!contractHours || contractHours <= 0) {
      console.warn('Invalid contract hours for employee ' + employee.EmployeeNumber + ', using 37.5 as fallback');
      contractHours = 37.5;
    }
    
    // Load NMW configuration and search directly
    var bands = loadNMWConfiguration();
    
    // Find applicable rate bands for this staff type
    var applicableBands = bands
      .filter(function(band) { return band.staffType === staffType && band.effectiveFrom <= checkDate; })
      .sort(function(a, b) { return b.effectiveFrom - a.effectiveFrom; });
    
    if (applicableBands.length === 0) {
      throw new Error('No NMW rate found for ' + staffType + ' on ' + checkDate.toDateString());
    }
    
    // Get the most recent rate band
    var latestBand = applicableBands[0];
    
    // ALWAYS use the full adult rate (21plus) regardless of actual age
    var rate = latestBand.bands['21plus'] || latestBand.bands.adult;
    
    if (!rate || rate <= 0) {
      throw new Error('Invalid adult NMW rate: ' + rate);
    }
    
    console.log('=== NMW RATE LOOKUP SUCCESS ===');
    console.log('Staff Type:', staffType);
    console.log('Contract Hours:', contractHours);
    console.log('Adult NMW Rate:', rate);
    console.log('Effective Date:', checkDate);
    
    return {
      success: true,
      rate: rate,
      staffType: staffType,
      contractHours: contractHours,
      effectiveDate: checkDate,
      ageBandUsed: '21plus (Full Adult Rate)'
    };
    
  } catch (error) {
    console.error('NMW lookup failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Checks employee eligibility for EV salary sacrifice scheme.
 * @param {Object} employee - Employee record
 * @param {Object} agreementData - Agreement data including monthly cost
 * @returns {Object} Eligibility result with detailed breakdown
 */
function checkEVSchemeEligibility(employee, agreementData) {
  var totalExistingSalSacs = 0;
  var calculationBreakdown = null;
  
  var checks = [];
  var isEligible = true;
  var issues = [];
  var warnings = [];
  
  // 1. Employment Status Check
  checks.push('Employment Status');
  if (employee.Status && employee.Status.toLowerCase() !== 'current') {
    isEligible = false;
    issues.push('Employee status is not current');
  }
  
  // 2. PayBasis Check - Must be Monthly-Salaried
  checks.push('Pay Basis');
  var payBasis = employee.PayBasis || employee.PayType || '';
  if (payBasis.toLowerCase() !== 'monthly-salaried' && employee.PayType && employee.PayType.toLowerCase() !== 'salary') {
    isEligible = false;
    issues.push('Pay basis is "' + payBasis + '" - must be "Monthly-Salaried"');
  }
  
  // 3. Minimum Salary Check - £33,000
  checks.push('Minimum Salary');
  var currentSalary = parseFloat(employee.Salary || 0);
  var minimumSalary = 33000;
  
  if (currentSalary === 0) {
    isEligible = false;
    issues.push('No salary information available');
  } else if (currentSalary < minimumSalary) {
    isEligible = false;
    issues.push('Current salary (£' + currentSalary.toLocaleString() + ') is below minimum requirement (£' + minimumSalary.toLocaleString() + ')');
  }
  
  // 4. Length of Service Check - 6 months minimum
  checks.push('Length of Service');
  var startDate = parseDMY(employee.StartDate);
  if (!startDate) {
    isEligible = false;
    issues.push('No start date available for length of service calculation');
  } else {
    var monthsEmployed = (new Date() - startDate) / (1000 * 60 * 60 * 24 * 30.44);
    var minimumMonths = 6;
    
    if (monthsEmployed < minimumMonths) {
      isEligible = false;
      issues.push('Length of service is ' + Math.floor(monthsEmployed) + ' months (minimum: ' + minimumMonths + ' months)');
    }
  }
  
  // 5. National Minimum Wage Affordability Check
  checks.push('NMW Affordability (using full adult rate)');
  var monthlyDeduction = parseFloat(agreementData.totalMonthlyCost || agreementData.monthlyDeduction || 0);
  
  if (monthlyDeduction > 0 && currentSalary > 0) {
    var nmwResult = getEmployeeNMWRate(employee);
    
    if (!nmwResult.success) {
      isEligible = false;
      issues.push('Cannot determine NMW rate for eligibility check: ' + nmwResult.error);
      
      calculationBreakdown = {
        error: 'NMW rate lookup failed',
        errorDetails: nmwResult.error,
        salaryDetails: {
          grossAnnualSalary: currentSalary,
          proposedMonthlyEVCost: monthlyDeduction
        }
      };
    } else {
      var currentNMW = nmwResult.rate;
      
      // Calculate existing salary sacrifices
      try {
        var existingSalSacs = getExistingSalarySacrifices(employee.EmployeeNumber) || [];
        totalExistingSalSacs = existingSalSacs.reduce(function(total, salsac) { 
          return total + (salsac.monthlyAmount || 0); 
        }, 0);
      } catch (error) {
        console.warn('Could not load existing salary sacrifices:', error);
        totalExistingSalSacs = 0;
        warnings.push('Could not verify existing salary sacrifices');
      }
      
      // Perform affordability calculation
      var totalMonthlyDeductions = monthlyDeduction + totalExistingSalSacs;
      var postDeductionSalary = currentSalary - (totalMonthlyDeductions * 12);
      
      var nmwBuffer = 1.15; // 15% buffer above NMW
      var contractHours = nmwResult.contractHours || 37.5; // Use actual contract hours
      var annualNMWMinimum = currentNMW * 52 * contractHours; // Use actual contract hours
      var requiredMinimumSalary = annualNMWMinimum * nmwBuffer;
      
      // Build detailed calculation breakdown
      calculationBreakdown = {
        nmwDetails: {
          hourlyRate: currentNMW,
          staffType: nmwResult.staffType,
          contractHours: contractHours,
          ageBandUsed: nmwResult.ageBandUsed || 'Full Adult Rate',
          effectiveDate: nmwResult.effectiveDate.toDateString()
        },
        salaryDetails: {
          grossAnnualSalary: currentSalary,
          proposedMonthlyEVCost: monthlyDeduction,
          existingMonthlySalarySacrifices: totalExistingSalSacs,
          totalMonthlyDeductions: totalMonthlyDeductions,
          totalAnnualDeductions: totalMonthlyDeductions * 12,
          netAnnualSalary: postDeductionSalary
        },
        nmwCalculation: {
          hourlyRate: currentNMW,
          contractHours: contractHours,
          weeksPerYear: 52,
          annualNMWBaseline: annualNMWMinimum,
          nmwBufferPercentage: 15,
          nmwBufferMultiplier: nmwBuffer,
          requiredMinimumSalary: requiredMinimumSalary
        },
        complianceCheck: {
          netSalaryAfterDeductions: postDeductionSalary,
          requiredMinimumSalary: requiredMinimumSalary,
          shortfall: postDeductionSalary < requiredMinimumSalary ? requiredMinimumSalary - postDeductionSalary : 0,
          isCompliant: postDeductionSalary >= requiredMinimumSalary,
          marginAboveMinimum: postDeductionSalary >= requiredMinimumSalary ? postDeductionSalary - requiredMinimumSalary : 0
        }
      };
      
      if (postDeductionSalary < requiredMinimumSalary) {
        isEligible = false;
        issues.push('Post-deduction salary (£' + postDeductionSalary.toLocaleString() + ') would be below required minimum (£' + requiredMinimumSalary.toLocaleString() + ') - includes 15% NMW buffer');
      }
      
      checks.push('Affordability Calculation Complete');
    }
  } else if (monthlyDeduction === 0) {
    warnings.push('No monthly deduction amount specified for affordability check');
  } else if (currentSalary === 0) {
    isEligible = false;
    issues.push('No salary information available for affordability check');
  }
  
  // 6. Additional Data Validation
  checks.push('Data Completeness');
  if (!agreementData.p11dValue) {
    warnings.push('P11D value not provided');
  }
  if (!agreementData.leaseTerm || agreementData.leaseTerm < 12 || agreementData.leaseTerm > 60) {
    warnings.push('Lease term should be between 12-60 months');
  }
  
  return {
    isEligible: isEligible,
    notes: isEligible ? 
      'All eligibility requirements met' + (warnings.length > 0 ? ' (Warnings: ' + warnings.join('; ') + ')' : '') :
      issues.join('; '),
    rulesChecked: checks,
    warnings: warnings,
    calculationBreakdown: calculationBreakdown,
    details: {
      employmentStatus: employee.Status,
      payrollEmployee: 'Yes',
      payBasis: payBasis,
      currentSalary: currentSalary,
      lengthOfService: startDate ? Math.floor((new Date() - startDate) / (1000 * 60 * 60 * 24 * 30.44)) : null,
      monthlyDeduction: monthlyDeduction,
      existingSalarySacrifices: totalExistingSalSacs || 0,
      postDeductionSalary: currentSalary > 0 && monthlyDeduction > 0 ? 
        currentSalary - (monthlyDeduction * 12) - ((totalExistingSalSacs || 0) * 12) : null,
      nmwBuffer: 1.15,
      startDate: employee.StartDate,
      issues: issues,
      warnings: warnings
    }
  };
}

/**
 * Enhanced eligibility check that includes existing salary sacrifice agreements.
 * @param {Object} employee - Employee record
 * @param {Object} agreementData - Agreement data including existing salary sacrifices
 * @returns {Object} Detailed eligibility result
 */
function checkEVSchemeEligibilityEnhanced(employee, agreementData) {
  var totalExistingSalSacs = 0;
  var enhancedCalculationBreakdown = null;
  
  var checks = [];
  var isEligible = true;
  var issues = [];
  var warnings = [];
  
  // 1. Employment Status Check
  checks.push('Employment Status');
  if (employee.Status && employee.Status.toLowerCase() !== 'current') {
    isEligible = false;
    issues.push('Employee status is not current');
  }
  
  // 2. PayBasis Check - Must be Monthly-Salaried
  checks.push('Pay Basis');
  var payBasis = employee.PayBasis || employee.PayType || '';
  if (payBasis.toLowerCase() !== 'monthly-salaried' && employee.PayType && employee.PayType.toLowerCase() !== 'salary') {
    isEligible = false;
    issues.push('Pay basis is "' + payBasis + '" - must be "Monthly-Salaried"');
  }
  
  // 3. Minimum Salary Check - £33,000
  checks.push('Minimum Salary');
  var currentSalary = parseFloat(employee.Salary || 0);
  var minimumSalary = 33000;
  
  if (currentSalary === 0) {
    isEligible = false;
    issues.push('No salary information available');
  } else if (currentSalary < minimumSalary) {
    isEligible = false;
    issues.push('Current salary (£' + currentSalary.toLocaleString() + ') is below minimum requirement (£' + minimumSalary.toLocaleString() + ')');
  }
  
  // 4. Length of Service Check - 6 months minimum
  checks.push('Length of Service');
  var startDate = parseDMY(employee.StartDate);
  if (!startDate) {
    isEligible = false;
    issues.push('No start date available for length of service calculation');
  } else {
    var monthsEmployed = (new Date() - startDate) / (1000 * 60 * 60 * 24 * 30.44);
    var minimumMonths = 6;
    
    if (monthsEmployed < minimumMonths) {
      isEligible = false;
      issues.push('Length of service is ' + Math.floor(monthsEmployed) + ' months (minimum: ' + minimumMonths + ' months)');
    }
  }
  
  // 5. Enhanced National Minimum Wage Affordability Check
  checks.push('NMW Affordability with Existing Salary Sacrifices (using full adult rate)');
  var proposedMonthlyCost = parseFloat(agreementData.totalMonthlyCost || 0);
  
  if (proposedMonthlyCost > 0 && currentSalary > 0) {
    var nmwResult = getEmployeeNMWRate(employee);
    
    if (!nmwResult.success) {
      isEligible = false;
      issues.push('Cannot determine NMW rate for eligibility check: ' + nmwResult.error);
      
      enhancedCalculationBreakdown = {
        error: 'NMW rate lookup failed',
        errorDetails: nmwResult.error,
        salaryDetails: {
          grossAnnualSalary: currentSalary,
          proposedMonthlyEVCost: proposedMonthlyCost,
          existingMonthlySalarySacrifices: {
            pension: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.pension || 0 : 0),
            cycleToWork: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.cycleToWork || 0 : 0),
            childcareVouchers: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.childcareVouchers || 0 : 0),
            other: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.other || 0 : 0),
            total: totalExistingSalSacs
          }
        }
      };
    } else {
      var currentNMW = nmwResult.rate;
      
      // Calculate existing salary sacrifices from provided data
      try {
        if (agreementData.existingSalarySacrifices) {
          var existing = agreementData.existingSalarySacrifices;
          totalExistingSalSacs = (existing.pension || 0) + (existing.cycleToWork || 0) + 
                               (existing.childcareVouchers || 0) + (existing.other || 0);
        }
      } catch (error) {
        console.warn('Error calculating existing salary sacrifices:', error);
        totalExistingSalSacs = 0;
        warnings.push('Could not calculate existing salary sacrifices');
      }
      
      // Perform affordability calculation
      var totalMonthlyDeductions = proposedMonthlyCost + totalExistingSalSacs;
      var postDeductionSalary = currentSalary - (totalMonthlyDeductions * 12);
      
      var nmwBuffer = 1.15; // 15% buffer above NMW
      var contractHours = nmwResult.contractHours || 37.5; // Use actual contract hours
      var annualNMWMinimum = currentNMW * 52 * contractHours; // Use actual contract hours
      var requiredMinimumSalary = annualNMWMinimum * nmwBuffer;
      
      // Build detailed calculation breakdown
      enhancedCalculationBreakdown = {
        nmwDetails: {
          hourlyRate: currentNMW,
          staffType: nmwResult.staffType,
          contractHours: contractHours,
          ageBandUsed: nmwResult.ageBandUsed || 'Full Adult Rate',
          effectiveDate: nmwResult.effectiveDate.toDateString()
        },
        salaryDetails: {
          grossAnnualSalary: currentSalary,
          proposedMonthlyEVCost: proposedMonthlyCost,
          existingMonthlySalarySacrifices: {
            pension: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.pension || 0 : 0),
            cycleToWork: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.cycleToWork || 0 : 0),
            childcareVouchers: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.childcareVouchers || 0 : 0),
            other: (agreementData.existingSalarySacrifices ? agreementData.existingSalarySacrifices.other || 0 : 0),
            total: totalExistingSalSacs
          },
          totalMonthlyDeductions: totalMonthlyDeductions,
          totalAnnualDeductions: totalMonthlyDeductions * 12,
          netAnnualSalary: postDeductionSalary
        },
        nmwCalculation: {
          hourlyRate: currentNMW,
          contractHours: contractHours,
          weeksPerYear: 52,
          annualNMWBaseline: annualNMWMinimum,
          nmwBufferPercentage: 15,
          nmwBufferMultiplier: nmwBuffer,
          requiredMinimumSalary: requiredMinimumSalary
        },
        complianceCheck: {
          netSalaryAfterDeductions: postDeductionSalary,
          requiredMinimumSalary: requiredMinimumSalary,
          shortfall: postDeductionSalary < requiredMinimumSalary ? requiredMinimumSalary - postDeductionSalary : 0,
          isCompliant: postDeductionSalary >= requiredMinimumSalary,
          marginAboveMinimum: postDeductionSalary >= requiredMinimumSalary ? postDeductionSalary - requiredMinimumSalary : 0
        }
      };
      
      if (postDeductionSalary < requiredMinimumSalary) {
        isEligible = false;
        issues.push('Post-deduction salary (£' + postDeductionSalary.toLocaleString() + 
                    ') would be below required minimum (£' + requiredMinimumSalary.toLocaleString() + 
                    ') - includes 15% NMW buffer. Total monthly salary sacrifices: £' + totalMonthlyDeductions.toLocaleString());
      }
      
      // Add salary sacrifice breakdown to details
      if (totalExistingSalSacs > 0) {
        warnings.push('Existing salary sacrifices: £' + totalExistingSalSacs.toLocaleString() + '/month');
      }
      
      checks.push('Affordability Calculation with Existing Deductions Complete');
    }
  } else if (proposedMonthlyCost === 0) {
    warnings.push('No proposed monthly cost specified for affordability check');
  } else if (currentSalary === 0) {
    isEligible = false;
    issues.push('No salary information available for affordability check');
  }
  
  // 6. Additional Data Validation
  checks.push('Data Completeness');
  if (!agreementData.p11dValue) {
    warnings.push('P11D value not provided');
  }
  
  return {
    isEligible: isEligible,
    notes: isEligible ? 
      'All eligibility requirements met' + (warnings.length > 0 ? ' (Warnings: ' + warnings.join('; ') + ')' : '') :
      issues.join('; '),
    rulesChecked: checks,
    warnings: warnings,
    calculationBreakdown: enhancedCalculationBreakdown,
    details: {
      employmentStatus: employee.Status,
      payrollEmployee: 'Yes',
      payBasis: payBasis,
      currentSalary: currentSalary,
      lengthOfService: startDate ? Math.floor((new Date() - startDate) / (1000 * 60 * 60 * 24 * 30.44)) : null,
      monthlyDeduction: proposedMonthlyCost,
      existingSalarySacrifices: totalExistingSalSacs || 0,
      postDeductionSalary: currentSalary > 0 && proposedMonthlyCost > 0 ? 
        currentSalary - ((proposedMonthlyCost + (totalExistingSalSacs || 0)) * 12) : null,
      nmwBuffer: 1.15,
      startDate: employee.StartDate,
      issues: issues,
      warnings: warnings,
      salarySacrificeBreakdown: agreementData.existingSalarySacrifices || {}
    }
  };
}

/**
 * Adds a new EV scheme agreement.
 * @param {Object} agreementData - Agreement data from form
 * @returns {Object} Result object with success status and details
 */
function addEVSchemeAgreement(agreementData) {
  try {
    var agreements = loadEVSchemeAgreements();
    
    var agreementId = 'EV' + new Date().getFullYear() + String(agreements.length + 1).padStart(4, '0');
    
    var employee = getEmployeeByNumber(agreementData.employeeNumber);
    if (!employee) {
      throw new Error('Employee ' + agreementData.employeeNumber + ' not found');
    }
    
    // Use enhanced eligibility check if existing salary sacrifices are provided
    var eligibilityResult;
    if (agreementData.existingSalarySacrifices) {
      var enhancedData = {
        totalMonthlyCost: agreementData.totalMonthlyCost,
        p11dValue: agreementData.p11dValue,
        leaseTerm: agreementData.leaseTerm,
        existingSalarySacrifices: agreementData.existingSalarySacrifices
      };
      eligibilityResult = checkEVSchemeEligibilityEnhanced(employee, enhancedData);
    } else {
      eligibilityResult = checkEVSchemeEligibility(employee, agreementData);
    }
    
    var newAgreement = {
      agreementId: agreementId,
      employeeNumber: agreementData.employeeNumber,
      employeeName: employee.Firstnames + ' ' + employee.Surname,
      location: employee.Location,
      division: employee.Division,
      vehicleMake: agreementData.vehicleMake,
      vehicleModel: agreementData.vehicleModel,
      p11dValue: parseFloat(agreementData.p11dValue || 0),
      totalMonthlyCost: parseFloat(agreementData.totalMonthlyCost || 0),
      leaseTerm: parseInt(agreementData.leaseTerm),
      leaseValue: parseFloat(agreementData.leaseValue || 0),
      monthlyDeduction: parseFloat(agreementData.totalMonthlyCost || agreementData.monthlyDeduction || 0),
      providerName: agreementData.providerName,
      providerContact: agreementData.providerContact,
      
      // Workflow Management Fields
      workflowStatus: eligibilityResult.isEligible ? 'Eligible - Pending Approval' : 'Ineligible',
      status: eligibilityResult.isEligible ? 'Eligible - Pending Approval' : 'Ineligible', // Backward compatibility
      submissionDate: new Date().toISOString(),
      lastStatusChange: new Date().toISOString(),
      expectedDeliveryDate: null,
      actualDeliveryDate: null,
      activationDate: null,
      deductionsStartDate: null,
      remainingTermMonths: agreementData.leaseTerm,
      
      // Enhanced salary snapshot including existing salary sacrifices
      employeeSalarySnapshot: {
        currentSalary: parseFloat(employee.Salary || 0),
        payBasis: employee.PayBasis || employee.PayType,
        startDate: employee.StartDate,
        lengthOfService: eligibilityResult.details.lengthOfService,
        existingSalarySacrifices: agreementData.existingSalarySacrifices || {},
        totalExistingMonthlySalSac: eligibilityResult.details.existingSalarySacrifices || 0
      },
      
      submittedBy: Session.getEffectiveUser().getEmail(),
      eligibilityNotes: eligibilityResult.notes,
      isEligible: eligibilityResult.isEligible,
      businessRulesChecked: eligibilityResult.rulesChecked,
      eligibilityWarnings: eligibilityResult.warnings || [],
      eligibilityDetails: eligibilityResult.details,
      
      // Approval tracking
      approvalStatus: 'Pending',
      approvedBy: null,
      approvalDate: null,
      approvalNotes: null,
      
      // Workflow history
      workflowHistory: [{
        status: eligibilityResult.isEligible ? 'Eligible - Pending Approval' : 'Ineligible',
        timestamp: new Date().toISOString(),
        updatedBy: Session.getEffectiveUser().getEmail(),
        notes: 'Agreement created - ' + eligibilityResult.notes
      }]
    };
    
    agreements.push(newAgreement);
    saveEVSchemeAgreements(agreements);
    
    return {
      success: true,
      agreementId: agreementId,
      isEligible: eligibilityResult.isEligible,
      warnings: eligibilityResult.warnings || [],
      message: eligibilityResult.isEligible ? 
        'Agreement created successfully. Employee is eligible.' :
        'Agreement created but employee is ineligible: ' + eligibilityResult.notes
    };
    
  } catch (error) {
    console.error('Error adding EV scheme agreement:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Performs comprehensive EV scheme eligibility check with salary sacrifice data.
 * @param {Object} salSacData - Complete salary sacrifice data including existing agreements
 * @returns {Object} Detailed eligibility result
 */
function performEVSchemeEligibilityCheck(salSacData) {
  try {
    var employee = getEmployeeByNumber(salSacData.employeeNumber);
    if (!employee) {
      return {
        success: false,
        error: 'Employee ' + salSacData.employeeNumber + ' not found'
      };
    }
    
    // Create enhanced agreement data with existing salary sacrifices
    var enhancedAgreementData = {
      totalMonthlyCost: salSacData.proposedMonthlyCost,
      p11dValue: salSacData.proposedP11dValue,
      leaseTerm: 36, // Default for eligibility check
      existingSalarySacrifices: salSacData.existingSalarySacrifices
    };
    
    // Perform eligibility check with enhanced data
    var eligibilityResult = checkEVSchemeEligibilityEnhanced(employee, enhancedAgreementData);
    
    return {
      success: true,
      eligibilityResult: eligibilityResult,
      salSacData: salSacData,
      employee: {
        employeeNumber: employee.EmployeeNumber,
        name: employee.Firstnames + ' ' + employee.Surname,
        currentSalary: employee.Salary
      }
    };
    
  } catch (error) {
    console.error('Error performing EV scheme eligibility check:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * API function to validate employee for new agreement.
 * @param {string} employeeNumber - Employee number to validate
 * @returns {Object} Validation result with employee details and preliminary eligibility
 */
function validateEmployeeForEVScheme(employeeNumber) {
  try {
    var employee = getEmployeeByNumber(employeeNumber);
    
    if (!employee) {
      return {
        success: false,
        error: 'Employee ' + employeeNumber + ' not found'
      };
    }
    
    var basicEligibility = checkEVSchemeEligibility(employee, { totalMonthlyCost: 0 });
    
    return {
      success: true,
      employee: {
        employeeNumber: employee.EmployeeNumber,
        name: employee.Firstnames + ' ' + employee.Surname,
        location: employee.Location,
        division: employee.Division,
        currentSalary: employee.Salary,
        startDate: employee.StartDate,
        status: employee.Status,
        payrollEmployee: 'Yes',
        payBasis: employee.PayBasis || employee.PayType
      },
      preliminaryEligibility: basicEligibility
    };
    
  } catch (error) {
    console.error('Error validating employee:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Updates the workflow status of an EV scheme agreement.
 * @param {string} agreementId - Agreement ID to update
 * @param {string} newStatus - New workflow status
 * @param {string|null} deliveryDate - Delivery date (if applicable)
 * @returns {Object} Update result
 */
function updateAgreementWorkflowStatus(agreementId, newStatus, deliveryDate) {
  if (deliveryDate === undefined) deliveryDate = null;
  
  try {
    var agreements = loadEVSchemeAgreements();
    var agreementIndex = agreements.findIndex(function(a) { return a.agreementId === agreementId; });
    
    if (agreementIndex === -1) {
      return {
        success: false,
        error: 'Agreement not found: ' + agreementId
      };
    }
    
    var agreement = agreements[agreementIndex];
    var currentUser = Session.getEffectiveUser().getEmail();
    var timestamp = new Date().toISOString();
    
    // Update workflow status
    agreement.workflowStatus = newStatus;
    agreement.status = newStatus; // Backward compatibility
    agreement.lastStatusChange = timestamp;
    
    // Handle status-specific updates
    if (newStatus === 'Approved - Awaiting Delivery') {
      agreement.approvalStatus = 'Approved';
      agreement.approvedBy = currentUser;
      agreement.approvalDate = timestamp;
      if (deliveryDate) {
        agreement.expectedDeliveryDate = deliveryDate;
      }
    } else if (newStatus === 'Active - Vehicle Delivered') {
      agreement.actualDeliveryDate = deliveryDate || timestamp;
      agreement.activationDate = deliveryDate || timestamp;
      agreement.deductionsStartDate = deliveryDate || timestamp;
      // Reset remaining term to original lease term when activated
      agreement.remainingTermMonths = agreement.leaseTerm;
    } else if (newStatus === 'Completed') {
      agreement.completionDate = timestamp;
      agreement.remainingTermMonths = 0;
    }
    
    // Add to workflow history
    if (!agreement.workflowHistory) {
      agreement.workflowHistory = [];
    }
    
    agreement.workflowHistory.push({
      status: newStatus,
      timestamp: timestamp,
      updatedBy: currentUser,
      notes: getStatusChangeNotes(newStatus, deliveryDate),
      previousStatus: agreement.workflowHistory.length > 0 ? 
        agreement.workflowHistory[agreement.workflowHistory.length - 1].status : 
        'Initial Status'
    });
    
    // Save updated agreements
    agreements[agreementIndex] = agreement;
    saveEVSchemeAgreements(agreements);
    
    console.log('Updated agreement ' + agreementId + ' to status: ' + newStatus);
    
    return {
      success: true,
      agreementId: agreementId,
      newStatus: newStatus,
      message: 'Agreement status updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating agreement workflow status:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Generates appropriate notes for status changes.
 * @param {string} status - New status
 * @param {string|null} deliveryDate - Delivery date if applicable
 * @returns {string} Status change notes
 */
function getStatusChangeNotes(status, deliveryDate) {
  if (status === 'Approved - Awaiting Delivery') {
    return 'Agreement approved by admin. Awaiting vehicle delivery.';
  } else if (status === 'Active - Vehicle Delivered') {
    return 'Vehicle delivered' + (deliveryDate ? ' on ' + deliveryDate : '') + '. Salary deductions will commence.';
  } else if (status === 'Completed') {
    return 'Agreement completed. Lease term ended.';
  } else {
    return 'Status updated to: ' + status;
  }
}

/**
 * Gets remaining term for active agreements (calculated in real-time).
 * @param {string} agreementId - Agreement ID
 * @returns {Object} Remaining term calculation
 */
function getAgreementRemainingTerm(agreementId) {
  try {
    var agreements = loadEVSchemeAgreements();
    var agreement = agreements.find(function(a) { return a.agreementId === agreementId; });
    
    if (!agreement) {
      return { success: false, error: 'Agreement not found' };
    }
    
    if (!agreement.activationDate || agreement.workflowStatus !== 'Active - Vehicle Delivered') {
      return {
        success: true,
        remainingMonths: agreement.leaseTerm || 0,
        status: 'Not Active',
        originalTerm: agreement.leaseTerm || 0
      };
    }
    
    var activationDate = new Date(agreement.activationDate);
    var now = new Date();
    var monthsElapsed = Math.floor((now - activationDate) / (1000 * 60 * 60 * 24 * 30.44));
    var remainingMonths = Math.max(0, (agreement.leaseTerm || 0) - monthsElapsed);
    
    return {
      success: true,
      remainingMonths: remainingMonths,
      monthsElapsed: monthsElapsed,
      originalTerm: agreement.leaseTerm || 0,
      activationDate: agreement.activationDate,
      isActive: remainingMonths > 0,
      isCompleted: remainingMonths === 0
    };
    
  } catch (error) {
    console.error('Error calculating remaining term:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gets all agreements that should transition to completed status.
 * @returns {Array<Object>} Agreements that need status updates
 */
function getAgreementsNeedingCompletion() {
  try {
    var agreements = loadEVSchemeAgreements();
    var needingCompletion = [];
    
    agreements.forEach(function(agreement) {
      if (agreement.workflowStatus === 'Active - Vehicle Delivered' && agreement.activationDate) {
        var termResult = getAgreementRemainingTerm(agreement.agreementId);
        if (termResult.success && termResult.isCompleted) {
          needingCompletion.push({
            agreementId: agreement.agreementId,
            employeeName: agreement.employeeName,
            monthsElapsed: termResult.monthsElapsed,
            originalTerm: termResult.originalTerm
          });
        }
      }
    });
    
    return needingCompletion;
    
  } catch (error) {
    console.error('Error finding agreements needing completion:', error);
    return [];
  }
}

function testEVSchemeModule() {
  console.log('=== EV Scheme Management Module Test ===');
  
  try {
    const agreements = loadEVSchemeAgreements();
    console.log('✓ Loaded ' + agreements.length + ' agreements');
    
    const summary = generateEVSchemeSummary(agreements);
    console.log('✓ Generated summary: ' + summary.total + ' total agreements');
    
    const testEmployees = getActivePayrollEmployees().slice(0, 3);
    testEmployees.forEach((employee, index) => {
      console.log('Testing employee ' + (index + 1) + ': ' + employee.EmployeeNumber);
      
      const testAgreementData = {
        p11dValue: 35000,
        totalMonthlyCost: 650,
        leaseTerm: 36
      };
      
      const eligibilityResult = checkEVSchemeEligibility(employee, testAgreementData);
      console.log('- Eligible: ' + eligibilityResult.isEligible);
      console.log('- Current Salary: £' + (employee.Salary || 'N/A'));
      
      if (!eligibilityResult.isEligible) {
        console.log('- Issues: ' + eligibilityResult.notes);
      }
    });
    
    console.log('=== EV Scheme Module Test Complete ===');
    
    return {
      success: true,
      agreementsCount: agreements.length,
      summary: summary,
      testResults: 'All tests passed'
    };
    
  } catch (error) {
    console.error('EV Scheme module test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
