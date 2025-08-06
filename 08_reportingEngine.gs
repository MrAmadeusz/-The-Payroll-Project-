/**
 * 08_reportingEngine.gs
 * Core reporting engine: dispatch, registry, access control, and output routing
 */

function getUserRole(email) {
  if (CONFIG.ADMIN_EMAILS.includes(email)) return 'admin';
  if (CONFIG.HR_EMAILS.includes(email)) return 'hr';
  if (CONFIG.MANAGER_EMAILS.includes(email)) return 'manager';
  if (CONFIG.LEISURE_EMAILS.includes(email)) return 'leisure';
  return 'unauthorised';
}

function runReport(reportName, options = {}) {
  const email = Session.getEffectiveUser().getEmail();
  const report = REPORT_REGISTRY[reportName];

  if (!report) {
    throw new Error(`Unknown report: ${reportName}`);
  }

  const userRole = getUserRole(email);
  if (!report.rolesAllowed.includes(userRole)) {
    throw new Error(`Access denied. Your role (${userRole}) is not permitted to run "${reportName}".`);
  }

  const data = report.loadData(options);
  const output = report.generate(data, options);
  const file = writeReportToStandaloneWorkbook(reportName, output, email);

  SpreadsheetApp.getUi().alert(
    `✅ Report "${reportName}" has been generated.`,
    `A private Google Sheet has been created and shared only with you:\n\n${file.getUrl()}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Registry of available reports.
 * Each entry defines how to load and process the data.
 */

const REPORT_REGISTRY = {
  'monthlyONSEmployeeCount': {
    label: '📊 Monthly ONS Headcount',
    description: 'Headcount summary for ONS monthly submission',
    group: '📈 Strategic Reports',
    rolesAllowed: ['admin', 'hr'],
    loadData: () => getAllEmployees(),
    generate: (data) => generateONSReportTable(data),
    defaultFileName: 'ONS_Headcount',
    exportFormat: 'sheet'
  },

  'lengthOfServiceMilestones': {
    label: '🎖️ Length of Service Milestones',
    description: 'Employees hitting 3, 5, 10, etc. year milestones in selected month',
    group: '📋 HR Reports',
    rolesAllowed: ['admin', 'hr'],
    loadData: () => getAllEmployees().filter(e => e.Status === 'Current'),
    generate: (data, options) => generateMilestoneReport(data, options),
    defaultFileName: 'Length_of_Service_Milestones',
    exportFormat: 'sheet'
  },

  'weeklyStarters': {
    label: '🟢 Weekly Starters Report',
    description: 'New starters from previous week (with smart duplicate detection)',
    group: '📋 HR Reports',
    rolesAllowed: ['admin', 'hr'],
    loadData: () => getAllEmployees(),
    generate: (data, options) => generateStartersReport(data, options),
    defaultFileName: 'Weekly_Starters_Report',
    exportFormat: 'sheet',
    usesDateRange: true
  },

  'weeklyLeavers': {
    label: '🔴 Weekly Leavers Report',
    description: 'Employee leavers with smart duplicate detection',
    group: '📋 HR Reports',
    rolesAllowed: ['admin', 'hr'],
    loadData: () => getAllEmployees(),
    generate: (data, options) => generateLeaversReport(data, options),
    defaultFileName: 'Weekly_Leavers_Report',
    exportFormat: 'sheet',
    usesDateRange: true
  },

  'employeeDataExplorer': {
    label: '📋 Custom Employee View',
    description: 'Build a filtered employee view with column selection and export',
    rolesAllowed: ['admin', 'hr'],
    isInteractive: true 
  },

  'weeklyDBSCheck': {
    label: '🔍 DBS Check: Starters & Transfers',
    description: 'Flags employees newly added to GM/Nights/Leisure rotas',
    group: '📋 HR Reports',
    rolesAllowed: ['admin', 'hr'],
    loadData: () => getAllEmployees(),
    generate: (data, options) => generateDBSCheckReport(data, options),
    defaultFileName: 'Weekly_DBS_Check',
    exportFormat: 'sheet',
    usesDateRange: true
  },

  'missingNINumbers': {
    label: '🆔 Missing NI Numbers',
    description: 'Current employees without National Insurance numbers',
    group: '📋 HR Reports',
    rolesAllowed: ['admin', 'hr'],
    loadData: () => getAllEmployees(),
    generate: (data, options) => generateMissingNIReport(data, options),
    defaultFileName: 'Missing_NI_Numbers_Report',
    exportFormat: 'sheet'
  },

  'contractCheck': {
    label: '📋 Contract Type Compliance',
    description: 'Current employees with incorrect contract type classifications',
    group: '📋 HR Reports',
    rolesAllowed: ['admin', 'hr'],
    loadData: () => getAllEmployees(),
    generate: (data, options) => generateContractCheckReport(data, options),
    defaultFileName: 'Contract_Check_Report',
    exportFormat: 'sheet'
  },

  'weeklyCurrentEmployees': {
    label: 'Leisure Report - Current Employees',
    description: 'Snapshot of all current employees as of report date',
    group: '📋 HR Reports',
    rolesAllowed: ['admin', 'leisure'],
    emailGroups: ['admin', 'leisure'], 
    loadData: () => getAllEmployees(),
    generate: (data, options) => generateCurrentEmployeesReport(data, options),
    defaultFileName: 'Leisure_Active_Employees',
    exportFormat: 'sheet',
    usesDateRange: false
  },

  'hoursDemographic': {
    label: 'Hours Demographic Report',
    description: 'Gender and contract hours analysis for current employees (+/- 30 Hours)',
    group: '📈 Strategic Reports',
    rolesAllowed: ['admin'],
    loadData: () => getAllEmployees(), 
    generate: (data, options) => generateHoursDemographicReport(data, options),
    defaultFileName: 'Monthly_Hours_Demographic',
    exportFormat: 'sheet',
    usesDateRange: false
  },

  'leaversAOEReport': {
    label: '💰 Leavers AOE Report',
    description: 'Weekly leavers with outstanding payments (non-zero ATT values)',
    group: '📋 HR Reports',
    rolesAllowed: ['admin'],
    loadData: () => getAllEmployees(),
    generate: (data, options) => generateLeaversAOEReport(data, options),
    defaultFileName: 'Weekly_Leavers_AOE_Report',
    exportFormat: 'sheet',
    usesDateRange: true
  }

  // Add additional reports here...

}
/**
 * Creates a new private Google Sheet containing the report output.
 * Shares view/edit access only with the report runner.
 */
function writeReportToStandaloneWorkbook(reportName, output, userEmail) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const safeName = `Report - ${reportName} - ${userEmail.split('@')[0]} - ${timestamp}`;

  const newSpreadsheet = SpreadsheetApp.create(safeName);
  const sheet = newSpreadsheet.getSheets()[0];
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  const file = DriveApp.getFileById(newSpreadsheet.getId());
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  file.addEditor(userEmail);
  return file;
}

/**
 * Prompts user for recipient email addresses
 */
function promptForRecipients() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Send Report To Recipients',
    'Enter email addresses (separate multiple emails with commas):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error('Send cancelled by user.');
  }
  
  const emailText = response.getResponseText().trim();
  if (!emailText) {
    throw new Error('No email addresses entered.');
  }
  
  // Split by comma and clean up
  const emails = emailText.split(',').map(email => email.trim()).filter(email => email);
  
  // Basic email validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const invalidEmails = emails.filter(email => !emailRegex.test(email));
  
  if (invalidEmails.length > 0) {
    throw new Error(`Invalid email addresses: ${invalidEmails.join(', ')}`);
  }
  
  if (emails.length > 10) {
    throw new Error('Maximum 10 recipients allowed.');
  }
  
  return emails;
}

function runReportWithMode(reportKey, mode = 'download') {
  try {
    const email = Session.getEffectiveUser().getEmail();
    const report = REPORT_REGISTRY[reportKey];

    if (!report) throw new Error(`Report not found: ${reportKey}`);
    
    const role = getUserRole(email);
    if (!report.rolesAllowed.includes(role)) {
      throw new Error(`Access denied for role: ${role}`);
    }

    const data = report.loadData();
    const output = report.generate(data);
    
    let recipients = [];
    let message = '';
    
    if (mode === 'send') {
      // Only the specified recipients - creator must add themselves if they want a copy
      recipients = promptForRecipients();
      message = `Report sent to: ${recipients.join(', ')}`;
    } else {
      // Download mode - only the creator gets access
      recipients = [email];
      message = 'Report created for download';
    }
    
    const file = writeReportToStandaloneWorkbookWithRecipients(reportKey, output, recipients);
    
    return {
      success: true,
      url: file.getUrl(),
      fileName: file.getName(),
      message: message,
      mode: mode,
      recipients: recipients.length
    };
    
  } catch (error) {
    console.error(`Report generation failed for ${reportKey}:`, error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Updated function to create workbook with multiple recipients
 */
function writeReportToStandaloneWorkbookWithRecipients(reportName, output, recipients) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const reportConfig = REPORT_REGISTRY[reportName];
  const requesterEmail = Session.getEffectiveUser().getEmail();
  const safeName = `${reportConfig?.defaultFileName || reportName} - ${requesterEmail.split('@')[0]} - ${timestamp}`;

  const newSpreadsheet = SpreadsheetApp.create(safeName);
  const sheet = newSpreadsheet.getActiveSheet();
  
  // Write the data
  if (output && output.length > 0) {
    sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    
    // Apply basic formatting
    if (output.length > 1) {
      // Header row formatting
      const headerRange = sheet.getRange(1, 1, 1, output[0].length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, output[0].length);
      
      // Freeze header row
      sheet.setFrozenRows(1);
    }
  }
  
  // Set sharing and add recipients
  const file = DriveApp.getFileById(newSpreadsheet.getId());
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  
  // Add all recipients as editors
  recipients.forEach(email => {
    try {
      file.addEditor(email);
    } catch (error) {
      console.warn(`Could not add ${email} as editor: ${error.message}`);
    }
  });
  
  return file;
}

