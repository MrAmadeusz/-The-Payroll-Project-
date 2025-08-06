function showReportsDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('08_reportsDashboard')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Reports Dashboard');
}

function getAvailableReportsForUser() {
  const email = Session.getEffectiveUser().getEmail();
  const role = getUserRole(email);

  const groups = {
    '📋 HR Reports': [],
    '📈 Strategic Reports': [],
    '💸 Finance Reports': []
  };

  Object.entries(REPORT_REGISTRY).forEach(([key, config]) => {
    if (!config.rolesAllowed.includes(role)) return;

    const group = config.group || '📋 HR Reports'; // default group
    groups[group].push({
      key,
      name: prettifyReportName(key),
      description: config.description,
      isInteractive: !!config.isInteractive
    });
  });

  return groups;
}

function runReportAndReturnLink(reportKey) {
  const email = Session.getEffectiveUser().getEmail();
  const report = REPORT_REGISTRY[reportKey];

  if (!report) throw new Error(`Report not found: ${reportKey}`);
  const role = getUserRole(email);

  if (!report.rolesAllowed.includes(role)) {
    throw new Error(`Access denied for role: ${role}`);
  }

  const data = report.loadData();
  const output = report.generate(data);
  const file = writeReportToStandaloneWorkbook(reportKey, output, email);
  return file.getUrl();
}

function prettifyReportName(key) {
  return key
    .replace(/([A-Z])/g, ' $1')
    .replace(/^./, s => s.toUpperCase())
    .replace('Ons', 'ONS')
    .trim();
}

/**
 * Smart email recipient function - uses CONFIG groups with overrides
 */
function getReportRecipients(reportKey) {
  const config = CONFIG.REPORT_EMAIL_CONFIG;
  
  // For reports that need multiple groups, we can specify them in the registry
  const report = REPORT_REGISTRY[reportKey];
  if (report && report.emailGroups) {
    let recipients = [];
    report.emailGroups.forEach(group => {
      switch (group) {
        case 'admin':
          recipients.push(...(CONFIG.ADMIN_EMAILS || []));
          break;
        case 'hr':
          recipients.push(...(CONFIG.HR_EMAILS || []));
          break;
        case 'leisure':
          recipients.push(...(CONFIG.LEISURE_EMAILS || []));
          break;
        case 'manager':
          recipients.push(...(CONFIG.MANAGER_EMAILS || []));
          break;
      }
    });
    
    // Add any additional recipients
    const additional = config.additionalRecipients?.[reportKey] || [];
    recipients.push(...additional);
    
    // Remove duplicates
    const uniqueRecipients = [...new Set(recipients)];
    console.log(`📧 ${reportKey} recipients (${report.emailGroups.join(' + ')}):`, uniqueRecipients);
    return uniqueRecipients;
  }
  
  // Fall back to existing single-group logic for other reports
  const targetGroup = config.groupOverrides?.[reportKey] || config.defaultGroup;
  
  let recipients = [];
  switch (targetGroup) {
    case 'admin':
      recipients = [...(CONFIG.ADMIN_EMAILS || [])];
      break;
    case 'hr':
      recipients = [...(CONFIG.HR_EMAILS || [])];
      break;
    case 'both':
      recipients = [...(CONFIG.ADMIN_EMAILS || []), ...(CONFIG.HR_EMAILS || [])];
      break;
    case 'manager':
      recipients = [...(CONFIG.MANAGER_EMAILS || [])];
      break;
    case 'leisure':
      recipients = [...(CONFIG.LEISURE_EMAILS || [])];
      break;
    default:
      console.warn(`Unknown group: ${targetGroup}, defaulting to admin`);
      recipients = [...(CONFIG.ADMIN_EMAILS || [])];
  }
  
  // Add any additional recipients
  const additional = config.additionalRecipients?.[reportKey] || [];
  recipients.push(...additional);
  
  // Remove duplicates
  const uniqueRecipients = [...new Set(recipients)];
  console.log(`📧 ${reportKey} recipients (${targetGroup}):`, uniqueRecipients);
  return uniqueRecipients;
}

/**
 * Main automation function - called by GUI trigger
 * Set trigger to: Time-driven → Week timer → Monday → 9am-10am
 */
function runWeeklyReportsAutomated() {
  try {
    console.log('🤖 Starting automated weekly reports...');
    
    // Calculate previous week's date range (Monday to Sunday)
    const dateRange = getPreviousWeekDateRange();
    console.log(`📅 Generating reports for: ${formatDateRangeForDisplay(dateRange)}`);
    
    const results = [];
    
    // Generate DBS Check Report
    try {
      console.log('🔍 Generating DBS Check Report...');
      const dbsResult = generateAndSendReport('weeklyDBSCheck', getReportRecipients('weeklyDBSCheck'), dateRange);
      results.push(`✅ DBS Check: ${dbsResult.dataRows} employees sent to ${dbsResult.recipients} recipients`);
    } catch (error) {
      console.error('❌ DBS Check failed:', error.message);
      results.push(`❌ DBS Check failed: ${error.message}`);
    }
    
    // Generate Starters Report
    try {
      console.log('🟢 Generating Starters Report...');
      const startersResult = generateAndSendReport('weeklyStarters', getReportRecipients('weeklyStarters'), dateRange);
      results.push(`✅ Starters: ${startersResult.dataRows} employees sent to ${startersResult.recipients} recipients`);
    } catch (error) {
      console.error('❌ Starters Report failed:', error.message);
      results.push(`❌ Starters failed: ${error.message}`);
    }
    
    // Generate Leavers Report
    try {
      console.log('🔴 Generating Leavers Report...');
      const leaversResult = generateAndSendReport('weeklyLeavers', getReportRecipients('weeklyLeavers'), dateRange);
      results.push(`✅ Leavers: ${leaversResult.dataRows} employees sent to ${leaversResult.recipients} recipients`);
    } catch (error) {
      console.error('❌ Leavers Report failed:', error.message);
      results.push(`❌ Leavers failed: ${error.message}`);
    }
    
    // Generate Missing NI Numbers Report
    try {
      console.log('🆔 Generating Missing NI Numbers Report...');
      const missingNIResult = generateAndSendReport('missingNINumbers', getReportRecipients('missingNINumbers'));
      results.push(`✅ Missing NI: ${missingNIResult.dataRows} employees sent to ${missingNIResult.recipients} recipients`);
    } catch (error) {
      console.error('❌ Missing NI Report failed:', error.message);
      results.push(`❌ Missing NI failed: ${error.message}`);
    }
    
    // Generate Contract Check Report
    try {
      console.log('📋 Generating Contract Check Report...');
      const contractResult = generateAndSendReport('contractCheck', getReportRecipients('contractCheck'));
      results.push(`✅ Contract Check: ${contractResult.dataRows} employees sent to ${contractResult.recipients} recipients`);
    } catch (error) {
      console.error('❌ Contract Check failed:', error.message);
      results.push(`❌ Contract Check failed: ${error.message}`);
    }
    
    // Generate Current Employees Report
    try {
      console.log('📋 Generating Current Employees Report...');
      const currentEmpResult = generateAndSendReport('weeklyCurrentEmployees', getReportRecipients('weeklyCurrentEmployees'));
      results.push(`✅ Current Employees: ${currentEmpResult.dataRows} employees sent to ${currentEmpResult.recipients} recipients`);
    } catch (error) {
      console.error('❌ Current Employees Report failed:', error.message);
      results.push(`❌ Current Employees failed: ${error.message}`);
    }
    
    // Generate Leavers AOE Report
    try {
      console.log('💰 Generating Leavers AOE Report...');
      const aoeResult = generateAndSendReport('leaversAOEReport', getReportRecipients('leaversAOEReport'), dateRange);
      results.push(`✅ Leavers AOE: ${aoeResult.dataRows} employees sent to ${aoeResult.recipients} recipients`);
    } catch (error) {
      console.error('❌ Leavers AOE Report failed:', error.message);
      results.push(`❌ Leavers AOE failed: ${error.message}`);
    }
    
    // Send summary to admins
    sendSummaryEmail(results, dateRange);
    
    console.log('🎉 Weekly reports automation completed successfully');
    
  } catch (error) {
    console.error('💥 Critical automation error:', error);
    sendErrorEmail(error);
  }
}

/**
 * Generates and sends a single report
 */
function generateAndSendReport(reportKey, recipients, dateRange) {
  const report = REPORT_REGISTRY[reportKey];
  if (!report) throw new Error(`Report not found: ${reportKey}`);
  
  const data = report.loadData();
  const output = report.generate(data, { dateRange: dateRange });
  
  const file = writeReportToStandaloneWorkbookWithRecipients(reportKey, output, recipients);
  
  return {
    success: true,
    url: file.getUrl(),
    fileName: file.getName(),
    recipients: recipients.length,
    dataRows: Math.max(0, output.length - 3) // Minus headers
  };
}

/**
 * Calculates previous week's date range (Monday to Sunday)
 */
function getPreviousWeekDateRange() {
  const today = new Date();
  const currentDay = today.getDay(); // 0=Sunday, 1=Monday, etc.
  
  // Calculate days to subtract to get to previous Monday
  const daysToLastMonday = currentDay === 0 ? 8 : currentDay + 6;
  
  const lastMonday = new Date(today);
  lastMonday.setDate(today.getDate() - daysToLastMonday);
  lastMonday.setHours(0, 0, 0, 0);
  
  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);
  lastSunday.setHours(23, 59, 59, 999);
  
  return {
    startDate: lastMonday,
    endDate: lastSunday
  };
}

/**
 * Sends summary email to admins
 */
function sendSummaryEmail(results, dateRange) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  const dateRangeDisplay = formatDateRangeForDisplay(dateRange);
  
  const subject = `Weekly Reports Summary - ${dateRangeDisplay}`;
  const message = `📊 WEEKLY REPORTS COMPLETED\n\n` +
    `📅 Period: ${dateRangeDisplay}\n` +
    `🕒 Generated: ${new Date().toLocaleString()}\n\n` +
    `RESULTS:\n${results.join('\n')}\n\n` +
    `This is an automated summary from the weekly reports system.`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send summary to ${email}:`, error.message);
    }
  });
}

/**
 * Sends error notification if automation fails
 */
function sendErrorEmail(error) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  const subject = '🚨 Weekly Reports Automation FAILED';
  const message = `💥 WEEKLY REPORTS AUTOMATION ERROR\n\n` +
    `🕒 Time: ${new Date().toLocaleString()}\n` +
    `❌ Error: ${error.message}\n\n` +
    `Please check the Apps Script logs and run reports manually if needed.`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send error email to ${email}:`, error.message);
    }
  });
}

/**
 * Monthly Anniversary Report Automation
 * Run on 20th of each month to send next month's anniversaries
 * Set trigger to: Time-driven → Month timer → 20th day → 9am-10am
 */
function runMonthlyAnniversaryReportAutomated() {
  try {
    console.log('🎂 Starting automated monthly anniversary report...');
    
    // Calculate next month's date range
    const dateRange = getNextMonthDateRange();
    const monthName = getMonthName(dateRange.startDate.getMonth());
    
    console.log(`📅 Generating anniversary report for: ${monthName} ${dateRange.startDate.getFullYear()}`);
    console.log(`📅 Date range: ${formatDateRangeForDisplay(dateRange)}`);
    
    // Generate and send the report
    const result = generateAndSendReport(
      'lengthOfServiceMilestones', 
      getReportRecipients('lengthOfServiceMilestones'), 
      dateRange
    );
    
    // Send confirmation email to admins
    const summary = [
      `✅ Anniversary Report: ${result.dataRows} employees sent to ${result.recipients} recipients`,
      `📅 Period: ${monthName} ${dateRange.startDate.getFullYear()}`,
      `📊 Report: ${result.fileName}`,
      `🔗 URL: ${result.url}`
    ];
    
    sendAnniversarySummaryEmail(summary, dateRange);
    
    console.log('🎉 Monthly anniversary report automation completed successfully');
    
  } catch (error) {
    console.error('💥 Anniversary automation error:', error);
    sendAnniversaryErrorEmail(error);
  }
}

/**
 * Calculates the date range for the next month (1st to last day)
 */
function getNextMonthDateRange() {
  const today = new Date();
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  
  // Start of next month
  const startDate = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), 1);
  startDate.setHours(0, 0, 0, 0);
  
  // End of next month (last day)
  const endDate = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0);
  endDate.setHours(23, 59, 59, 999);
  
  return {
    startDate: startDate,
    endDate: endDate
  };
}

/**
 * Gets month name from month number (0-11)
 */
function getMonthName(monthNumber) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[monthNumber];
}

/**
 * Sends summary email for anniversary automation
 */
function sendAnniversarySummaryEmail(results, dateRange) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  const monthName = getMonthName(dateRange.startDate.getMonth());
  const year = dateRange.startDate.getFullYear();
  
  const subject = `Monthly Anniversary Report Sent - ${monthName} ${year}`;
  const message = `🎂 MONTHLY ANNIVERSARY REPORT COMPLETED\n\n` +
    `📅 Period: ${monthName} ${year}\n` +
    `🕒 Generated: ${new Date().toLocaleString()}\n\n` +
    `RESULTS:\n${results.join('\n')}\n\n` +
    `Recipients: ${getReportRecipients('lengthOfServiceMilestones').join(', ')}\n\n` +
    `This is an automated summary from the monthly anniversary report system.`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send anniversary summary to ${email}:`, error.message);
    }
  });
}

/**
 * Sends error notification if anniversary automation fails
 */
function sendAnniversaryErrorEmail(error) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  const subject = '🚨 Monthly Anniversary Report FAILED';
  const message = `💥 MONTHLY ANNIVERSARY REPORT ERROR\n\n` +
    `🕒 Time: ${new Date().toLocaleString()}\n` +
    `❌ Error: ${error.message}\n\n` +
    `Please check the Apps Script logs and run the report manually if needed.\n\n` +
    `Recipients configured: ${getReportRecipients('lengthOfServiceMilestones').join(', ')}`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send anniversary error email to ${email}:`, error.message);
    }
  });
}

/**
 * Test function to manually run anniversary automation for next month
 */
function testMonthlyAnniversaryAutomation() {
  console.log('🧪 Testing monthly anniversary automation...');
  runMonthlyAnniversaryReportAutomated();
}

/**
 * Test function to verify the new email configuration system
 */
function testNewEmailSystem() {
  console.log('📧 Testing new email configuration system...');
  
  const testReports = [
    'weeklyDBSCheck',
    'weeklyStarters', 
    'weeklyLeavers',
    'missingNINumbers',
    'contractCheck',
    'lengthOfServiceMilestones',
    'weeklyCurrentEmployees',
    'leaversAOEReport'  
  ];
  
  console.log('\n=== EMAIL DISTRIBUTION TEST ===');
  console.log(`Default group: ${CONFIG.REPORT_EMAIL_CONFIG?.defaultGroup || 'admin'}`);
  console.log(`Admin emails: ${CONFIG.ADMIN_EMAILS?.length || 0}`);
  console.log(`HR emails: ${CONFIG.HR_EMAILS?.length || 0}`);
  console.log(`Leisure emails: ${CONFIG.LEISURE_EMAILS?.length || 0}`);
  
  testReports.forEach(reportKey => {
    const recipients = getReportRecipients(reportKey);
    console.log(`\n📋 ${reportKey}:`);
    console.log(`   Recipients: ${recipients.length}`);
    console.log(`   Emails: ${recipients.join(', ')}`);
  });
  
  return {
    defaultGroup: CONFIG.REPORT_EMAIL_CONFIG?.defaultGroup,
    adminCount: CONFIG.ADMIN_EMAILS?.length,
    hrCount: CONFIG.HR_EMAILS?.length,
    leisureCount: CONFIG.LEISURE_EMAILS?.length,
    testResults: testReports.map(reportKey => ({
      report: reportKey,
      recipients: getReportRecipients(reportKey)
    }))
  };
}

/**
 * Test function to preview what weekly automation will generate
 */
function testWeeklyAutomationPreview() {
  console.log('🤖 WEEKLY AUTOMATION PREVIEW - NO EMAILS SENT');
  console.log('=' * 50);
  
  try {
    const dateRange = getPreviousWeekDateRange();
    console.log(`📅 Would generate reports for: ${formatDateRangeForDisplay(dateRange)}`);
    
    const weeklyReports = [
      'weeklyDBSCheck',
      'weeklyStarters',
      'weeklyLeavers', 
      'missingNINumbers',
      'contractCheck',
      'weeklyCurrentEmployees',
      'leaversAOEReport' 
    ];
    
    console.log('\n📊 REPORTS THAT WOULD BE GENERATED:');
    weeklyReports.forEach((reportKey, index) => {
      const recipients = getReportRecipients(reportKey);
      const report = REPORT_REGISTRY[reportKey];
      console.log(`${index + 1}. ${report?.label || reportKey}`);
      console.log(`   Recipients: ${recipients.length} (${recipients.join(', ')})`);
    });
    
    // Show total unique recipients
    const allRecipients = new Set();
    weeklyReports.forEach(reportKey => {
      getReportRecipients(reportKey).forEach(email => allRecipients.add(email));
    });
    
    console.log(`\n📧 TOTAL UNIQUE RECIPIENTS: ${allRecipients.size}`);
    console.log(`   ${Array.from(allRecipients).join(', ')}`);
    
    console.log('\n✅ This is a preview only - no reports generated, no emails sent');
    
    return {
      dateRange: dateRange,
      reportCount: weeklyReports.length,
      uniqueRecipients: allRecipients.size,
      recipients: Array.from(allRecipients)
    };
    
  } catch (error) {
    console.error('❌ Preview failed:', error.message);
    return { error: error.message };
  }
}

/**
 * Manual function to run weekly automation immediately (for testing)
 * WARNING: This will generate and send actual reports!
 */
function runWeeklyAutomationNow() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    'Confirm Weekly Automation',
    'This will generate and send ALL weekly reports immediately.\n\nAre you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation === ui.Button.YES) {
    console.log('🚀 Running weekly automation manually...');
    runWeeklyReportsAutomated();
  } else {
    console.log('❌ Weekly automation cancelled by user');
  }
}

/**
 * Monthly Hours Demographic Report Automation
 * Run on last day of each month (GUI trigger)
 * Set trigger to: Time-driven → Month timer → Last day → 9am-10am
 */
function runMonthlyHoursDemographicAutomated() {
  try {
    console.log('📊 Starting automated monthly hours demographic report...');
    
    const today = new Date();
    const currentMonth = getMonthName(today.getMonth());
    const currentYear = today.getFullYear();
    
    console.log(`📅 Generating hours demographic report for: ${currentMonth} ${currentYear}`);
    
    // Generate and send the report
    const result = generateAndSendReport(
      'hoursDemographic', 
      getReportRecipients('hoursDemographic')
    );
    
    // Send confirmation email to admins
    const summary = [
      `✅ Hours Demographic Report: ${result.dataRows} employees analyzed and sent to ${result.recipients} recipients`,
      `📅 Period: ${currentMonth} ${currentYear}`,
      `📊 Report: ${result.fileName}`,
      `🔗 URL: ${result.url}`
    ];
    
    sendHoursDemographicSummaryEmail(summary, currentMonth, currentYear);
    
    console.log('🎉 Monthly hours demographic report automation completed successfully');
    
  } catch (error) {
    console.error('💥 Hours demographic automation error:', error);
    sendHoursDemographicErrorEmail(error);
  }
}

/**
 * Sends summary email for hours demographic automation
 */
function sendHoursDemographicSummaryEmail(results, monthName, year) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  
  const subject = `Monthly Hours Demographic Report Sent - ${monthName} ${year}`;
  const message = `📊 MONTHLY HOURS DEMOGRAPHIC REPORT COMPLETED\n\n` +
    `📅 Period: ${monthName} ${year}\n` +
    `🕒 Generated: ${new Date().toLocaleString()}\n\n` +
    `RESULTS:\n${results.join('\n')}\n\n` +
    `Recipients: ${getReportRecipients('hoursDemographic').join(', ')}\n\n` +
    `This is an automated summary from the monthly hours demographic report system.`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send hours demographic summary to ${email}:`, error.message);
    }
  });
}

/**
 * Sends error notification if hours demographic automation fails
 */
function sendHoursDemographicErrorEmail(error) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  const subject = '🚨 Monthly Hours Demographic Report FAILED';
  const message = `💥 MONTHLY HOURS DEMOGRAPHIC REPORT ERROR\n\n` +
    `🕒 Time: ${new Date().toLocaleString()}\n` +
    `❌ Error: ${error.message}\n\n` +
    `Please check the Apps Script logs and run the report manually if needed.\n\n` +
    `Recipients configured: ${getReportRecipients('hoursDemographic').join(', ')}`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send hours demographic error email to ${email}:`, error.message);
    }
  });
}

/**
 * Test function to manually run hours demographic automation
 */
function testMonthlyHoursDemographicAutomation() {
  console.log('🧪 Testing monthly hours demographic automation...');
  runMonthlyHoursDemographicAutomated();
}

/**
 * Monthly Snapshot Automation - Run on 28th of each month
 * Set trigger to: Time-driven → Month timer → 28th day → 9am-10am
 */
function runMonthlySnapshotsAutomated() {
  try {
    console.log('📸 Starting automated monthly snapshots...');
    
    const today = new Date();
    const currentMonth = getMonthName(today.getMonth());
    const currentYear = today.getFullYear();
    
    console.log(`📅 Creating snapshots for: ${currentMonth} ${currentYear}`);
    
    const results = [];
    
    // Create Gross to Nett Snapshot
    try {
      console.log('💰 Creating Gross to Nett snapshot...');
      const grossResult = saveMonthlyGrossToNettSnapshot();
      results.push(`✅ Gross to Nett: ${grossResult.recordCount} records saved as ${grossResult.fileName}`);
    } catch (error) {
      console.error('❌ Gross to Nett snapshot failed:', error.message);
      results.push(`❌ Gross to Nett snapshot failed: ${error.message}`);
    }
    
    // Create Payments Snapshot
    try {
      console.log('💳 Creating Payments snapshot...');
      const paymentsResult = saveMonthlyPaymentsSnapshot();
      results.push(`✅ Payments: ${paymentsResult.recordCount} records saved as ${paymentsResult.fileName}`);
    } catch (error) {
      console.error('❌ Payments snapshot failed:', error.message);
      results.push(`❌ Payments snapshot failed: ${error.message}`);
    }
    
    // Send summary to admins
    sendSnapshotSummaryEmail(results, currentMonth, currentYear);
    
    console.log('🎉 Monthly snapshots automation completed successfully');
    
  } catch (error) {
    console.error('💥 Critical snapshot automation error:', error);
    sendSnapshotErrorEmail(error);
  }
}

/**
 * Sends summary email for snapshot automation
 */
function sendSnapshotSummaryEmail(results, monthName, year) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  
  const subject = `Monthly Snapshots Created - ${monthName} ${year}`;
  const message = `📸 MONTHLY SNAPSHOTS COMPLETED\n\n` +
    `📅 Period: ${monthName} ${year}\n` +
    `🕒 Generated: ${new Date().toLocaleString()}\n\n` +
    `RESULTS:\n${results.join('\n')}\n\n` +
    `These snapshots are used for the weekly Leavers AOE (Amount Owing Employee) report.\n\n` +
    `This is an automated summary from the monthly snapshot system.`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send snapshot summary to ${email}:`, error.message);
    }
  });
}

/**
 * Sends error notification if snapshot automation fails
 */
function sendSnapshotErrorEmail(error) {
  const adminEmails = CONFIG.ADMIN_EMAILS;
  const subject = '🚨 Monthly Snapshots Automation FAILED';
  const message = `💥 MONTHLY SNAPSHOTS AUTOMATION ERROR\n\n` +
    `🕒 Time: ${new Date().toLocaleString()}\n` +
    `❌ Error: ${error.message}\n\n` +
    `Please check the Apps Script logs and run snapshots manually if needed.\n\n` +
    `Note: This will affect the weekly Leavers AOE report until snapshots are created.`;
  
  adminEmails.forEach(email => {
    try {
      GmailApp.sendEmail(email, subject, message);
    } catch (error) {
      console.error(`Failed to send snapshot error email to ${email}:`, error.message);
    }
  });
}

/**
 * Test function to manually run snapshot automation
 */
function testMonthlySnapshotsAutomation() {
  console.log('🧪 Testing monthly snapshots automation...');
  runMonthlySnapshotsAutomated();
}

/**
 * Test the new Leavers AOE report integration
 */
function testLeaversAOEIntegration() {
  console.log('🧪 Testing Leavers AOE Report Integration...');
  
  try {
    // Check if snapshots exist
    console.log('1. Checking for required snapshots...');
    const grossData = loadLatestGrossToNettSnapshot();
    console.log(`   ✅ Gross to Nett snapshot: ${grossData.length} records`);
    
    // Test report generation
    console.log('2. Testing report generation...');
    const employees = getAllEmployees();
    const result = generateLeaversAOEReport(employees);
    console.log(`   ✅ Report generated: ${result.length} rows`);
    
    // Test registry integration
    console.log('3. Testing registry integration...');
    const reportConfig = REPORT_REGISTRY['leaversAOEReport'];
    if (reportConfig) {
      console.log(`   ✅ Registry entry found: ${reportConfig.label}`);
      console.log(`   ✅ Roles allowed: ${reportConfig.rolesAllowed.join(', ')}`);
    } else {
      console.warn('   ⚠️ Registry entry not found - add to REPORT_REGISTRY');
    }
    
    // Test recipient configuration
    console.log('4. Testing recipient configuration...');
    const recipients = getReportRecipients('leaversAOEReport');
    console.log(`   ✅ Recipients: ${recipients.length} (${recipients.join(', ')})`);
    
    console.log('\n✅ All integration tests passed');
    
    return {
      success: true,
      snapshotRecords: grossData.length,
      reportRows: result.length,
      recipients: recipients.length
    };
    
  } catch (error) {
    console.error('❌ Integration test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Complete system test - snapshots + report + automation
 */
function testCompleteAOESystem() {
  console.log('🔄 Testing Complete AOE System...');
  
  try {
    // Step 1: Test snapshot creation
    console.log('\n=== STEP 1: Testing Snapshot Creation ===');
    const snapshotResult = testMonthlySnapshots();
    if (!snapshotResult.grossToNett?.success) {
      throw new Error('Snapshot creation failed');
    }
    
    // Step 2: Test report generation
    console.log('\n=== STEP 2: Testing Report Generation ===');
    const reportResult = testLeaversAOEReport();
    if (!reportResult.success) {
      throw new Error('Report generation failed');
    }
    
    // Step 3: Test integration
    console.log('\n=== STEP 3: Testing Integration ===');
    const integrationResult = testLeaversAOEIntegration();
    if (!integrationResult.success) {
      throw new Error('Integration test failed');
    }
    
    console.log('\n🎉 COMPLETE AOE SYSTEM TEST PASSED');
    console.log(`📊 Summary:`);
    console.log(`   Snapshots: Gross(${snapshotResult.grossRecords}) + Payments(${snapshotResult.paymentsRecords})`);
    console.log(`   Report: ${reportResult.dataRows} leavers with outstanding payments`);
    console.log(`   Recipients: ${integrationResult.recipients} admin users`);
    
    return {
      success: true,
      snapshots: snapshotResult,
      report: reportResult,
      integration: integrationResult
    };
    
  } catch (error) {
    console.error('❌ Complete system test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}
