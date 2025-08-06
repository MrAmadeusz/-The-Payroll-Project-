/**
 * Creates the main menu when the spreadsheet opens.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getEffectiveUser().getEmail();
  const isAdmin = CONFIG.ADMIN_EMAILS.includes(userEmail);

  // === Standard HR & Payroll menu (all users) ===
  const mainMenu = ui.createMenu('🏢 HR & Payroll');
  mainMenu.addItem('👥 Show Employee Stats', 'showEmployeeStats');
  mainMenu.addSeparator();
  mainMenu.addItem('💰 NMW Advanced Checker', 'showEnhancedNMWComplianceSidebar');
  mainMenu.addItem('📍 NMW Location Checker', 'showNMWLocationSidebar');
  mainMenu.addSeparator();
  mainMenu.addItem('🤱 Maternity Dashboard', 'showMaternityDashboard'); 
  mainMenu.addSeparator();
  mainMenu.addItem('➕ Submit Feature Request', 'showFeatureRequestSubmissionForm');
  mainMenu.addToUi();

  // === Admin-only tools in separate menu ===
  if (isAdmin) {
    const adminMenu = ui.createMenu('🔐 Admin Tools');
    adminMenu.addItem('🔄 File Comparison Tool', 'showFileComparisonTool');
    adminMenu.addItem('🔍 Pre-Payroll Validation', 'showPrePayrollValidationDashboard');
    adminMenu.addSeparator();
    adminMenu.addItem('Take Data Snapshot', 'saveWeeklyEmployeeSnapshot');
    adminMenu.addItem('📅 Show Pay Periods', 'showPayPeriods');
    adminMenu.addItem('⏰ Show Upcoming Deadlines', 'showUpcomingDeadlines');
    adminMenu.addSeparator();
    adminMenu.addItem('💡 Feature Request Dashboard', 'showFeatureRequestDashboard');
    adminMenu.addSeparator();
    adminMenu.addItem('➕ New Maternity Leave', 'showNewMaternityCaseForm'); 
    adminMenu.addItem('🧪 Test Maternity (Console)', 'testMaternityModuleConsole'); // Debug tool
    adminMenu.addItem('🗓️ Team Absence Planner', 'showTeamAbsencePlanner');
    adminMenu.addItem('🚗 EV Scheme Management', 'showEVSchemeDashboard');
    adminMenu.addSeparator();
    adminMenu.addItem('📚 Journal Export Dashboard', 'showJournalExportSidebar');
    adminMenu.addToUi();
  }

  // === Reporting Menu (Currently Admin-only. Future Access TBC)
  if (isAdmin) {
    const reportingMenu = ui.createMenu ('📊 Reporting');
    reportingMenu.addItem('Reports Dashboard', 'showReportsDashboard')
    reportingMenu.addItem('📊 Generate Monthly ONS Report', 'showMonthlyONSEmployeeCount');
    reportingMenu.addToUi();
  }
}

// === JOURNAL FUNCTIONS ===

/**
 * Shows the journal export sidebar for multi-journal processing
 */
function showJournalExportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('12_journalSidebar')
    .setTitle("Export Payroll Journals")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// === Employee Functions (using existing functions) ===

/**
 * Shows basic employee statistics.
 */
function showEmployeeStats() {
  try {
    const allEmployees = getAllEmployees();
    const activePayroll = getActivePayrollEmployees();
    const locations = listUniqueLocations();
    const divisions = listUniqueDivisions();
    const payTypes = listUniquePayTypes();
    
    const stats = `📊 EMPLOYEE STATISTICS
═══════════════════════

👥 Total Employees: ${allEmployees.length}
💼 Active Payroll: ${activePayroll.length}
📍 Locations: ${locations.length}
🏢 Divisions: ${divisions.length}
💰 Pay Types: ${payTypes.join(', ')}

Top 3 Locations:
${locations.slice(0, 3).map(loc => `• ${loc}`).join('\n')}`;
    
    SpreadsheetApp.getUi().alert('Employee Statistics', stats, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', `Failed to load employee data: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === Payroll Functions (using existing functions) ===

/**
 * Shows current pay periods.
 */
function showPayPeriods() {
  try {
    const hourlyPeriod = getCurrentPayPeriod('Hourly');
    const salariedPeriod = getCurrentPayPeriod('Salaried');
    
    let info = '📅 CURRENT PAY PERIODS\n══════════════════════\n\n';
    
    if (hourlyPeriod) {
      info += `⏰ Hourly Staff:\n`;
      info += `   Period: ${hourlyPeriod.periodStart.toLocaleDateString()} - ${hourlyPeriod.periodEnd.toLocaleDateString()}\n`;
      info += `   Pay Date: ${hourlyPeriod.payDate.toLocaleDateString()}\n\n`;
    }
    
    if (salariedPeriod) {
      info += `💼 Salaried Staff:\n`;
      info += `   Period: ${salariedPeriod.periodStart.toLocaleDateString()} - ${salariedPeriod.periodEnd.toLocaleDateString()}\n`;
      info += `   Pay Date: ${salariedPeriod.payDate.toLocaleDateString()}\n`;
    }
    
    SpreadsheetApp.getUi().alert('Current Pay Periods', info, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', `Failed to load pay periods: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows upcoming payroll deadlines.
 */
function showUpcomingDeadlines() {
  try {
    const deadlines = getUpcomingPayrollDeadlines(30);
    
    let info = '⏰ UPCOMING DEADLINES (30 Days)\n═══════════════════════════════\n\n';
    
    if (deadlines.length === 0) {
      info += 'No upcoming deadlines in the next 30 days.';
    } else {
      deadlines.forEach(deadline => {
        const icon = deadline.type === 'cutoff' ? '✂️' : '💰';
        info += `${icon} ${deadline.date.toLocaleDateString()} - ${deadline.type}\n`;
        info += `   ${deadline.period.staffType} (${deadline.daysUntil} days)\n\n`;
      });
    }
    
    SpreadsheetApp.getUi().alert('Upcoming Deadlines', info, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', `Failed to load deadlines: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Simple payroll calendar test.
 */
function testPayrollCalendar() {
  try {
    const calendar = getPayrollCalendar();
    const validation = validateEmployeePayTypeConsistency();
    
    let result = '🧪 PAYROLL CALENDAR TEST\n═══════════════════════\n\n';
    result += `✅ Calendar loaded: ${calendar.length} periods\n`;
    result += `${validation.isValid ? '✅' : '❌'} Employee consistency: ${validation.summary}\n`;
    
    SpreadsheetApp.getUi().alert('Payroll Test Results', result, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Test Failed', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === NMW Functions (using existing functions) ===

/**
 * Shows NMW system health.
 */
function showNMWSystemHealth() {
  try {
    const health = getNMWSystemHealth();
    
    let status = `🏥 NMW SYSTEM HEALTH\n═══════════════════\n\nOverall: ${health.overall.toUpperCase()}\n\n`;
    
    Object.entries(health.components).forEach(([component, info]) => {
      const icon = info.status === 'healthy' ? '✅' : info.status === 'warning' ? '⚠️' : '❌';
      status += `${icon} ${component}: ${info.message}\n`;
    });
    
    SpreadsheetApp.getUi().alert('NMW System Health', status, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Health Check Failed', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === Feature Request Functions (Module 05) ===

/**
 * Shows feature request analytics in a popup.
 */
function showFeatureRequestAnalytics() {
  try {
    const analytics = getFeatureRequestAnalytics();
    
    let report = `📊 FEATURE REQUEST ANALYTICS\n═══════════════════════════\n\n`;
    
    // Overview
    report += `📈 OVERVIEW:\n`;
    report += `• Total Requests: ${analytics.overview.total}\n`;
    report += `• Active: ${analytics.overview.active}\n`;
    report += `• Completed: ${analytics.overview.completed}\n`;
    report += `• Needing Attention: ${analytics.overview.needingAttention}\n\n`;
    
    // By Status
    report += `📋 BY STATUS:\n`;
    Object.entries(analytics.byStatus).forEach(([status, count]) => {
      if (count > 0) report += `• ${status}: ${count}\n`;
    });
    
    report += `\n🎯 BY PRIORITY:\n`;
    Object.entries(analytics.byPriority).forEach(([priority, count]) => {
      if (count > 0) report += `• ${priority}: ${count}\n`;
    });
    
    report += `\n📂 BY CATEGORY:\n`;
    Object.entries(analytics.byCategory).forEach(([category, count]) => {
      if (count > 0) report += `• ${category}: ${count}\n`;
    });
    
    SpreadsheetApp.getUi().alert('Feature Request Analytics', report, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Analytics Error', `Failed to generate analytics: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Shows feature request system health.
 */
function showFeatureRequestSystemHealth() {
  try {
    const health = getFeatureRequestSystemHealth();
    
    let status = `🏥 FEATURE REQUEST SYSTEM HEALTH\n═════════════════════════════\n\nOverall: ${health.overall.toUpperCase()}\n\n`;
    
    Object.entries(health.components).forEach(([component, info]) => {
      const icon = info.status === 'healthy' ? '✅' : info.status === 'warning' ? '⚠️' : '❌';
      status += `${icon} ${component}: ${info.message}\n`;
    });
    
    status += `\nLast Checked: ${health.lastChecked.toLocaleString()}`;
    
    SpreadsheetApp.getUi().alert('Feature Request System Health', status, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Health Check Failed', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === Helper Functions ===

/**
 * Shows simple help information.
 */
function showQuickHelp() {
  const help = `🆘 QUICK HELP
════════════

🏢 AVAILABLE FUNCTIONS:

👥 EMPLOYEES:
• Show Employee Stats - Basic statistics
• Test Employee Directory - Run diagnostics

📅 PAYROLL:
• Show Pay Periods - Current periods
• Show Upcoming Deadlines - Next 30 days
• Test Payroll Calendar - System check

📚 JOURNALS:
• Journal Export Dashboard - Multi-journal interface
  - Investment Journal - Process investment payroll
  - Hourly Journal - Process hourly staff + tips
  - Salaried Journal - Process salaried staff + tips
  - Hourly Accrual - Copy accrual files
  - Cross Charge Journal - Inter-hotel transfers

💰 NMW COMPLIANCE:
• NMW Advanced Checker - Full filtering UI
• NMW Location Checker - Simple location check
• Test NMW System - Run diagnostics
• NMW System Health - Component status

🤱 MATERNITY MANAGEMENT:
• Maternity Dashboard - View all active cases
• New Maternity Leave - Create new case (Admin)
• Test Maternity Module - Run diagnostics (Admin)

💡 FEATURE REQUESTS:
• Feature Request Dashboard - Manage all requests
• Submit Feature Request - Create new request
• Feature Request Analytics - View statistics
• Test Feature Requests - Run diagnostics
• Feature Request Health - System status

🔧 TROUBLESHOOTING:
• Run the test functions if something breaks
• Check system health for component status
• Reset employee cache if data seems stale`;
  
  SpreadsheetApp.getUi().alert('Quick Help', help, SpreadsheetApp.getUi().ButtonSet.OK);
}
