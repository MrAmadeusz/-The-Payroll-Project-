/**
 * 08_reportBuilder.gs
 * Dynamic employee data view builder
 */

function showReportBuilderModal() {
  const html = HtmlService.createHtmlOutputFromFile('08_reportBuilderModal.html')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Custom Employee Report Builder');
}

/**
 * Returns all employee data for filtering UI setup.
 */
function getAllEmployees() {
  return loadAllEmployeeData(); // delegated to global cache-aware loader
}

/**
 * Generates a preview for the given filter and column options.
 */
function previewFilteredEmployeeView(options) {
  if (!options) throw new Error("Options object is null or undefined.");
  console.log("🔍 Filters:", JSON.stringify(options.filters));
  console.log("🧮 Columns:", options.columns);

  const all = getAllEmployees();
  const fullOutput = generateFilteredEmployeeView(all, options);
  return fullOutput.slice(0, 21); // header + top 20
}

/**
 * Executes full report and exports it
 */
function runEmployeeReportBuilder(options, mode = 'download') {
  const all = getAllEmployees();
  const output = generateFilteredEmployeeView(all, options);
  const reportName = 'Custom_Employee_View';

  const recipients = (mode === 'send') ? promptForRecipients() : [Session.getEffectiveUser().getEmail()];

  const file = writeReportToStandaloneWorkbookWithRecipients(reportName, output, recipients);

  SpreadsheetApp.getUi().alert(
    `✅ Custom report generated successfully.\n\nFile: ${file.getName()}\n\nAccess: ${recipients.join(', ')}`
  );
}

/**
 * Core filtering and formatting logic
 */
function generateFilteredEmployeeView(data, options = {}) {
  if (!options) throw new Error("Options object is null or undefined.");

  const filters = options.filters || {};
  const selectedColumns = options.columns || null;

  console.log("🔍 generateFilteredEmployeeView filters:", JSON.stringify(filters));
  console.log("🧮 Selected columns:", selectedColumns);

  const filtered = getEmployeesMatchingFilters(filters);

  const columns = selectedColumns && selectedColumns.length
    ? selectedColumns
    : Object.keys(filtered[0] || {});

  const output = [columns];
  filtered.forEach(row => {
    output.push(columns.map(col => row[col] || ''));
  });

  return output;
}

/**
 * Safe load from Drive (uses existing global method)
 */
function loadAllEmployeeData() {
  const folder = DriveApp.getFolderById(CONFIG.EMPLOYEE_MASTER_FOLDER_ID);
  const file = getCsvFileByKeyword(folder, 'All');
  return parseCsvToObjects(file);
}

/**
 * Register the builder as an interactive report in the dashboard.
 */
function extendReportRegistryWithBuilder() {
  REPORT_REGISTRY['reportBuilder'] = {
    label: '📋 Custom Employee View',
    description: 'Build a filtered employee view with column selection and export',
    rolesAllowed: ['admin', 'hr'],
    isInteractive: true
  };
}
/**
 * 08_reportBuilder.gs
 * Dynamic employee data view builder
 */

function showReportBuilderModal() {
  const html = HtmlService.createHtmlOutputFromFile('08_reportBuilderModal.html')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Custom Employee Report Builder');
}

/**
 * Returns all employee data for filtering UI setup.
 */
function getAllEmployees() {
  return loadAllEmployeeData(); // delegated to global cache-aware loader
}

/**
 * Generates a preview for the given filter and column options.
 */
function previewFilteredEmployeeView(options) {
  if (!options) throw new Error("Options object is null or undefined.");
  console.log("🔍 Filters:", JSON.stringify(options.filters));
  console.log("🧮 Columns:", options.columns);

  const all = getAllEmployees();
  const fullOutput = generateFilteredEmployeeView(all, options);
  return fullOutput.slice(0, 21); // header + top 20
}

/**
 * Executes full report and exports it
 */
function runEmployeeReportBuilder(options, mode = 'download') {
  const all = getAllEmployees();
  const output = generateFilteredEmployeeView(all, options);
  const reportName = 'Custom_Employee_View';

  const recipients = (mode === 'send') ? promptForRecipients() : [Session.getEffectiveUser().getEmail()];

  const file = writeReportToStandaloneWorkbookWithRecipients(reportName, output, recipients);

  SpreadsheetApp.getUi().alert(
    `✅ Custom report generated successfully.\n\nFile: ${file.getName()}\n\nAccess: ${recipients.join(', ')}`
  );
}

/**
 * Core filtering and formatting logic
 */
function generateFilteredEmployeeView(data, options = {}) {
  if (!options) throw new Error("Options object is null or undefined.");

  const filters = options.filters || {};
  const selectedColumns = options.columns || null;

  console.log("🔍 generateFilteredEmployeeView filters:", JSON.stringify(filters));
  console.log("🧮 Selected columns:", selectedColumns);

  const filtered = getEmployeesMatchingFilters(filters);

  const columns = selectedColumns && selectedColumns.length
    ? selectedColumns
    : Object.keys(filtered[0] || {});

  const output = [columns];
  filtered.forEach(row => {
    output.push(columns.map(col => row[col] || ''));
  });

  return output;
}

/**
 * Safe load from Drive (uses existing global method)
 */
function loadAllEmployeeData() {
  const folder = DriveApp.getFolderById(CONFIG.EMPLOYEE_MASTER_FOLDER_ID);
  const file = getCsvFileByKeyword(folder, 'All');
  return parseCsvToObjects(file);
}

/**
 * Register the builder as an interactive report in the dashboard.
 */
function extendReportRegistryWithBuilder() {
  REPORT_REGISTRY['reportBuilder'] = {
    label: '📋 Custom Employee View',
    description: 'Build a filtered employee view with column selection and export',
    rolesAllowed: ['admin', 'hr'],
    isInteractive: true
  };
}
