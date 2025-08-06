/**
 * ONS Headcount Report Logic
 * Used by REPORT_REGISTRY['monthlyONSEmployeeCount']
 */

/**
 * Loads employee data for the ONS headcount report.
 * Assumes getAllEmployees() returns an array of objects.
 */
function loadAllEmployeesForONS() {
  return getAllEmployees(); // Existing global function
}

/**
 * Transforms employee data into a simple tabular output.
 * Output is a 2D array: [header, ...rows]
 */
function generateONSReportTable(data) {
  const header = ['EE Number', 'Full Name', 'Status', 'Location', 'Division'];
  const rows = data.map(emp => [
    emp.eeNumber,
    emp.fullName,
    emp.status,
    emp.location,
    emp.division
  ]);
  return [header, ...rows];
}
