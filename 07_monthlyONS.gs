/**
 * Gets monthly ONS summary: unique paid employees + total payment.
 *
 * @param {number} year - Calendar year
 * @param {number} fmMonth - Financial month (1 = April, ..., 12 = March)
 * @returns {{ employeeCount: number, totalPayment: number }}
 */
function getMonthlyONSReportSummary(year, fmMonth) {
  const calendarMonth = getCalendarMonthFromFM(fmMonth);
  const monthName = getMonthName(calendarMonth);
  const fm = fmMonth.toString().padStart(2, '0');
  const keyword = `${fm} ${monthName} ${year}`;

  const hourlyFolder = DriveApp.getFolderById(CONFIG.HOURLY_PAYMENT_REPORT_FOLDER_ID);
  const salariedFolder = DriveApp.getFolderById(CONFIG.SALARIED_PAYMENT_REPORT_FOLDER_ID);

  const hourlyFile = getCsvFileByKeyword(hourlyFolder, keyword);
  const salariedFile = getCsvFileByKeyword(salariedFolder, keyword);

  const hourlyData = parseCsvToObjects(hourlyFile);
  const salariedData = parseCsvToObjects(salariedFile);

  const combined = [...hourlyData, ...salariedData];

  const paidEmployees = combined.filter(row => {
    const pay = parseFloat(row["Total Payment"]);
    return row["Employee Number"] && !isNaN(pay) && pay > 0;
  });

  const uniqueEmployeeSet = new Set(paidEmployees.map(r => r["Employee Number"]));
  const totalPayment = paidEmployees.reduce((sum, row) => {
    return sum + parseFloat(row["Total Payment"] || 0);
  }, 0);

  return {
    employeeCount: uniqueEmployeeSet.size,
    totalPayment: parseFloat(totalPayment.toFixed(2))
  };
}

/**
 * Launches a modal with dropdowns for year and FM month selection.
 */
function showMonthlyONSEmployeeCount() {
  const html = HtmlService.createHtmlOutputFromFile('07_onsPickerModal.html')
    .setWidth(1200)   // max comfortable width
    .setHeight(700);  // near-full height
  SpreadsheetApp.getUi().showModalDialog(html, '📊 ONS Monthly Report');
}


/**
 * Called from client modal to return the report data.
 *
 * @param {number} year
 * @param {number} fmMonth
 * @returns {{success: boolean, message: string}}
 */
function runONSReport(year, fmMonth) {
  try {
    const summary = getMonthlyONSReportSummary(year, fmMonth);
    const calendarMonth = getCalendarMonthFromFM(fmMonth);
    const monthName = getMonthName(calendarMonth);
    const formattedPay = new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(summary.totalPayment);

    return {
      success: true,
      message: `📊 ONS Report for ${monthName} ${year}\n\n👥 Employees Paid: ${summary.employeeCount}\n💰 Total Payment: ${formattedPay}`
    };
  } catch (error) {
    return {
      success: false,
      message: `❌ Error: ${error.message}`
    };
  }
}

/**
 * Converts FM month (1–12, April–March) to calendar month (1–12, Jan–Dec)
 */
function getCalendarMonthFromFM(fmMonth) {
  return ((fmMonth + 2) % 12) + 1;
}

/**
 * Gets month name from calendar month index (1–12)
 */
function getMonthName(calendarMonth) {
  const names = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return names[calendarMonth - 1];
}
