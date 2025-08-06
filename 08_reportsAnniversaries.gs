/**
 * Updated Anniversary Report with simple start/end date range
 * Replaces the month-only approach
 * FIXED: Timezone normalization issue - all dates normalized to midnight to ensure proper comparisons
 */

function generateMilestoneReport(data, options = {}) {
  // Use provided date range or prompt for one
  const dateRange = options.dateRange || promptForDateRange("Select Date Range for Anniversary Report");
  
  // NORMALIZE date range to midnight to avoid timezone comparison issues
  dateRange.startDate.setHours(0, 0, 0, 0);
  dateRange.endDate.setHours(23, 59, 59, 999);
  
  const milestones = [3, 5, 10, 15, 20, 25, 30, 35, 40];
  const today = new Date();
  const currentYear = today.getFullYear();
  
  const matches = data.filter(emp => {
    const rawDate = emp.ContinuousServiceStartDate?.trim();
    if (!rawDate || rawDate.length < 8) return false;

    const parsed = parseDMY(rawDate);
    if (!parsed) return false;

    // Calculate years of service
    const years = currentYear - parsed.getFullYear();
    
    // Check if this employee hits a milestone this year
    if (!milestones.includes(years)) return false;
    
    // Create the anniversary date for this year (normalize to midnight)
    const anniversaryDate = new Date(currentYear, parsed.getMonth(), parsed.getDate());
    anniversaryDate.setHours(0, 0, 0, 0);
    
    // Check if anniversary falls within selected date range
    return isDateInRange(anniversaryDate, dateRange);
  });

  // Sort by anniversary date within the range
  matches.sort((a, b) => {
    const dateA = parseDMY(a.ContinuousServiceStartDate);
    const dateB = parseDMY(b.ContinuousServiceStartDate);
    const annivA = new Date(currentYear, dateA.getMonth(), dateA.getDate());
    const annivB = new Date(currentYear, dateB.getMonth(), dateB.getDate());
    annivA.setHours(0, 0, 0, 0); // Normalize to midnight
    annivB.setHours(0, 0, 0, 0); // Normalize to midnight
    return annivA - annivB;
  });

  const header = ['EE Number', 'Full Name', 'Start Date', 'Anniversary Date', 'Years of Service', 'Location', 'Division'];
  
  const rows = matches.map(emp => {
    const parsed = parseDMY(emp.ContinuousServiceStartDate);
    const years = currentYear - parsed.getFullYear();
    const anniversaryDate = new Date(currentYear, parsed.getMonth(), parsed.getDate());
    anniversaryDate.setHours(0, 0, 0, 0); // Normalize to midnight
    const formattedAnniversary = Utilities.formatDate(anniversaryDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    
    return [
      emp.EmployeeNumber,
      `${emp.Firstnames} ${emp.Surname}`,
      emp.ContinuousServiceStartDate,
      formattedAnniversary,
      years,
      emp.Location,
      emp.Division
    ];
  });

  return [header, ...rows];
}
