/**
 * 👥 employeeDirectory.gs
 * Employee data management module for HR/payroll system.
 * Provides cached, indexed access to employee master data with advanced filtering.
 */

// === Module-Level Cache Variables ===
let _employeeCache = null;
let _cachedFilterOptions = null;
let _employeeIndexes = {};

// === Cache Management ===

/**
 * Loads and caches the employee master CSV from Drive.
 * Implements performance optimizations including indexing and pre-computed filter arrays.
 * 
 * @returns {Array<Object>} Array of employee objects
 */
function getAllEmployees() {
  if (_employeeCache) return _employeeCache;

  _employeeCache = timeOperation("Loading employee data", () => {
    const folder = DriveApp.getFolderById(CONFIG.EMPLOYEE_MASTER_FOLDER_ID);
    const file = getCsvFileByKeyword(folder, "All");
    return parseCsvToObjects(file);
  });

  console.log(`Loaded ${_employeeCache.length} employees`);

  // Build performance optimizations
  _buildIndexesAndCaches();
  
  return _employeeCache;
}

/**
 * Builds indexes and pre-computes filter options for performance.
 * @private
 */
function _buildIndexesAndCaches() {
  timeOperation("Building indexes and filter caches", () => {
    // Build indexes for fast filtering
    _employeeIndexes = {
      Location: buildDataIndex(_employeeCache, 'Location'),
      Division: buildDataIndex(_employeeCache, 'Division'),
      PayType: buildDataIndex(_employeeCache, 'PayType'),
      Status: buildDataIndex(_employeeCache, 'Status')
    };

    // Pre-compute filter options
    _cachedFilterOptions = {
      location: extractUniqueValues(_employeeCache, 'Location'),
      division: extractUniqueValues(_employeeCache, 'Division'),
      jobTitle: extractUniqueValues(_employeeCache, 'JobTitle'),
      payType: extractUniqueValues(_employeeCache, 'PayType'),
      status: extractUniqueValues(_employeeCache, 'Status')
    };
  });

  console.log(`Built indexes for ${Object.keys(_employeeIndexes).length} fields`);
  console.log(`Cached filter options: ${JSON.stringify(Object.keys(_cachedFilterOptions).reduce((acc, key) => {
    acc[key] = _cachedFilterOptions[key].length;
    return acc;
  }, {}))}`);
}

/**
 * Clears all cached employee data and indexes.
 * Call this when employee data changes.
 */
function resetEmployeeCache() {
  _employeeCache = null;
  _cachedFilterOptions = null;
  _employeeIndexes = {};
  console.log("Employee cache cleared");
}

// === Core Lookup Functions ===

/**
 * Finds an employee by their employee number.
 * 
 * @param {string|number} employeeNumber - Employee number to search for
 * @returns {Object|null} Employee object or null if not found
 */
function getEmployeeByNumber(employeeNumber) {
  return getAllEmployees().find(emp => emp.EmployeeNumber == employeeNumber) || null;
}

/**
 * Gets employees by pay type (Salary/Hourly).
 * 
 * @param {string} payTypeInput - Pay type ("Salary", "Hourly", "salaried", "hourly", etc.)
 * @returns {Array<Object>} Matching employees
 */
function getEmployeesByPayType(payTypeInput) {
  const normalizedPayType = normalizePayType(payTypeInput);
  return getAllEmployees().filter(emp => emp.PayType === normalizedPayType);
}

/**
 * Gets current active payroll employees, optionally filtered by pay type.
 * 
 * @param {string} [payTypeInput] - Optional pay type filter
 * @returns {Array<Object>} Active payroll employees
 */
function getActivePayrollEmployees(payTypeInput = null) {
  const normalizedPayType = payTypeInput ? normalizePayType(payTypeInput) : null;

  return getAllEmployees().filter(emp => {
    const isCurrent = emp.Status?.toLowerCase() === "current";
    const isPayroll = emp.IsPayrollEmployee?.toLowerCase() === "yes";
    const payTypeMatch = normalizedPayType ? emp.PayType === normalizedPayType : true;
    return isCurrent && isPayroll && payTypeMatch;
  });
}

/**
 * Gets employees who were active on a specific date.
 * Considers start and termination dates.
 * 
 * @param {Date|string} targetDate - Date to check employment status
 * @param {string} [payTypeInput] - Optional pay type filter
 * @returns {Array<Object>} Employees active on the target date
 */
function getEmployeesActiveOn(targetDate, payTypeInput = null) {
  const normalizedPayType = payTypeInput ? normalizePayType(payTypeInput) : null;
  const date = new Date(targetDate);

  return getAllEmployees().filter(emp => {
    const isCurrent = emp.Status?.toLowerCase() === "current";
    const isPayroll = emp.IsPayrollEmployee?.toLowerCase() === "yes";
    const startDate = parseDMY(emp.StartDate);
    const endDate = parseDMY(emp.TerminationDate);
    const isEmployed = startDate && startDate <= date && (!endDate || date <= endDate);
    const payTypeMatch = normalizedPayType ? emp.PayType === normalizedPayType : true;
    return isCurrent && isPayroll && isEmployed && payTypeMatch;
  });
}

/**
 * Gets current employees at a specific location.
 * 
 * @param {string} locationName - Location name to filter by
 * @returns {Array<Object>} Current employees at the location
 */
function getCurrentEmployeesByLocation(locationName) {
  return getActivePayrollEmployees().filter(emp =>
    emp.Location?.toLowerCase() === locationName.toLowerCase()
  );
}

/**
 * Gets current employees in a specific division.
 * 
 * @param {string} divisionName - Division name to filter by
 * @returns {Array<Object>} Current employees in the division
 */
function getCurrentEmployeesByDivision(divisionName) {
  return getActivePayrollEmployees().filter(emp =>
    emp.Division?.toLowerCase() === divisionName.toLowerCase()
  );
}

// === Advanced Filtering ===

/**
 * Gets employees matching multiple filter criteria with optimized performance.
 * Uses pre-built indexes for fast filtering.
 * 
 * @param {Object} filters - Filter criteria {location: "...", division: "...", etc.}
 * @returns {Array<Object>} Matching employees
 */
function getEmployeesMatchingFilters(filters) {
  // Ensure data is loaded
  getAllEmployees();
  
  // Field mapping for UI keys to database fields
  const FIELD_MAP = {
    location: "Location",
    division: "Division",
    jobTitle: "JobTitle",
    payType: "PayType",
    status: "Status"
  };

  return timeOperation(`Filtering with ${Object.keys(filters).length} criteria`, () => {
    return filterDataWithIndexes(_employeeCache, filters, FIELD_MAP, _employeeIndexes);
  });
}

/**
 * Gets the count of employees matching filter criteria.
 * More efficient than getting full results when only count is needed.
 * 
 * @param {Object} filters - Filter criteria
 * @returns {number} Count of matching employees
 */
function getEmployeeMatchCount(filters) {
  return getEmployeesMatchingFilters(filters).length;
}

// === Filter Options & Metadata ===

/**
 * Gets all available filter options for UI components.
 * Returns cached results for performance.
 * 
 * @returns {Object} Filter options {location: [...], division: [...], etc.}
 */
function getAllFilterOptionsForModal() {
  // Ensure data is loaded and cached
  getAllEmployees();
  return _cachedFilterOptions;
}

/**
 * Gets unique locations from employee data.
 * @returns {Array<string>} Sorted array of unique locations
 */
function listUniqueLocations() {
  return getAllFilterOptionsForModal().location;
}

/**
 * Gets unique divisions from employee data.
 * @returns {Array<string>} Sorted array of unique divisions
 */
function listUniqueDivisions() {
  return getAllFilterOptionsForModal().division;
}

/**
 * Gets unique job titles from employee data.
 * @returns {Array<string>} Sorted array of unique job titles
 */
function listUniqueJobTitles() {
  return getAllFilterOptionsForModal().jobTitle;
}

/**
 * Gets unique pay types from employee data.
 * @returns {Array<string>} Sorted array of unique pay types
 */
function listUniquePayTypes() {
  return getAllFilterOptionsForModal().payType;
}

/**
 * Gets unique statuses from employee data.
 * @returns {Array<string>} Sorted array of unique statuses
 */
function listUniqueStatuses() {
  return getAllFilterOptionsForModal().status;
}

// === Convenience Aliases ===

/**
 * Alias for listUniqueLocations() - shorter function name for common use.
 * @returns {Array<string>} Unique locations
 */
function location() {
  return listUniqueLocations();
}

/**
 * Alias for listUniqueDivisions() - shorter function name for common use.
 * @returns {Array<string>} Unique divisions
 */
function division() {
  return listUniqueDivisions();
}

// === Testing & Debugging ===

/**
 * Comprehensive test function for debugging the employee directory module.
 * Tests data loading, caching, filtering, and performance.
 */
function testEmployeeDirectory() {
  console.log("=== Employee Directory Module Test ===");
  
  // Test basic loading
  console.log("1. Testing data loading...");
  const employees = getAllEmployees();
  console.log(`✓ Loaded ${employees.length} employees`);
  
  // Test caching
  console.log("2. Testing cache performance...");
  const start = new Date().getTime();
  getAllEmployees(); // Should be instant from cache
  const cacheTime = new Date().getTime() - start;
  console.log(`✓ Cache access took ${cacheTime}ms`);
  
  // Test filter options
  console.log("3. Testing filter options...");
  const filterOptions = getAllFilterOptionsForModal();
  for (const key in filterOptions) {
    console.log(`   ${key}: ${filterOptions[key].length} options`);
  }
  
  // Test filtering performance
  console.log("4. Testing filter performance...");
  if (filterOptions.location.length > 0) {
    const testFilter = { location: filterOptions.location[0] };
    const results = getEmployeesMatchingFilters(testFilter);
    console.log(`✓ Filter test returned ${results.length} results`);
  }
  
  // Test specific lookups
  console.log("5. Testing specific functions...");
  const activeEmployees = getActivePayrollEmployees();
  console.log(`✓ Active payroll employees: ${activeEmployees.length}`);
  
  const salaryEmployees = getEmployeesByPayType("Salary");
  console.log(`✓ Salary employees: ${salaryEmployees.length}`);
  
  console.log("=== Employee Directory Test Complete ===");
}

/**
 * Performance benchmark for the employee directory module.
 * Useful for optimizing large datasets.
 */
function benchmarkEmployeeDirectory() {
  console.log("=== Employee Directory Performance Benchmark ===");
  
  // Clear cache to test cold start
  resetEmployeeCache();
  
  // Benchmark initial load
  const coldStart = new Date().getTime();
  getAllEmployees();
  const coldTime = new Date().getTime() - coldStart;
  console.log(`Cold start (load + index): ${coldTime}ms`);
  
  // Benchmark cache access
  const warmStart = new Date().getTime();
  getAllEmployees();
  const warmTime = new Date().getTime() - warmStart;
  console.log(`Warm cache access: ${warmTime}ms`);
  
  // Benchmark filtering
  const filterOptions = getAllFilterOptionsForModal();
  if (filterOptions.location.length > 0) {
    const filterStart = new Date().getTime();
    getEmployeesMatchingFilters({ location: filterOptions.location[0] });
    const filterTime = new Date().getTime() - filterStart;
    console.log(`Single filter operation: ${filterTime}ms`);
  }
  
  console.log("=== Benchmark Complete ===");
}

/**
 * Gets all employees who should be paid in a specific payroll period.
 * Accounts for employees who were active during any part of the period.
 * 
 * @param {Object} period - Pay period object from calendar
 * @returns {Array<Object>} Employees eligible for this pay period
 */
function getEmployeesForPayPeriod(period) {
  if (!period || !period.periodStart || !period.periodEnd) {
    throw new Error("Invalid pay period provided");
  }
  
  const payType = period.staffType === "Salaried" ? "Salary" : "Hourly";
  
  return getAllEmployees().filter(emp => {
    // Must match pay type
    if (emp.PayType !== payType) return false;
    
    // Must be payroll employee
    if (emp.IsPayrollEmployee?.toLowerCase() !== "yes") return false;
    
    // Must be current status
    if (emp.Status?.toLowerCase() !== "current") return false;
    
    // Check if employed during any part of the period
    const startDate = parseDMY(emp.StartDate);
    const endDate = parseDMY(emp.TerminationDate);
    
    if (!startDate) return false;
    
    // Employee started before or during period AND
    // (no end date OR ended after period started)
    return startDate <= period.periodEnd && 
           (!endDate || endDate >= period.periodStart);
  });
}

/**
 * Gets payroll context for a specific employee.
 * Includes current period, previous period, and next period.
 * 
 * @param {string|number} employeeNumber - Employee number
 * @returns {Object|null} Pay period context for employee
 */
function getEmployeePayPeriodContext(employeeNumber) {
  const employee = getEmployeeByNumber(employeeNumber);
  if (!employee) return null;
  
  // Determine staff type from pay type
  let staffType;
  if (employee.PayType === "Salary") {
    staffType = "Salaried";
  } else if (employee.PayType === "Hourly") {
    staffType = "Hourly";
  } else {
    return null; // Unknown pay type
  }
  
  const current = getCurrentPayPeriod(staffType);
  const previous = current ? getPreviousPayPeriod(current) : null;
  const next = current ? getNextPayPeriod(current) : null;
  
  return {
    employee: employee,
    staffType: staffType,
    current: current,
    previous: previous,
    next: next,
    isEligibleForCurrentPeriod: current ? 
      getEmployeesForPayPeriod(current).some(emp => emp.EmployeeNumber == employeeNumber) : false
  };
}

/**
 * Gets employees who were active during a specific historical period.
 * More precise than getEmployeesForPayPeriod for historical analysis.
 * 
 * @param {Object} period - Historical pay period
 * @param {string} [payType] - Optional pay type filter
 * @returns {Array<Object>} Employees who were active during the period
 */
function getEmployeesActiveInPeriod(period, payType = null) {
  if (!period || !period.periodStart || !period.periodEnd) {
    throw new Error("Invalid pay period provided");
  }
  
  const targetPayType = payType || (period.staffType === "Salaried" ? "Salary" : "Hourly");
  
  return getAllEmployees().filter(emp => {
    // Pay type filter
    if (emp.PayType !== targetPayType) return false;
    
    // Must be payroll employee
    if (emp.IsPayrollEmployee?.toLowerCase() !== "yes") return false;
    
    // Check employment dates
    const startDate = parseDMY(emp.StartDate);
    const endDate = parseDMY(emp.TerminationDate);
    
    if (!startDate) return false;
    
    // Was employed during the entire period OR any part of it
    return startDate <= period.periodEnd && 
           (!endDate || endDate >= period.periodStart);
  });
}
