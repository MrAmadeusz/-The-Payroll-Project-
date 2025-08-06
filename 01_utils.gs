/**
 * Loads unique standardized locations from the mapping file.
 * Shared across rate checkers and NMW checker.
 *
 * @returns {string[]} Sorted list of known locations
 */
function getKnownLocations() {
  const file = DriveApp.getFileById(CONFIG.LOCATION_MAPPING_FILE_ID);
  const csv = file.getBlob().getDataAsString();
  const rows = Utilities.parseCsv(csv);

  const standardColIndex = rows[0].indexOf('Standard');
  if (standardColIndex === -1) {
    throw new Error("Missing 'Standard' column in location mapping file.");
  }

  const uniqueStandards = new Set();
  for (let i = 1; i < rows.length; i++) {
    const value = rows[i][standardColIndex].trim();
    if (value) uniqueStandards.add(value);
  }

  return Array.from(uniqueStandards).sort();
}

/**
 * Finds a CSV file by keyword in a Drive folder.
 */
function getCsvFileByKeyword(folder, keyword) {
  const normalise = str => str.toLowerCase().replace(/\s+/g, ' ').trim();
  const target = normalise(keyword);

  const files = folder.getFilesByType(MimeType.CSV);
  while (files.hasNext()) {
    const file = files.next();
    const name = normalise(file.getName());
    if (name.includes(target)) {
      return file;
    }
  }
  throw new Error(`CSV file with keyword '${keyword}' not found in folder.`);
}


/**
 * Finds a Google Sheet file by keyword in a Drive folder.
 */
function getGoogleSheetFileByKeyword(folder, keyword) {
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().toLowerCase().includes(keyword.toLowerCase())) {
      return file;
    }
  }
  throw new Error(`Google Sheet with keyword '${keyword}' not found in folder.`);
}

/**
 * Parses a CSV blob into an array of objects.
 */
function parseCsvToObjects(file) {
  const csv = Utilities.parseCsv(file.getBlob().getDataAsString());
  const [headers, ...rows] = csv;
  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

/**
 * Parses a string into a number or null if empty or 0.
 */
function parseAsNullableNumber(value) {
  const num = parseFloat(value);
  return isNaN(num) || num === 0 ? null : num;
}

/**
 * Builds a map of Location → LocationCode from CSV file.
 */
function parseLocationMap(file) {
  const csv = Utilities.parseCsv(file.getBlob().getDataAsString());
  const [headers, ...rows] = csv;

  const locationIndex = headers.indexOf("Location");
  const codeIndex = headers.indexOf("LocationCode");

  const map = {};
  rows.forEach(r => {
    if (r[locationIndex]) {
      map[r[locationIndex]] = r[codeIndex];
    }
  });
  return map;
}

function getFileByKeyword(folder, keyword) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().toLowerCase().includes(keyword.toLowerCase())) {
      return file;
    }
  }
  throw new Error(`No file matching '${keyword}' found in folder '${folder.getName()}'`);
}

/**
 * Loads and parses JSON from a Google Drive file.
 * Generic utility for any JSON-based data source.
 * 
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} Parsed JSON object
 */
function loadJsonFromDrive(fileId) {
  const file = DriveApp.getFileById(fileId);
  const json = file.getBlob().getDataAsString();
  return JSON.parse(json);
}

/**
 * Finds records within a date range.
 * Generic utility for any date-based filtering.
 * 
 * @param {Array} records - Array of objects with date fields
 * @param {Date} targetDate - Date to search for
 * @param {string} startField - Name of start date field
 * @param {string} endField - Name of end date field
 * @returns {Array} Matching records
 */
function findRecordsByDateRange(records, targetDate, startField, endField) {
  return records.filter(record => {
    const start = new Date(record[startField]);
    const end = record[endField] ? new Date(record[endField]) : null;
    return start && end && targetDate >= start && targetDate <= end;
  });
}


/**
 * Safely parses DD/MM/YYYY date strings into Date objects.
 * Generic utility for UK date format parsing.
 * 
 * @param {string} dateStr - Date string in DD/MM/YYYY format
 * @returns {Date|null} Parsed date or null if invalid
 */
function parseDMY(dateStr) {
  if (!dateStr || typeof dateStr !== "string") return null;
  const parts = dateStr.trim().split("/");
  if (parts.length !== 3) return null;
  const [day, month, year] = parts;
  const parsed = new Date(`${year}-${month}-${day}`);
  return isNaN(parsed) ? null : parsed;
}

/**
 * Determines if a parsed start date hits a milestone in the given month.
 * Returns the milestone (in years) if matched, or null if not.
 *
 * @param {Date|null} startDate
 * @param {number[]} milestones
 * @param {number} targetMonth (1–12)
 * @param {number} currentYear
 * @returns {number|null}
 */
function getMilestoneYear(startDate, milestones, targetMonth, currentYear) {
  if (!startDate || !(startDate instanceof Date) || isNaN(startDate)) return null;
  const years = currentYear - startDate.getFullYear();
  const anniversaryMonth = startDate.getMonth() + 1;
  return milestones.includes(years) && anniversaryMonth === targetMonth ? years : null;
}


/**
 * Builds an index from an array of objects for faster lookups.
 * Generic utility for creating field-based indexes.
 * 
 * @param {Array<Object>} data - Array of objects to index
 * @param {string} keyField - Field name to index by
 * @returns {Object} Index object where keys are field values, values are arrays of matching objects
 */
function buildDataIndex(data, keyField) {
  const index = {};
  data.forEach(item => {
    const key = (item[keyField] || '').trim();
    if (key) {
      if (!index[key]) {
        index[key] = [];
      }
      index[key].push(item);
    }
  });
  return index;
}

/**
 * Extracts unique values from an array of objects for a specific field.
 * Generic utility for building filter dropdowns.
 * 
 * @param {Array<Object>} data - Array of objects
 * @param {string} fieldName - Field to extract unique values from
 * @param {boolean} sorted - Whether to sort the results (default: true)
 * @returns {Array<string>} Unique values, optionally sorted
 */
function extractUniqueValues(data, fieldName, sorted = true) {
  const uniqueSet = new Set();
  data.forEach(item => {
    const value = item[fieldName];
    if (value && typeof value === 'string' && value.trim()) {
      uniqueSet.add(value.trim());
    }
  });
  const result = Array.from(uniqueSet);
  return sorted ? result.sort() : result;
}

/**
 * Filters an array of objects using multiple criteria with indexed optimization.
 * Generic utility for complex filtering operations.
 * 
 * @param {Array<Object>} data - Data to filter
 * @param {Object} filters - Filter criteria {fieldName: value, ...}
 * @param {Object} fieldMap - Maps filter keys to actual field names
 * @param {Object} indexes - Pre-built indexes for optimization {fieldName: index, ...}
 * @returns {Array<Object>} Filtered results
 */
function filterDataWithIndexes(data, filters, fieldMap = {}, indexes = {}) {
  if (!filters || Object.keys(filters).length === 0) return data;
  
  // Find the best index to start with (most restrictive)
  let candidates = data;
  let usedFilter = null;
  
  for (const filterKey in filters) {
    const fieldName = fieldMap[filterKey] || filterKey;
    if (indexes[fieldName] && filters[filterKey]) {
      candidates = indexes[fieldName][filters[filterKey]] || [];
      usedFilter = filterKey;
      break;
    }
  }
  
  // Apply remaining filters
  return candidates.filter(item => {
    for (const filterKey in filters) {
      if (filterKey === usedFilter) continue; // Skip already applied filter
      
      const fieldName = fieldMap[filterKey] || filterKey;
      const filterValue = filters[filterKey];
      const itemValue = (item[fieldName] || '').toLowerCase().trim();
      
      if (itemValue !== filterValue.toLowerCase().trim()) {
        return false;
      }
    }
    return true;
  });
}

/**
 * Performance timing utility for debugging slow operations.
 * 
 * @param {string} label - Description of the operation
 * @param {Function} operation - Function to time
 * @returns {*} Result of the operation
 */
function timeOperation(label, operation) {
  const start = new Date().getTime();
  const result = operation();
  const elapsed = new Date().getTime() - start;
  console.log(`${label} completed in ${elapsed}ms`);
  return result;
}

/**
 * Normalizes pay type input to standard values.
 * Generic utility for pay type canonicalization.
 * 
 * @param {string} payTypeInput - Input pay type (various formats)
 * @returns {string} Normalized pay type
 */
function normalizePayType(payTypeInput) {
  const canonical = { 
    salaried: "Salary", 
    salary: "Salary", 
    hourly: "Hourly" 
  };
  return canonical[payTypeInput.toLowerCase()] || payTypeInput;
}
/**
 * Finds the next or previous record in a sorted array.
 * Generic utility for navigating ordered datasets.
 * 
 * @param {Array} records - Sorted array of records
 * @param {Object} currentRecord - Current record to find adjacent to
 * @param {string} keyField - Field to match records by
 * @param {number} direction - 1 for next, -1 for previous
 * @returns {Object|null} Adjacent record or null
 */
function getAdjacentRecord(records, currentRecord, keyField, direction = 1) {
  const index = records.findIndex(r => r[keyField] === currentRecord[keyField]);
  if (index === -1) return null;
  
  const targetIndex = index + direction;
  return (targetIndex >= 0 && targetIndex < records.length) ? records[targetIndex] : null;
}

/**
 * Creates a cached function wrapper that stores results for repeated calls.
 * Generic utility for any expensive operation that benefits from caching.
 * 
 * @param {Function} fn - Function to cache
 * @param {Function} keyGenerator - Optional function to generate cache keys
 * @param {number} ttlMs - Time to live in milliseconds (default: 5 minutes)
 * @returns {Function} Cached version of the function
 */
function createCachedFunction(fn, keyGenerator = null, ttlMs = 300000) {
  const cache = new Map();
  
  return function(...args) {
    const key = keyGenerator ? keyGenerator(...args) : JSON.stringify(args);
    const now = Date.now();
    
    // Check if cached result exists and is still valid
    if (cache.has(key)) {
      const { result, timestamp } = cache.get(key);
      if (now - timestamp < ttlMs) {
        return result;
      }
      cache.delete(key);
    }
    
    // Execute function and cache result
    const result = fn.apply(this, args);
    cache.set(key, { result, timestamp: now });
    
    return result;
  };
}

/**
 * Validates data against a schema with detailed error reporting.
 * Generic utility for any data validation needs.
 * 
 * @param {Object} data - Data to validate
 * @param {Object} schema - Validation schema
 * @returns {Object} Validation result with detailed errors
 */
function validateDataSchema(data, schema) {
  const errors = [];
  const warnings = [];
  
  for (const field in schema) {
    const rules = schema[field];
    const value = data[field];
    
    // Required field check
    if (rules.required && (value === undefined || value === null || value === '')) {
      errors.push(`${field} is required`);
      continue;
    }
    
    // Skip other checks if field is not required and empty
    if (!rules.required && (value === undefined || value === null || value === '')) {
      continue;
    }
    
    // Type check
    if (rules.type && typeof value !== rules.type) {
      errors.push(`${field} must be of type ${rules.type}`);
    }
    
    // Range check for numbers
    if (rules.min !== undefined && value < rules.min) {
      errors.push(`${field} must be at least ${rules.min}`);
    }
    
    if (rules.max !== undefined && value > rules.max) {
      errors.push(`${field} must be at most ${rules.max}`);
    }
    
    // Custom validation function
    if (rules.validator && typeof rules.validator === 'function') {
      const customResult = rules.validator(value, data);
      if (customResult !== true) {
        errors.push(customResult || `${field} failed custom validation`);
      }
    }
    
    // Enum check
    if (rules.enum && !rules.enum.includes(value)) {
      errors.push(`${field} must be one of: ${rules.enum.join(', ')}`);
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings
  };
}

/**
 * Creates a batch processor for handling large datasets efficiently.
 * Generic utility for any batch processing needs with progress tracking.
 * 
 * @param {Array} items - Items to process
 * @param {Function} processor - Function to process each item
 * @param {Object} options - Processing options
 * @returns {Object} Batch processing results
 */
function processBatch(items, processor, options = {}) {
  const {
    batchSize = 100,
    onProgress = null,
    onError = 'continue', // 'continue' | 'stop'
    maxErrors = 10
  } = options;
  
  const results = [];
  const errors = [];
  const startTime = Date.now();
  
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    
    for (const item of batch) {
      try {
        const result = processor(item, i + batch.indexOf(item));
        results.push(result);
      } catch (error) {
        errors.push({
          item: item,
          index: i + batch.indexOf(item),
          error: error.message
        });
        
        if (onError === 'stop' || errors.length >= maxErrors) {
          break;
        }
      }
    }
    
    // Progress callback
    if (onProgress) {
      const progress = Math.min(i + batchSize, items.length);
      onProgress(progress, items.length, results.length, errors.length);
    }
    
    // Break if too many errors
    if (errors.length >= maxErrors) {
      break;
    }
  }
  
  return {
    results: results,
    errors: errors,
    totalProcessed: results.length + errors.length,
    processingTime: Date.now() - startTime,
    success: errors.length < maxErrors
  };
}

/**
 * Groups array of objects by multiple fields simultaneously.
 * Enhanced version of the basic groupBy for complex grouping scenarios.
 * 
 * @param {Array<Object>} data - Data to group
 * @param {Array<string>} groupFields - Fields to group by
 * @returns {Object} Nested grouping structure
 */
function groupByMultipleFields(data, groupFields) {
  if (groupFields.length === 0) return { all: data };
  
  const [firstField, ...remainingFields] = groupFields;
  const groups = {};
  
  data.forEach(item => {
    const key = item[firstField] || 'Unknown';
    if (!groups[key]) groups[key] = [];
    groups[key].push(item);
  });
  
  // Recursively group by remaining fields
  if (remainingFields.length > 0) {
    Object.keys(groups).forEach(key => {
      groups[key] = groupByMultipleFields(groups[key], remainingFields);
    });
  }
  
  return groups;
}

/**
 * Calculates comprehensive statistics for numerical arrays.
 * Generic utility for any statistical analysis needs.
 * 
 * @param {Array<number>} values - Numerical values to analyze
 * @returns {Object} Statistical summary
 */
function calculateStatistics(values) {
  if (!Array.isArray(values) || values.length === 0) {
    return { count: 0, sum: 0, mean: 0, median: 0, min: 0, max: 0, stdDev: 0 };
  }
  
  const numericValues = values.filter(v => typeof v === 'number' && !isNaN(v));
  const count = numericValues.length;
  
  if (count === 0) {
    return { count: 0, sum: 0, mean: 0, median: 0, min: 0, max: 0, stdDev: 0 };
  }
  
  const sum = numericValues.reduce((acc, val) => acc + val, 0);
  const mean = sum / count;
  
  const sorted = [...numericValues].sort((a, b) => a - b);
  const median = count % 2 === 0 
    ? (sorted[count / 2 - 1] + sorted[count / 2]) / 2
    : sorted[Math.floor(count / 2)];
  
  const variance = numericValues.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / count;
  const stdDev = Math.sqrt(variance);
  
  return {
    count: count,
    sum: parseFloat(sum.toFixed(2)),
    mean: parseFloat(mean.toFixed(2)),
    median: parseFloat(median.toFixed(2)),
    min: Math.min(...numericValues),
    max: Math.max(...numericValues),
    stdDev: parseFloat(stdDev.toFixed(2)),
    variance: parseFloat(variance.toFixed(2))
  };
}

/**
 * Creates a progress tracking system for long-running operations.
 * Generic utility for any operation that needs progress monitoring.
 * 
 * @param {string} operationId - Unique identifier for the operation
 * @returns {Object} Progress tracker with update and check methods
 */
function createProgressTracker(operationId) {
  const progressStore = PropertiesService.getScriptProperties();
  
  return {
    update: function(status, message, progress = null) {
      const progressData = {
        status: status,
        message: message,
        progress: progress,
        updatedAt: new Date().toISOString()
      };
      progressStore.setProperty(operationId, JSON.stringify(progressData));
    },
    
    get: function() {
      const data = progressStore.getProperty(operationId);
      return data ? JSON.parse(data) : null;
    },
    
    complete: function(finalMessage = 'Completed') {
      this.update('completed', finalMessage, 100);
    },
    
    fail: function(errorMessage) {
      this.update('failed', errorMessage, null);
    },
    
    cleanup: function() {
      progressStore.deleteProperty(operationId);
    }
  };
}

/**
 * Safely executes a function with error handling and retries.
 * Generic utility for any operation that might fail and need retries.
 * 
 * @param {Function} fn - Function to execute
 * @param {Object} options - Execution options
 * @returns {Object} Execution result
 */
function safeExecute(fn, options = {}) {
  const {
    maxRetries = 3,
    retryDelay = 1000,
    onRetry = null,
    fallback = null
  } = options;
  
  let lastError;
  
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const result = fn();
      return {
        success: true,
        result: result,
        attempts: attempt + 1
      };
    } catch (error) {
      lastError = error;
      
      if (attempt < maxRetries) {
        if (onRetry) onRetry(attempt + 1, error);
        Utilities.sleep(retryDelay);
      }
    }
  }
  
  // All retries failed
  if (fallback && typeof fallback === 'function') {
    try {
      const fallbackResult = fallback(lastError);
      return {
        success: true,
        result: fallbackResult,
        attempts: maxRetries + 1,
        usedFallback: true
      };
    } catch (fallbackError) {
      lastError = fallbackError;
    }
  }
  
  return {
    success: false,
    error: lastError.message,
    attempts: maxRetries + 1
  };
}

/**
 * Creates a configurable data export system.
 * Generic utility for exporting data in multiple formats.
 * 
 * @param {Array<Object>} data - Data to export
 * @param {Object} options - Export configuration
 * @returns {Object} Export result
 */
function exportDataToFormat(data, options = {}) {
  const {
    format = 'csv',
    filename = `export_${new Date().getTime()}`,
    folderId = null,
    columns = null,
    headers = true
  } = options;
  
  try {
    if (format === 'csv') {
      return _exportToCSV(data, { filename, folderId, columns, headers });
    } else if (format === 'json') {
      return _exportToJSON(data, { filename, folderId });
    } else if (format === 'sheet') {
      return _exportToSheet(data, { filename, columns, headers });
    } else {
      throw new Error(`Unsupported export format: ${format}`);
    }
  } catch (error) {
    return {
      success: false,
      error: error.message,
      format: format
    };
  }
}

/**
 * Private helper for CSV export.
 * @private
 */
function _exportToCSV(data, options) {
  const { filename, folderId, columns, headers } = options;
  
  if (!data || data.length === 0) {
    throw new Error("No data to export");
  }
  
  // Determine columns
  const exportColumns = columns || Object.keys(data[0]);
  
  // Build CSV content
  let csvContent = '';
  
  // Add headers if requested
  if (headers) {
    csvContent += exportColumns.join(',') + '\n';
  }
  
  // Add data rows
  data.forEach(row => {
    const values = exportColumns.map(col => {
      const value = row[col] || '';
      // Escape commas and quotes
      return typeof value === 'string' && (value.includes(',') || value.includes('"')) 
        ? `"${value.replace(/"/g, '""')}"` 
        : value;
    });
    csvContent += values.join(',') + '\n';
  });
  
  // Save to Drive
  const blob = Utilities.newBlob(csvContent, 'text/csv', `${filename}.csv`);
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const file = folder.createFile(blob);
  
  return {
    success: true,
    format: 'csv',
    fileId: file.getId(),
    fileName: file.getName(),
    url: file.getUrl(),
    rowsExported: data.length
  };
}

/**
 * Private helper for JSON export.
 * @private
 */
function _exportToJSON(data, options) {
  const { filename, folderId } = options;
  
  const jsonContent = JSON.stringify(data, null, 2);
  const blob = Utilities.newBlob(jsonContent, 'application/json', `${filename}.json`);
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const file = folder.createFile(blob);
  
  return {
    success: true,
    format: 'json',
    fileId: file.getId(),
    fileName: file.getName(),
    url: file.getUrl(),
    recordsExported: data.length
  };
}

/**
 * Private helper for Sheet export.
 * @private
 */
function _exportToSheet(data, options) {
  const { filename, columns, headers } = options;
  
  if (!data || data.length === 0) {
    throw new Error("No data to export");
  }
  
  const exportColumns = columns || Object.keys(data[0]);
  const ss = SpreadsheetApp.create(filename);
  const sheet = ss.getActiveSheet();
  
  let row = 1;
  
  // Add headers if requested
  if (headers) {
    sheet.getRange(row, 1, 1, exportColumns.length).setValues([exportColumns]);
    row++;
  }
  
  // Add data
  const dataRows = data.map(item => exportColumns.map(col => item[col] || ''));
  if (dataRows.length > 0) {
    sheet.getRange(row, 1, dataRows.length, exportColumns.length).setValues(dataRows);
  }
  
  return {
    success: true,
    format: 'sheet',
    fileId: ss.getId(),
    fileName: ss.getName(),
    url: ss.getUrl(),
    rowsExported: data.length
  };
}

/**
 * Simple date range picker - just start and end dates
 */
function promptForDateRange(title = "Select Date Range for Report") {
  const ui = SpreadsheetApp.getUi();
  
  // Get start date
  const startResponse = ui.prompt(
    title,
    'Enter START date (DD/MM/YYYY):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startResponse.getSelectedButton() !== ui.Button.OK) {
    throw new Error('Report cancelled by user.');
  }
  
  const startDate = parseDMY(startResponse.getResponseText().trim());
  if (!startDate) {
    throw new Error('Invalid start date format. Please use DD/MM/YYYY.');
  }
  
  // Get end date
  const endResponse = ui.prompt(
    title,
    'Enter END date (DD/MM/YYYY):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endResponse.getSelectedButton() !== ui.Button.OK) {
    throw new Error('Report cancelled by user.');
  }
  
  const endDate = parseDMY(endResponse.getResponseText().trim());
  if (!endDate) {
    throw new Error('Invalid end date format. Please use DD/MM/YYYY.');
  }
  
  // Basic validation
  if (startDate > endDate) {
    throw new Error('Start date must be before end date.');
  }
  
  return {
    startDate: startDate,
    endDate: endDate
  };
}

function isDateInRange(dateToCheck, dateRange) {
  let checkDate;
  
  if (typeof dateToCheck === 'string') {
    checkDate = parseDMY(dateToCheck);
  } else if (dateToCheck instanceof Date) {
    checkDate = dateToCheck;
  } else {
    return false;
  }
  
  if (!checkDate) return false;
  
  return checkDate >= dateRange.startDate && checkDate <= dateRange.endDate;
}

function formatDateRangeForDisplay(dateRange) {
  const formatDate = (date) => {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  };
  
  const start = formatDate(dateRange.startDate);
  const end = formatDate(dateRange.endDate);
  
  return `${start} - ${end}`;
}

function saveWeeklyEmployeeSnapshot() {
  const sourceFolder = DriveApp.getFolderById(CONFIG.EMPLOYEE_MASTER_FOLDER_ID);
  const snapshotFolder = DriveApp.getFolderById(CONFIG.EMPLOYEE_SNAPSHOTS_FOLDER_ID);
  const sourceFile = getCsvFileByKeyword(sourceFolder, 'All');

  if (!sourceFile) throw new Error("❌ No 'All Employees' CSV file found to snapshot.");

  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const snapshotName = `EmployeeSnapshot_${date}.csv`;

  sourceFile.makeCopy(snapshotName, snapshotFolder);
}

function loadLatestEmployeeSnapshot() {
  const folder = DriveApp.getFolderById(CONFIG.EMPLOYEE_SNAPSHOTS_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.CSV);

  let latestFile = null;
  let latestDate = 0;

  while (files.hasNext()) {
    const file = files.next();
    const match = file.getName().match(/EmployeeSnapshot_(\d{4}-\d{2}-\d{2})\.csv/);
    if (match) {
      const fileDate = new Date(match[1]).getTime();
      if (fileDate > latestDate) {
        latestDate = fileDate;
        latestFile = file;
      }
    }
  }

  if (!latestFile) throw new Error('❌ No previous snapshot file found in snapshot folder.');

  return parseCsvToObjects(latestFile);
}



/**
 * Creates a lookup index for employees using normalized employee numbers.
 * This ensures consistent lookups regardless of input format.
 * 
 * @param {Array<Object>} employees - Array of employee objects  
 * @param {string} employeeNumberField - Field name containing employee numbers
 * @returns {Object} Lookup index with normalized employee numbers as keys
 */
function createNormalizedEmployeeIndex(employees, employeeNumberField = 'EmployeeNumber') {
  const index = {};
  const duplicates = [];
  
  employees.forEach(employee => {
    const original = employee[employeeNumberField];
    const normalized = normalizeEmployeeNumber(original);
    
    if (normalized) {
      if (index[normalized]) {
        // Duplicate found
        duplicates.push({
          employeeNumber: normalized,
          employee1: index[normalized],
          employee2: employee
        });
        console.warn(`Duplicate normalized employee number: ${normalized}`);
      } else {
        index[normalized] = employee;
      }
    }
  });
  
  console.log(`Created normalized employee index with ${Object.keys(index).length} entries`);
  if (duplicates.length > 0) {
    console.warn(`Found ${duplicates.length} duplicate employee numbers after normalization`);
  }
  
  return {
    index: index,
    duplicates: duplicates
  };
}

/**
 * Test function to validate employee number normalization
 */
function testEmployeeNumberNormalization() {
  console.log('=== Employee Number Normalization Test ===');
  
  const testCases = [
    "1234",
    "567", 
    "00012345",
    "12345678",
    123,
    "abc123",
    "",
    null,
    "  1234  ",
    "0000567",
    "123456789" // Too long
  ];
  
  testCases.forEach(testCase => {
    const result = normalizeEmployeeNumber(testCase);
    console.log(`Input: "${testCase}" → Output: "${result}"`);
  });
  
  // Test with sample data
  const sampleEmployees = [
    { EmployeeNumber: "1234", Name: "John Doe" },
    { EmployeeNumber: "00005678", Name: "Jane Smith" },
    { EmployeeNumber: "567", Name: "Bob Johnson" },
    { EmployeeNumber: "invalid", Name: "Invalid Employee" }
  ];
  
  console.log('Testing array normalization...');
  const summary = normalizeEmployeeNumbersInArray(sampleEmployees);
  console.log('Normalization summary:', summary);
  console.log('Updated employees:', sampleEmployees);
  
  console.log('=== Test Complete ===');
}

/**
 * PRODUCTION-READY SNAPSHOT FUNCTIONS
 * These bypass the Google Apps Script MIME type bug
 * Add these to your 01_utils.gs file
 */

/**
 * Creates Gross to Nett snapshot (bypasses API MIME type bug)
 * REPLACES: saveMonthlyGrossToNettSnapshot()
 */
function saveMonthlyGrossToNettSnapshot() {
  try {
    console.log('📸 Creating Gross to Nett snapshot...');
    
    const sourceFolder = DriveApp.getFolderById(CONFIG.PAYROLL_PREVIEW_FILE_ID);
    const snapshotFolder = DriveApp.getFolderById(CONFIG.GROSS_TO_NETT_FOLDER_ID);
    
    // Find target files by name pattern (ignore MIME type)
    const targetFiles = [];
    const fileIterator = sourceFolder.getFiles();
    
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      const fileName = file.getName();
      
      if (fileName.toLowerCase().includes('gross to nett') && !fileName.endsWith('.xlsx')) {
        targetFiles.push(file);
      }
    }
    
    if (targetFiles.length !== 2) {
      throw new Error(`Expected 2 "Gross to Nett" files, found ${targetFiles.length}`);
    }
    
    console.log(`✅ Found files: ${targetFiles.map(f => f.getName()).join(', ')}`);
    
    // Read data using SpreadsheetApp (bypasses MIME type bug)
    const allData = [];
    for (const file of targetFiles) {
      console.log(`📖 Reading: ${file.getName()}`);
      const data = readFileAsSpreadsheet(file);
      console.log(`📊 Extracted ${data.length} rows`);
      allData.push(...data);
    }
    
    console.log(`📊 Total combined data: ${allData.length} rows`);
    
    if (allData.length === 0) {
      throw new Error('No data found in the files');
    }
    
    // Create snapshot
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const snapshotName = `GrossToNett_Snapshot_${timestamp}`;
    
    const newSheet = SpreadsheetApp.create(snapshotName);
    writeObjectArrayToSheet(newSheet, allData);
    
    // Move to snapshot folder
    const file = DriveApp.getFileById(newSheet.getId());
    snapshotFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    
    console.log(`✅ Snapshot saved: ${snapshotName} (${allData.length} records)`);
    
    return {
      success: true,
      fileName: snapshotName,
      fileId: newSheet.getId(),
      recordCount: allData.length
    };
    
  } catch (error) {
    console.error(`❌ Gross to Nett snapshot failed: ${error.message}`);
    throw error;
  }
}

/**
 * Creates Payments snapshot (same approach as Gross to Nett)
 * REPLACES: saveMonthlyPaymentsSnapshot()
 */
function saveMonthlyPaymentsSnapshot() {
  try {
    console.log('📸 Creating Payments snapshot...');
    
    const sourceFolder = DriveApp.getFolderById(CONFIG.PAYROLL_PREVIEW_FILE_ID);
    const snapshotFolder = DriveApp.getFolderById(CONFIG.PAYMENTS_AND_DEDUCTIONS_FOLDER_ID);
    
    // Find target files by name pattern (ignore MIME type)
    const targetFiles = [];
    const fileIterator = sourceFolder.getFiles();
    
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      const fileName = file.getName();
      
      if (fileName.toLowerCase().includes('payments') && !fileName.endsWith('.xlsx')) {
        targetFiles.push(file);
      }
    }
    
    if (targetFiles.length !== 2) {
      throw new Error(`Expected 2 "Payments" files, found ${targetFiles.length}`);
    }
    
    console.log(`✅ Found files: ${targetFiles.map(f => f.getName()).join(', ')}`);
    
    // Read data using SpreadsheetApp (bypasses MIME type bug)
    const allData = [];
    for (const file of targetFiles) {
      console.log(`📖 Reading: ${file.getName()}`);
      const data = readFileAsSpreadsheet(file);
      console.log(`📊 Extracted ${data.length} rows`);
      allData.push(...data);
    }
    
    console.log(`📊 Total combined data: ${allData.length} rows`);
    
    if (allData.length === 0) {
      throw new Error('No data found in the files');
    }
    
    // Create snapshot
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const snapshotName = `Payments_Snapshot_${timestamp}`;
    
    const newSheet = SpreadsheetApp.create(snapshotName);
    writeObjectArrayToSheet(newSheet, allData);
    
    // Move to snapshot folder
    const file = DriveApp.getFileById(newSheet.getId());
    snapshotFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    
    console.log(`✅ Snapshot saved: ${snapshotName} (${allData.length} records)`);
    
    return {
      success: true,
      fileName: snapshotName,
      fileId: newSheet.getId(),
      recordCount: allData.length
    };
    
  } catch (error) {
    console.error(`❌ Payments snapshot failed: ${error.message}`);
    throw error;
  }
}

/**
 * Helper: Read file as spreadsheet (ignores MIME type)
 */
function readFileAsSpreadsheet(file) {
  try {
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const sheet = spreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) {
      return [];
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const [headers, ...rows] = data;
    
    return rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      return obj;
    });
    
  } catch (error) {
    throw new Error(`Could not read ${file.getName()} as spreadsheet: ${error.message}`);
  }
}

/**
 * Helper: Write object array to sheet
 */
function writeObjectArrayToSheet(spreadsheet, data) {
  if (data.length === 0) return;
  
  const sheet = spreadsheet.getActiveSheet();
  const headers = Object.keys(data[0]);
  
  // Prepare data array: [headers row, ...data rows]
  const sheetData = [
    headers,
    ...data.map(row => headers.map(header => row[header] || ''))
  ];
  
  // Write data
  sheet.getRange(1, 1, sheetData.length, headers.length).setValues(sheetData);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // Auto-resize and freeze
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

/**
 * Loads latest Gross to Nett snapshot (updated to work with new naming)
 * REPLACES: loadLatestGrossToNettSnapshot()
 */
function loadLatestGrossToNettSnapshot() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.GROSS_TO_NETT_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    let latestFile = null;
    let latestTimestamp = '';
    
    while (files.hasNext()) {
      const file = files.next();
      const match = file.getName().match(/GrossToNett_Snapshot_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2})/);
      if (match) {
        const timestamp = match[1];
        if (timestamp > latestTimestamp) {
          latestTimestamp = timestamp;
          latestFile = file;
        }
      }
    }
    
    if (!latestFile) {
      throw new Error('No Gross to Nett snapshot found. Run saveMonthlyGrossToNettSnapshot() first.');
    }
    
    console.log(`📂 Loading snapshot: ${latestFile.getName()}`);
    return readFileAsSpreadsheet(latestFile);
    
  } catch (error) {
    console.error(`❌ Failed to load Gross to Nett snapshot: ${error.message}`);
    throw error;
  }
}

/**
 * Test both snapshot functions
 */
function testBothSnapshots() {
  console.log('🧪 Testing both snapshot functions...');
  
  try {
    // Test Gross to Nett
    console.log('\n=== Testing Gross to Nett ===');
    const grossResult = saveMonthlyGrossToNettSnapshot();
    console.log(`✅ Gross to Nett: ${grossResult.recordCount} records`);
    
    // Test Payments
    console.log('\n=== Testing Payments ===');
    const paymentsResult = saveMonthlyPaymentsSnapshot();
    console.log(`✅ Payments: ${paymentsResult.recordCount} records`);
    
    // Test loading
    console.log('\n=== Testing Load ===');
    const loadedData = loadLatestGrossToNettSnapshot();
    console.log(`✅ Loaded: ${loadedData.length} records`);
    
    console.log('\n🎉 ALL TESTS PASSED!');
    
    return {
      success: true,
      grossToNett: grossResult,
      payments: paymentsResult,
      loadedRecords: loadedData.length
    };
    
  } catch (error) {
    console.error(`❌ Test failed: ${error.message}`);
    return { 
      success: false, 
      error: error.message 
    };
  }
}

/**
 * MISSING FUNCTION: normalizeEmployeeNumber
 * Add this to your 01_utils.gs file
 */

/**
 * Normalizes employee numbers to ensure consistent matching across different data sources.
 * Handles various formats like: "1234", "00001234", "12345678", etc.
 * 
 * @param {string|number} employeeNumber - Employee number in any format
 * @returns {string|null} Normalized employee number or null if invalid
 */
function normalizeEmployeeNumber(employeeNumber) {
  // Handle null, undefined, or empty values
  if (employeeNumber === null || employeeNumber === undefined || employeeNumber === '') {
    return null;
  }
  
  // Convert to string and trim whitespace
  let normalized = String(employeeNumber).trim();
  
  // Return null for obviously invalid values
  if (normalized === '' || normalized.toLowerCase() === 'null') {
    return null;
  }
  
  // Remove any non-numeric characters (in case there are letters/symbols)
  normalized = normalized.replace(/[^0-9]/g, '');
  
  // Return null if no numeric characters found
  if (normalized === '') {
    return null;
  }
  
  // Handle different scenarios based on length
  if (normalized.length <= 4) {
    // Short numbers: pad to 8 digits with leading zeros
    // "1234" → "00001234"
    return normalized.padStart(8, '0');
  } else if (normalized.length <= 8) {
    // Medium numbers: pad to 8 digits
    // "12345" → "00012345"
    return normalized.padStart(8, '0');
  } else if (normalized.length > 8) {
    // Long numbers: might be already padded or have extra digits
    // Take the last 8 digits if longer than 8
    return normalized.slice(-8);
  }
  
  return normalized;
}

/**
 * Normalizes employee numbers in an array of objects (modifies in place)
 * 
 * @param {Array<Object>} employees - Array of employee objects
 * @param {string} fieldName - Field name containing employee numbers (default: 'EmployeeNumber')
 * @returns {Object} Summary of normalization results
 */
function normalizeEmployeeNumbersInArray(employees, fieldName = 'EmployeeNumber') {
  let normalized = 0;
  let invalid = 0;
  let unchanged = 0;
  
  employees.forEach(employee => {
    const original = employee[fieldName];
    const normalizedValue = normalizeEmployeeNumber(original);
    
    if (normalizedValue === null) {
      invalid++;
    } else if (normalizedValue !== String(original).trim()) {
      employee[fieldName] = normalizedValue;
      normalized++;
    } else {
      unchanged++;
    }
  });
  
  return {
    total: employees.length,
    normalized: normalized,
    unchanged: unchanged,
    invalid: invalid
  };
}

/**
 * Test the normalization function with various inputs
 */
function testEmployeeNumberNormalization() {
  console.log('🧪 Testing Employee Number Normalization');
  console.log('========================================');
  
  const testCases = [
    "1234",           // Short number
    "567",            // Very short
    "00012345",       // Already padded
    "12345678",       // Full 8 digits
    "123456789",      // Too long
    123,              // Numeric input
    "abc123def",      // With letters
    "",               // Empty
    null,             // Null
    "  1234  ",       // With spaces
    "0000567",        // Leading zeros
    "30014743"        // Typical employee number
  ];
  
  console.log('INPUT → OUTPUT');
  console.log('--------------');
  
  testCases.forEach(testCase => {
    const result = normalizeEmployeeNumber(testCase);
    console.log(`"${testCase}" → "${result}"`);
  });
  
  // Test with array
  console.log('\n📋 Testing array normalization...');
  const sampleEmployees = [
    { EmployeeNumber: "1234", Name: "John Doe" },
    { EmployeeNumber: "00005678", Name: "Jane Smith" },
    { EmployeeNumber: "567", Name: "Bob Johnson" },
    { EmployeeNumber: "invalid", Name: "Invalid Employee" },
    { EmployeeNumber: "30014743", Name: "Real Employee" }
  ];
  
  console.log('\nBefore normalization:');
  sampleEmployees.forEach(emp => {
    console.log(`${emp.Name}: "${emp.EmployeeNumber}"`);
  });
  
  const summary = normalizeEmployeeNumbersInArray(sampleEmployees);
  
  console.log('\nAfter normalization:');
  sampleEmployees.forEach(emp => {
    console.log(`${emp.Name}: "${emp.EmployeeNumber}"`);
  });
  
  console.log('\nNormalization Summary:');
  console.log(`Total: ${summary.total}`);
  console.log(`Normalized: ${summary.normalized}`);
  console.log(`Unchanged: ${summary.unchanged}`);
  console.log(`Invalid: ${summary.invalid}`);
  
  return summary;
}

/**
 * Quick fix: Test AOE report again with the normalization function now available
 */
function testAOEReportWithNormalization() {
  console.log('💰 Testing AOE Report with normalization function...');
  
  try {
    // First test the normalization function
    console.log('\n🧪 Testing normalization function...');
    testEmployeeNumberNormalization();
    
    // Now test the AOE report
    console.log('\n📋 Testing AOE Report...');
    const aoeResult = testLeaversAOEReport();
    
    if (aoeResult.success) {
      console.log('\n🎉 SUCCESS! AOE Report now works with normalization');
      console.log(`📄 Report rows: ${aoeResult.totalRows}`);
      console.log(`💰 Leavers with outstanding payments: ${aoeResult.dataRows}`);
      console.log(`🔗 Report URL: ${aoeResult.url}`);
    } else {
      console.log(`❌ AOE Report still failed: ${aoeResult.error}`);
    }
    
    return aoeResult;
    
  } catch (error) {
    console.error(`❌ Test failed: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}
