/**
 * FILE COMPARISON TOOL - CORE ENGINE
 * Universal file difference detection with modular analysis system
 * Module: 13_fileCompare.gs
 */

/**
 * Main entry point - shows the file comparison dashboard
 * Add this to your main menu system
 */
function showFileComparisonTool() {
  const html = HtmlService.createHtmlOutputFromFile('13_fileCompareDashboard')
    .setTitle("File Comparison Tool")
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Core comparison function - called from dashboard
 * @param {Object} config - Comparison configuration
 * @returns {Object} Comparison results with differences
 */
function compareFiles(config) {
  const tracker = createProgressTracker(`file_compare_${Date.now()}`);
  
  try {
    tracker.update('starting', 'Initializing file comparison...', 0);
    
    // Validate inputs
    if (!config.fileA || !config.fileB) {
      throw new Error('Both files must be selected for comparison');
    }
    
    // Load files using universal adapter
    tracker.update('loading', 'Loading File A...', 10);
    const dataA = loadFileAsGrid(config.fileA);
    
    tracker.update('loading', 'Loading File B...', 30);
    const dataB = loadFileAsGrid(config.fileB);
    
    // Validate data structure
    if (!dataA.length || !dataB.length) {
      throw new Error('One or both files appear to be empty or unreadable');
    }
    
    // Perform core comparison
    tracker.update('comparing', 'Analyzing differences...', 50);
    const results = performCoreComparison(dataA, dataB, config);
    
    // Detect content type and suggest analysis modules
    tracker.update('analyzing', 'Detecting content type...', 80);
    const contentType = detectContentType(results.headersA, dataA);
    results.suggestedModules = getSuggestedAnalysisModules(contentType);
    
    // Add metadata
    results.metadata = {
      fileA: {
        name: config.fileA.name || 'File A',
        rows: dataA.length,
        columns: dataA[0] ? dataA[0].length : 0
      },
      fileB: {
        name: config.fileB.name || 'File B', 
        rows: dataB.length,
        columns: dataB[0] ? dataB[0].length : 0
      },
      comparedAt: new Date().toISOString(),
      processingTime: Date.now() - tracker.get().updatedAt
    };
    
    tracker.complete('Comparison completed successfully');
    
    console.log(`File comparison completed: ${results.summary.totalDifferences} differences found`);
    return results;
    
  } catch (error) {
    tracker.fail(error.message);
    console.error('File comparison failed:', error);
    throw error;
  }
}

/**
 * Universal file loader - handles CSV, GSheet, Excel
 * @param {Object} fileInfo - File information object
 * @returns {Array<Array>} 2D array representing the file data
 */
function loadFileAsGrid(fileInfo) {
  try {
    const file = DriveApp.getFileById(fileInfo.id);
    const fileName = file.getName().toLowerCase();
    
    // Determine file type and use appropriate loader
    if (fileName.endsWith('.csv')) {
      return loadCsvAsGrid(file);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      return loadExcelAsGrid(file);
    } else {
      // Assume Google Sheet or try as spreadsheet
      return loadSpreadsheetAsGrid(file);
    }
    
  } catch (error) {
    throw new Error(`Failed to load file: ${error.message}`);
  }
}

/**
 * Load CSV file as 2D grid
 * @param {File} file - Google Drive file
 * @returns {Array<Array>} 2D grid
 */
function loadCsvAsGrid(file) {
  const csvContent = file.getBlob().getDataAsString();
  const parsedData = Utilities.parseCsv(csvContent);
  
  if (!parsedData || parsedData.length === 0) {
    throw new Error('CSV file appears to be empty or invalid');
  }
  
  return parsedData;
}

/**
 * Load Excel file as 2D grid using SheetJS
 * @param {File} file - Google Drive file  
 * @returns {Array<Array>} 2D grid
 */
function loadExcelAsGrid(file) {
  // Note: This would require SheetJS integration
  // For now, provide fallback
  try {
    // Try to open as spreadsheet (if it's been converted)
    return loadSpreadsheetAsGrid(file);
  } catch (error) {
    throw new Error('Excel file support requires SheetJS integration. Please convert to Google Sheets format.');
  }
}

/**
 * Load Google Sheet as 2D grid
 * @param {File} file - Google Drive file
 * @returns {Array<Array>} 2D grid
 */
function loadSpreadsheetAsGrid(file) {
  const spreadsheet = SpreadsheetApp.openById(file.getId());
  const sheet = spreadsheet.getActiveSheet();
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    return [];
  }
  
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/**
 * Core comparison logic - handles the actual difference detection
 * @param {Array<Array>} dataA - File A as 2D grid
 * @param {Array<Array>} dataB - File B as 2D grid  
 * @param {Object} config - Comparison configuration
 * @returns {Object} Detailed comparison results
 */
function performCoreComparison(dataA, dataB, config) {
  const results = {
    differences: [],
    summary: {
      totalCellsA: 0,
      totalCellsB: 0,
      identicalCells: 0,
      modifiedCells: 0,
      addedRows: 0,
      deletedRows: 0,
      addedColumns: 0,
      deletedColumns: 0,
      totalDifferences: 0
    },
    headersA: dataA[0] || [],
    headersB: dataB[0] || [],
    comparisonMode: config.mode || 'positional'
  };
  
  // Handle different comparison modes
  if (config.mode === 'key-based' && config.keyColumn) {
    return performKeyBasedComparison(dataA, dataB, config, results);
  } else {
    return performPositionalComparison(dataA, dataB, config, results);
  }
}

/**
 * Positional comparison - row by row, cell by cell
 * @param {Array<Array>} dataA - File A data
 * @param {Array<Array>} dataB - File B data
 * @param {Object} config - Configuration
 * @param {Object} results - Results object to populate
 * @returns {Object} Comparison results
 */
function performPositionalComparison(dataA, dataB, config, results) {
  const maxRows = Math.max(dataA.length, dataB.length);
  const maxColsA = dataA[0] ? dataA[0].length : 0;
  const maxColsB = dataB[0] ? dataB[0].length : 0;
  const maxCols = Math.max(maxColsA, maxColsB);
  
  results.summary.totalCellsA = dataA.length * maxColsA;
  results.summary.totalCellsB = dataB.length * maxColsB;
  
  // Track structural differences
  if (dataA.length !== dataB.length) {
    if (dataA.length > dataB.length) {
      results.summary.deletedRows = dataA.length - dataB.length;
    } else {
      results.summary.addedRows = dataB.length - dataA.length;
    }
  }
  
  if (maxColsA !== maxColsB) {
    if (maxColsA > maxColsB) {
      results.summary.deletedColumns = maxColsA - maxColsB;
    } else {
      results.summary.addedColumns = maxColsB - maxColsA;
    }
  }
  
  // Compare cells
  for (let row = 0; row < maxRows; row++) {
    const rowA = dataA[row] || [];
    const rowB = dataB[row] || [];
    
    // Handle entire row additions/deletions
    if (!dataA[row] && dataB[row]) {
      results.differences.push({
        type: 'row_added',
        row: row,
        column: null,
        valueA: null,
        valueB: rowB,
        description: `Row ${row + 1} added`
      });
      results.summary.totalDifferences++;
      continue;
    }
    
    if (dataA[row] && !dataB[row]) {
      results.differences.push({
        type: 'row_deleted', 
        row: row,
        column: null,
        valueA: rowA,
        valueB: null,
        description: `Row ${row + 1} deleted`
      });
      results.summary.totalDifferences++;
      continue;
    }
    
    // Compare individual cells
    for (let col = 0; col < maxCols; col++) {
      const cellA = rowA[col] !== undefined ? rowA[col] : null;
      const cellB = rowB[col] !== undefined ? rowB[col] : null;
      
      if (!compareCellValues(cellA, cellB, config)) {
        const difference = {
          type: determineCellChangeType(cellA, cellB),
          row: row,
          column: col,
          valueA: cellA,
          valueB: cellB,
          description: `Cell [${row + 1}, ${col + 1}] changed`
        };
        
        // Add column name if headers available
        if (row > 0) { // Skip header row for column names
          const headerA = results.headersA[col];
          const headerB = results.headersB[col]; 
          if (headerA || headerB) {
            difference.columnName = headerA || headerB;
            difference.description = `${difference.columnName} changed in row ${row + 1}`;
          }
        }
        
        results.differences.push(difference);
        results.summary.modifiedCells++;
        results.summary.totalDifferences++;
      } else {
        results.summary.identicalCells++;
      }
    }
    
    // Progress update for large files
    if (row % 1000 === 0 && row > 0) {
      console.log(`Compared ${row} rows...`);
    }
  }
  
  return results;
}

/**
 * Key-based comparison - match rows by primary key column
 * @param {Array<Array>} dataA - File A data
 * @param {Array<Array>} dataB - File B data
 * @param {Object} config - Configuration with keyColumn
 * @param {Object} results - Results object to populate
 * @returns {Object} Comparison results
 */
function performKeyBasedComparison(dataA, dataB, config, results) {
  const keyColIndex = parseInt(config.keyColumn);
  
  if (keyColIndex < 0 || keyColIndex >= Math.max(dataA[0]?.length || 0, dataB[0]?.length || 0)) {
    throw new Error('Invalid key column specified');
  }
  
  // Build indexes for both files (skip header row)
  const indexA = buildDataIndex(dataA.slice(1).map((row, idx) => ({
    originalRow: idx + 1,
    data: row,
    key: row[keyColIndex]
  })), 'key');
  
  const indexB = buildDataIndex(dataB.slice(1).map((row, idx) => ({
    originalRow: idx + 1, 
    data: row,
    key: row[keyColIndex]
  })), 'key');
  
  const keysA = new Set(Object.keys(indexA));
  const keysB = new Set(Object.keys(indexB));
  const allKeys = new Set([...keysA, ...keysB]);
  
  // Compare by keys
  allKeys.forEach(key => {
    const recordsA = indexA[key] || [];
    const recordsB = indexB[key] || [];
    
    if (recordsA.length === 0) {
      // Key only in B - added
      recordsB.forEach(record => {
        results.differences.push({
          type: 'row_added',
          key: key,
          row: record.originalRow,
          column: null,
          valueA: null,
          valueB: record.data,
          description: `Record with key '${key}' added`
        });
      });
      results.summary.addedRows += recordsB.length;
    } else if (recordsB.length === 0) {
      // Key only in A - deleted
      recordsA.forEach(record => {
        results.differences.push({
          type: 'row_deleted',
          key: key,
          row: record.originalRow,
          column: null,
          valueA: record.data,
          valueB: null,
          description: `Record with key '${key}' deleted`
        });
      });
      results.summary.deletedRows += recordsA.length;
    } else {
      // Key in both - compare cells (assuming 1:1 match for simplicity)
      const recordA = recordsA[0];
      const recordB = recordsB[0];
      
      for (let col = 0; col < Math.max(recordA.data.length, recordB.data.length); col++) {
        const cellA = recordA.data[col];
        const cellB = recordB.data[col];
        
        if (!compareCellValues(cellA, cellB, config)) {
          results.differences.push({
            type: determineCellChangeType(cellA, cellB),
            key: key,
            row: recordA.originalRow,
            column: col,
            valueA: cellA,
            valueB: cellB,
            columnName: results.headersA[col] || results.headersB[col],
            description: `${results.headersA[col] || `Column ${col + 1}`} changed for key '${key}'`
          });
          results.summary.modifiedCells++;
        } else {
          results.summary.identicalCells++;
        }
      }
    }
  });
  
  results.summary.totalDifferences = results.differences.length;
  results.summary.totalCellsA = dataA.length * (dataA[0]?.length || 0);
  results.summary.totalCellsB = dataB.length * (dataB[0]?.length || 0);
  
  return results;
}

/**
 * Compare two cell values with configurable sensitivity
 * @param {*} valueA - Value from file A
 * @param {*} valueB - Value from file B
 * @param {Object} config - Comparison configuration
 * @returns {boolean} True if values are considered equal
 */
function compareCellValues(valueA, valueB, config) {
  // Handle null/undefined/empty cases
  if (valueA === null && valueB === null) return true;
  if (valueA === undefined && valueB === undefined) return true;
  if (valueA === '' && valueB === '') return true;
  
  // Convert to strings for comparison
  let strA = String(valueA || '');
  let strB = String(valueB || '');
  
  // Apply comparison options
  if (config.ignoreCase) {
    strA = strA.toLowerCase();
    strB = strB.toLowerCase();
  }
  
  if (config.ignoreWhitespace) {
    strA = strA.trim();
    strB = strB.trim();
  }
  
  return strA === strB;
}

/**
 * Determine the type of cell change
 * @param {*} valueA - Original value
 * @param {*} valueB - New value
 * @returns {string} Change type
 */
function determineCellChangeType(valueA, valueB) {
  if ((valueA === null || valueA === undefined || valueA === '') && 
      (valueB !== null && valueB !== undefined && valueB !== '')) {
    return 'cell_added';
  }
  
  if ((valueA !== null && valueA !== undefined && valueA !== '') && 
      (valueB === null || valueB === undefined || valueB === '')) {
    return 'cell_deleted';
  }
  
  return 'cell_modified';
}

/**
 * Detect content type from headers and sample data
 * @param {Array} headers - Column headers
 * @param {Array<Array>} sampleData - First few rows of data
 * @returns {string} Detected content type
 */
function detectContentType(headers, sampleData) {
  const headerStr = headers.join(' ').toLowerCase();
  
  // Payroll detection
  const payrollKeywords = ['employee', 'gross', 'net', 'salary', 'wage', 'hours', 'overtime', 'deduction'];
  if (payrollKeywords.some(keyword => headerStr.includes(keyword))) {
    return 'payroll';
  }
  
  // Sales detection
  const salesKeywords = ['product', 'price', 'quantity', 'revenue', 'sale', 'customer', 'order'];
  if (salesKeywords.some(keyword => headerStr.includes(keyword))) {
    return 'sales';
  }
  
  // HR detection
  const hrKeywords = ['hire', 'department', 'position', 'status', 'manager', 'location'];
  if (hrKeywords.some(keyword => headerStr.includes(keyword))) {
    return 'hr';
  }
  
  return 'generic';
}

/**
 * Get suggested analysis modules based on content type
 * @param {string} contentType - Detected content type
 * @returns {Array} Array of suggested module names
 */
function getSuggestedAnalysisModules(contentType) {
  const modules = {
    'payroll': ['payroll-intelligence'],
    'sales': ['sales-analysis'],
    'hr': ['hr-analysis'],
    'generic': []
  };
  
  return modules[contentType] || [];
}

/**
 * Get available files for comparison (from various sources)
 * @returns {Array} Array of file objects with id, name, type
 */
function getAvailableFiles() {
  const files = [];
  
  try {
    // Add files from common folders
    const foldersToCheck = [
      { id: CONFIG.EMPLOYEE_MASTER_FOLDER_ID, name: 'Employee Master' },
      { id: CONFIG.PAYROLL_DATA_FOLDER_ID, name: 'Payroll Data' },
      { id: CONFIG.GROSS_TO_NETT_FOLDER_ID, name: 'Gross to Nett' },
      { id: CONFIG.PAYMENTS_AND_DEDUCTIONS_FOLDER_ID, name: 'Payments & Deductions' }
    ];
    
    foldersToCheck.forEach(folderInfo => {
      try {
        const folder = DriveApp.getFolderById(folderInfo.id);
        const fileIterator = folder.getFiles();
        
        while (fileIterator.hasNext()) {
          const file = fileIterator.next();
          const fileName = file.getName().toLowerCase();
          
          // Only include supported formats
          if (fileName.endsWith('.csv') || 
              fileName.endsWith('.xlsx') || 
              fileName.endsWith('.xls') ||
              file.getMimeType() === MimeType.GOOGLE_SHEETS) {
            
            files.push({
              id: file.getId(),
              name: file.getName(),
              folder: folderInfo.name,
              type: getFileType(file),
              modified: file.getLastUpdated(),
              size: file.getSize()
            });
          }
        }
      } catch (folderError) {
        console.warn(`Could not access folder ${folderInfo.name}:`, folderError.message);
      }
    });
    
  } catch (error) {
    console.error('Error getting available files:', error);
  }
  
  // Sort by modification date, newest first
  files.sort((a, b) => b.modified - a.modified);
  
  return files;
}

/**
 * Determine file type from file object
 * @param {File} file - Google Drive file
 * @returns {string} File type
 */
function getFileType(file) {
  const mimeType = file.getMimeType();
  const name = file.getName().toLowerCase();
  
  if (name.endsWith('.csv')) return 'CSV';
  if (name.endsWith('.xlsx') || name.endsWith('.xls')) return 'Excel';
  if (mimeType === MimeType.GOOGLE_SHEETS) return 'Google Sheets';
  
  return 'Unknown';
}

/**
 * Export comparison results to various formats
 * @param {Object} results - Comparison results
 * @param {string} format - Export format ('csv', 'sheet', 'json')
 * @returns {Object} Export result with file info
 */
function exportComparisonResults(results, format = 'csv') {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const filename = `FileComparison_${timestamp}`;
  
  // Prepare export data
  const exportData = results.differences.map(diff => ({
    'Change Type': diff.type,
    'Row': diff.row + 1,
    'Column': diff.column !== null ? diff.column + 1 : 'N/A',
    'Column Name': diff.columnName || 'N/A',
    'Key': diff.key || 'N/A',
    'Value A': diff.valueA,
    'Value B': diff.valueB,
    'Description': diff.description
  }));
  
  // Add summary as first rows
  const summaryData = [
    { 'Change Type': 'SUMMARY', 'Description': `Total Differences: ${results.summary.totalDifferences}` },
    { 'Change Type': 'SUMMARY', 'Description': `Modified Cells: ${results.summary.modifiedCells}` },
    { 'Change Type': 'SUMMARY', 'Description': `Added Rows: ${results.summary.addedRows}` },
    { 'Change Type': 'SUMMARY', 'Description': `Deleted Rows: ${results.summary.deletedRows}` },
    { 'Change Type': '---', 'Description': '--- DETAILED CHANGES ---' }
  ];
  
  const allData = [...summaryData, ...exportData];
  
  return exportDataToFormat(allData, {
    format: format,
    filename: filename,
    folderId: null // Save to root for now
  });
}

/**
 * Test function for the core comparison engine
 */
function testFileComparisonEngine() {
  console.log('🧪 Testing File Comparison Engine...');
  
  try {
    // Test with sample data
    const dataA = [
      ['Name', 'Age', 'City'],
      ['John', '25', 'London'],
      ['Jane', '30', 'Manchester'],
      ['Bob', '35', 'Birmingham']
    ];
    
    const dataB = [
      ['Name', 'Age', 'City'],
      ['John', '26', 'London'],        // Age changed
      ['Jane', '30', 'Liverpool'],     // City changed  
      ['Alice', '28', 'Glasgow']       // New person, Bob deleted
    ];
    
    const config = {
      mode: 'positional',
      ignoreCase: false,
      ignoreWhitespace: true
    };
    
    console.log('Testing positional comparison...');
    const results = performCoreComparison(dataA, dataB, config);
    
    console.log('📊 Comparison Results:');
    console.log(`Total differences: ${results.summary.totalDifferences}`);
    console.log(`Modified cells: ${results.summary.modifiedCells}`);
    console.log(`Identical cells: ${results.summary.identicalCells}`);
    
    console.log('📋 Sample differences:');
    results.differences.slice(0, 5).forEach(diff => {
      console.log(`${diff.type}: Row ${diff.row + 1}, Col ${diff.column + 1} - "${diff.valueA}" → "${diff.valueB}"`);
    });
    
    // Test content detection
    console.log('🔍 Testing content detection...');
    const contentType = detectContentType(dataA[0], dataA);
    console.log(`Detected content type: ${contentType}`);
    
    console.log('✅ File Comparison Engine test completed successfully!');
    
    return {
      success: true,
      totalDifferences: results.summary.totalDifferences,
      contentType: contentType
    };
    
  } catch (error) {
    console.error('❌ File Comparison Engine test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
