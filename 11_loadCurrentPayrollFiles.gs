/**
 * COMPLETE PRE-PAYROLL FILE LOADER 
 * Production-ready file loading system for PRE-payroll validation
 * Handles the 4 specific Google Sheets with flexible date handling
 * 
 * FIXES:
 * - Better error handling and recovery
 * - Improved file detection and selection
 * - Enhanced metadata and debugging
 * - Flexible column name matching
 * - Robust spreadsheet reading
 */

/**
 * MAIN FUNCTION: Load current PRE-payroll files for validation
 * FIXED VERSION with comprehensive error handling
 */
function loadCurrentPayrollFiles() {
  const startTime = Date.now();
  
  try {
    console.log("🔍 Loading PRE-payroll files for validation...");
    
    const payrollData = {
      hourlyGrossToNett: [],
      hourlyPaymentsDeductions: [],
      salaryGrossToNett: [],
      salaryPaymentsDeductions: [],
      totalRecords: 0,
      filesSummary: {},
      loadedAt: new Date().toISOString(),
      loadStrategy: 'mostRecent',
      errors: [],
      warnings: []
    };
    
    // Validate folder access first
    let previewFolder;
    try {
      previewFolder = DriveApp.getFolderById(CONFIG.PAYROLL_PREVIEW_FILE_ID);
      console.log(`✅ Successfully accessed folder: ${previewFolder.getName()}`);
    } catch (folderError) {
      throw new Error(`Cannot access payroll preview folder (${CONFIG.PAYROLL_PREVIEW_FILE_ID}): ${folderError.message}`);
    }
    
    // 1. Load Hourly - Gross to Nett file
    console.log("⏰ Loading Hourly - Gross to Nett...");
    try {
      payrollData.hourlyGrossToNett = loadSpecificPayrollFile(previewFolder, "Hourly - Gross to Nett", "hourlyGrossToNett");
      console.log(`✅ Loaded ${payrollData.hourlyGrossToNett.length} Hourly Gross to Nett records`);
    } catch (error) {
      console.warn(`⚠️ Could not load Hourly Gross to Nett: ${error.message}`);
      payrollData.errors.push(`Hourly Gross to Nett: ${error.message}`);
      payrollData.hourlyGrossToNett = [];
    }
    
    // 2. Load Hourly - Payments & Deductions file
    console.log("💰 Loading Hourly - Payments & Deductions...");
    try {
      payrollData.hourlyPaymentsDeductions = loadSpecificPayrollFile(previewFolder, "Hourly - Payments & Deductions", "hourlyPaymentsDeductions");
      console.log(`✅ Loaded ${payrollData.hourlyPaymentsDeductions.length} Hourly Payments & Deductions records`);
    } catch (error) {
      console.warn(`⚠️ Could not load Hourly Payments & Deductions: ${error.message}`);
      payrollData.errors.push(`Hourly Payments & Deductions: ${error.message}`);
      payrollData.hourlyPaymentsDeductions = [];
    }
    
    // 3. Load Salary - Gross to Nett file
    console.log("💼 Loading Salary - Gross to Nett...");
    try {
      payrollData.salaryGrossToNett = loadSpecificPayrollFile(previewFolder, "Salary - Gross to Nett", "salaryGrossToNett");
      console.log(`✅ Loaded ${payrollData.salaryGrossToNett.length} Salary Gross to Nett records`);
    } catch (error) {
      console.warn(`⚠️ Could not load Salary Gross to Nett: ${error.message}`);
      payrollData.errors.push(`Salary Gross to Nett: ${error.message}`);
      payrollData.salaryGrossToNett = [];
    }
    
    // 4. Load Salary - Payments & Deductions file
    console.log("🏢 Loading Salary - Payments & Deductions...");
    try {
      payrollData.salaryPaymentsDeductions = loadSpecificPayrollFile(previewFolder, "Salary - Payments & Deductions", "salaryPaymentsDeductions");
      console.log(`✅ Loaded ${payrollData.salaryPaymentsDeductions.length} Salary Payments & Deductions records`);
    } catch (error) {
      console.warn(`⚠️ Could not load Salary Payments & Deductions: ${error.message}`);
      payrollData.errors.push(`Salary Payments & Deductions: ${error.message}`);
      payrollData.salaryPaymentsDeductions = [];
    }
    
    // Calculate totals and validate
    payrollData.totalRecords = 
      payrollData.hourlyGrossToNett.length + 
      payrollData.hourlyPaymentsDeductions.length + 
      payrollData.salaryGrossToNett.length + 
      payrollData.salaryPaymentsDeductions.length;
    
    // Create summary for debugging and UI display
    payrollData.filesSummary = {
      hourlyGrossToNett: payrollData.hourlyGrossToNett.length,
      hourlyPaymentsDeductions: payrollData.hourlyPaymentsDeductions.length,
      salaryGrossToNett: payrollData.salaryGrossToNett.length,
      salaryPaymentsDeductions: payrollData.salaryPaymentsDeductions.length,
      total: payrollData.totalRecords,
      errorsCount: payrollData.errors.length,
      warningsCount: payrollData.warnings.length
    };
    
    // Validate that we have some data
    if (payrollData.totalRecords === 0) {
      if (payrollData.errors.length > 0) {
        throw new Error(`No payroll data loaded. Errors: ${payrollData.errors.join('; ')}`);
      } else {
        throw new Error("No payroll data found in any of the 4 required files");
      }
    }
    
    // Check if we're missing critical files
    const criticalMissing = [];
    if (payrollData.hourlyGrossToNett.length === 0) criticalMissing.push("Hourly Gross to Nett");
    if (payrollData.salaryGrossToNett.length === 0) criticalMissing.push("Salary Gross to Nett");
    
    if (criticalMissing.length > 0) {
      payrollData.warnings.push(`Missing critical files: ${criticalMissing.join(', ')}`);
      console.warn(`⚠️ Missing critical files: ${criticalMissing.join(', ')}`);
    }
    
    const loadTime = Date.now() - startTime;
    console.log(`📊 PRE-payroll file loading completed in ${loadTime}ms`);
    console.log(`📈 Summary:`, payrollData.filesSummary);
    
    if (payrollData.errors.length > 0) {
      console.warn(`⚠️ Errors encountered:`, payrollData.errors);
    }
    
    // Add performance metadata
    payrollData.performanceMetrics = {
      loadTime: loadTime,
      filesProcessed: 4,
      recordsPerSecond: Math.round(payrollData.totalRecords / (loadTime / 1000)),
      successfulFiles: 4 - payrollData.errors.length
    };
    
    return payrollData;
    
  } catch (error) {
    console.error("❌ Failed to load PRE-payroll files:", error);
    throw new Error(`PRE-payroll file loading failed: ${error.message}`);
  }
}

/**
 * IMPROVED: Load a specific payroll file by name pattern with flexible date handling
 * Uses "most recently modified" strategy to automatically handle changing dates
 */
function loadSpecificPayrollFile(folder, basePattern, dataType) {
  try {
    const files = folder.getFiles();
    const matchingFiles = [];
    
    // Find ALL files matching the base pattern
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // Enhanced pattern matching with multiple strategies
      const isMatch = checkFileNameMatch(fileName, basePattern);
      const isNotExcluded = !fileName.includes('PayrollExceptions') && 
                           !fileName.includes('ARCHIVED') &&
                           !fileName.includes('_OLD') &&
                           !fileName.includes('BACKUP');
      const isGoogleSheet = file.getMimeType() === MimeType.GOOGLE_SHEETS;
      
      if (isMatch && isNotExcluded && isGoogleSheet) {
        matchingFiles.push({
          file: file,
          fileName: fileName,
          lastModified: file.getLastUpdated(),
          dateCreated: file.getDateCreated(),
          size: file.getSize(),
          id: file.getId()
        });
      }
    }
    
    if (matchingFiles.length === 0) {
      throw new Error(`No files found matching pattern: "${basePattern}"`);
    }
    
    // Strategy: Use the MOST RECENTLY MODIFIED file
    const latestFile = matchingFiles.reduce((latest, current) => {
      return current.lastModified > latest.lastModified ? current : latest;
    });
    
    console.log(`📄 Selected ${basePattern} file: ${latestFile.fileName} (modified: ${latestFile.lastModified.toLocaleString()})`);
    
    // Warn if multiple files found (helps with debugging)
    if (matchingFiles.length > 1) {
      console.warn(`⚠️ Found ${matchingFiles.length} files matching "${basePattern}". Using most recent: ${latestFile.fileName}`);
      const otherFiles = matchingFiles
        .filter(f => f !== latestFile)
        .map(f => `${f.fileName} (${f.lastModified.toLocaleString()})`)
        .slice(0, 3); // Show max 3 other files
      console.warn(`📋 Other files found:`, otherFiles);
    }
    
    // Read the selected Google Sheet with optimization for large files
    console.log(`📖 Reading ${latestFile.fileName}...`);
    const fileData = readLargeFileOptimized(latestFile.file);
    
    // Validate file has data
    if (!fileData || fileData.length === 0) {
      console.warn(`⚠️ File ${latestFile.fileName} appears to be empty`);
      return [];
    }
    
    // Add comprehensive metadata to each record
    fileData.forEach(record => {
      record._sourceFile = latestFile.fileName;
      record._sourceType = dataType;
      record._fileId = latestFile.file.getId();
      record._loadedAt = new Date().toISOString();
      record._lastModified = latestFile.lastModified.toISOString();
      record._dateCreated = latestFile.dateCreated.toISOString();
      record._filesFound = matchingFiles.length; // For debugging/auditing
      record._fileSize = latestFile.size;
    });
    
    console.log(`✅ Successfully loaded ${fileData.length} records from ${latestFile.fileName}`);
    
    return fileData;
    
  } catch (error) {
    console.error(`❌ Error loading ${basePattern}:`, error);
    throw new Error(`Failed to load ${basePattern}: ${error.message}`);
  }
}

/**
 * IMPROVED: Check if filename matches the pattern with flexible matching
 */
function checkFileNameMatch(fileName, pattern) {
  const normalizedFileName = fileName.toLowerCase();
  const normalizedPattern = pattern.toLowerCase();
  
  // Direct substring match
  if (normalizedFileName.includes(normalizedPattern)) {
    return true;
  }
  
  // Split pattern and check all parts exist
  const patternParts = normalizedPattern.split(/[\s\-&]+/).filter(part => part.length > 0);
  const allPartsMatch = patternParts.every(part => normalizedFileName.includes(part));
  
  if (allPartsMatch) {
    return true;
  }
  
  // Alternative patterns for common variations
  const alternativePatterns = {
    "hourly - gross to nett": ["hourly gross", "hourly nett", "hourly g2n", "hourly_gross"],
    "hourly - payments & deductions": ["hourly payments", "hourly deductions", "hourly p&d", "hourly_payments"],
    "salary - gross to nett": ["salary gross", "salary nett", "salary g2n", "salary_gross"],
    "salary - payments & deductions": ["salary payments", "salary deductions", "salary p&d", "salary_payments"]
  };
  
  const alternatives = alternativePatterns[normalizedPattern] || [];
  return alternatives.some(alt => normalizedFileName.includes(alt));
}

/**
 * IMPROVED: Optimized file reader for large Google Sheets
 * Reads data in chunks to avoid memory/timeout issues
 */
function readLargeFileOptimized(file) {
  try {
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    const sheet = spreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) {
      console.warn(`⚠️ Sheet appears empty: ${lastRow} rows, ${lastCol} columns`);
      return [];
    }
    
    console.log(`📊 File dimensions: ${lastRow} rows × ${lastCol} columns`);
    
    // Read headers first
    let headers;
    try {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      console.log(`📋 Headers: ${headers.slice(0, 5).join(', ')}${headers.length > 5 ? '...' : ''}`);
    } catch (headerError) {
      console.error(`❌ Failed to read headers: ${headerError.message}`);
      throw new Error(`Cannot read headers from sheet: ${headerError.message}`);
    }
    
    // For very large files (>10k rows), read in batches
    const dataRows = lastRow - 1; // Exclude header
    const data = [];
    
    if (dataRows > 10000) {
      console.log(`📦 Large file detected (${dataRows} rows). Reading in batches...`);
      
      const batchSize = 5000;
      const batches = Math.ceil(dataRows / batchSize);
      
      for (let batch = 0; batch < batches; batch++) {
        const startRow = 2 + (batch * batchSize); // +2 because row 1 is headers and sheets are 1-indexed
        const rowsToRead = Math.min(batchSize, dataRows - (batch * batchSize));
        
        if (rowsToRead <= 0) break;
        
        console.log(`📥 Reading batch ${batch + 1}/${batches}: rows ${startRow} to ${startRow + rowsToRead - 1}`);
        
        try {
          const batchData = sheet.getRange(startRow, 1, rowsToRead, lastCol).getValues();
          data.push(...batchData);
          
          // Small delay to prevent quota issues
          if (batch < batches - 1) {
            Utilities.sleep(100);
          }
        } catch (batchError) {
          console.error(`❌ Failed to read batch ${batch + 1}: ${batchError.message}`);
          throw new Error(`Failed to read data batch ${batch + 1}: ${batchError.message}`);
        }
      }
    } else {
      // For smaller files, read all at once
      console.log(`📥 Reading all ${dataRows} data rows at once`);
      try {
        const allData = sheet.getRange(2, 1, dataRows, lastCol).getValues();
        data.push(...allData);
      } catch (dataError) {
        console.error(`❌ Failed to read data: ${dataError.message}`);
        throw new Error(`Failed to read sheet data: ${dataError.message}`);
      }
    }
    
    // Convert to objects with error handling
    const records = [];
    let successfulRows = 0;
    let errorRows = 0;
    
    data.forEach((row, index) => {
      try {
        const obj = {};
        headers.forEach((header, colIndex) => {
          const cellValue = row[colIndex];
          obj[header] = cellValue !== undefined && cellValue !== null ? cellValue : '';
        });
        records.push(obj);
        successfulRows++;
      } catch (rowError) {
        console.warn(`⚠️ Error processing row ${index + 2}: ${rowError.message}`);
        errorRows++;
      }
    });
    
    console.log(`✅ Successfully processed ${successfulRows} records${errorRows > 0 ? ` (${errorRows} rows had errors)` : ''}`);
    
    if (errorRows > successfulRows * 0.1) { // More than 10% errors
      console.warn(`⚠️ High error rate: ${errorRows} errors out of ${data.length} rows`);
    }
    
    return records;
    
  } catch (error) {
    console.error(`❌ Failed to read file ${file.getName()}: ${error.message}`);
    throw new Error(`Could not read ${file.getName()} as spreadsheet: ${error.message}`);
  }
}

/**
 * IMPROVED: Validate that all required PRE-payroll files are present and accessible
 * Comprehensive validation with detailed error reporting
 */
function validatePrePayrollFilesExist() {
  try {
    console.log("🔍 Validating PRE-payroll files exist...");
    
    const previewFolder = DriveApp.getFolderById(CONFIG.PAYROLL_PREVIEW_FILE_ID);
    const requiredFiles = [
      "Hourly - Gross to Nett",
      "Hourly - Payments & Deductions", 
      "Salary - Gross to Nett",
      "Salary - Payments & Deductions"
    ];
    
    const foundFiles = [];
    const missingFiles = [];
    const duplicateFiles = [];
    
    requiredFiles.forEach(pattern => {
      try {
        const files = previewFolder.getFiles();
        const matchingFiles = [];
        
        while (files.hasNext()) {
          const file = files.next();
          const fileName = file.getName();
          
          if (checkFileNameMatch(fileName, pattern) && 
              !fileName.includes('PayrollExceptions') && 
              !fileName.includes('ARCHIVED') &&
              !fileName.includes('_OLD') &&
              file.getMimeType() === MimeType.GOOGLE_SHEETS) {
            
            matchingFiles.push({
              pattern: pattern,
              fileName: fileName,
              fileId: file.getId(),
              lastModified: file.getLastUpdated(),
              dateCreated: file.getDateCreated(),
              size: file.getSize()
            });
          }
        }
        
        if (matchingFiles.length === 0) {
          missingFiles.push(pattern);
        } else if (matchingFiles.length === 1) {
          foundFiles.push(matchingFiles[0]);
        } else {
          // Multiple files found - flag as potential issue
          const mostRecent = matchingFiles.reduce((latest, current) => {
            return current.lastModified > latest.lastModified ? current : latest;
          });
          
          foundFiles.push(mostRecent);
          duplicateFiles.push({
            pattern: pattern,
            count: matchingFiles.length,
            files: matchingFiles,
            selectedFile: mostRecent.fileName
          });
        }
        
      } catch (error) {
        missingFiles.push(`${pattern} (Error: ${error.message})`);
      }
    });
    
    const allFilesFound = missingFiles.length === 0;
    const hasDuplicates = duplicateFiles.length > 0;
    
    console.log(`📊 File validation: ${foundFiles.length}/${requiredFiles.length} file types found`);
    
    if (hasDuplicates) {
      console.warn(`⚠️ Duplicate files detected for ${duplicateFiles.length} file types`);
      duplicateFiles.forEach(dup => {
        console.warn(`   ${dup.pattern}: ${dup.count} files found, using ${dup.selectedFile}`);
      });
    }
    
    return {
      success: allFilesFound,
      foundFiles: foundFiles,
      missingFiles: missingFiles,
      duplicateFiles: duplicateFiles,
      hasDuplicates: hasDuplicates,
      message: allFilesFound ? 
        (hasDuplicates ? 
          `All required files found, but ${duplicateFiles.length} have duplicates` :
          "All required PRE-payroll files found") : 
        `Missing files: ${missingFiles.join(', ')}`,
      checkedAt: new Date().toISOString(),
      folderChecked: CONFIG.PAYROLL_PREVIEW_FILE_ID,
      recommendations: generateFileRecommendations(allFilesFound, hasDuplicates, missingFiles, duplicateFiles)
    };
    
  } catch (error) {
    console.error("❌ File validation failed:", error);
    return {
      success: false,
      error: error.message,
      message: `File validation failed: ${error.message}`,
      checkedAt: new Date().toISOString(),
      folderChecked: CONFIG.PAYROLL_PREVIEW_FILE_ID
    };
  }
}

/**
 * Generate specific recommendations based on file validation results
 */
function generateFileRecommendations(allFilesFound, hasDuplicates, missingFiles, duplicateFiles) {
  const recommendations = [];
  
  if (!allFilesFound) {
    recommendations.push("❌ Upload missing payroll files to the preview folder");
    missingFiles.forEach(file => {
      recommendations.push(`   - Missing: ${file}`);
    });
  }
  
  if (hasDuplicates) {
    recommendations.push("⚠️ Archive or delete duplicate files to avoid confusion");
    duplicateFiles.forEach(dup => {
      recommendations.push(`   - ${dup.pattern}: ${dup.count} files found`);
    });
  }
  
  if (allFilesFound && !hasDuplicates) {
    recommendations.push("✅ File structure is optimal for validation");
  }
  
  return recommendations;
}

/**
 * DEBUG FUNCTION: Test file loading with detailed diagnostics
 */
function testFileLoadingSystem() {
  console.log("🧪 Testing file loading system...");
  
  try {
    // Step 1: Test folder access
    console.log("1️⃣ Testing folder access...");
    const folder = DriveApp.getFolderById(CONFIG.PAYROLL_PREVIEW_FILE_ID);
    console.log(`✅ Folder access: ${folder.getName()}`);
    
    // Step 2: Test file validation
    console.log("2️⃣ Testing file validation...");
    const validation = validatePrePayrollFilesExist();
    console.log(`✅ File validation: ${validation.success ? 'PASSED' : 'FAILED'}`);
    console.log(`📋 Found: ${validation.foundFiles?.length || 0}, Missing: ${validation.missingFiles?.length || 0}`);
    
    // Step 3: Test actual file loading
    console.log("3️⃣ Testing file loading...");
    const loadResult = loadCurrentPayrollFiles();
    console.log(`✅ File loading: ${loadResult.totalRecords} total records`);
    console.log(`📊 Summary:`, loadResult.filesSummary);
    
    // Step 4: Test data quality
    console.log("4️⃣ Testing data quality...");
    const qualityCheck = checkDataQuality(loadResult);
    console.log(`✅ Data quality: ${qualityCheck.overall}`);
    
    return {
      success: true,
      folderAccess: true,
      fileValidation: validation,
      loadResult: {
        success: true,
        totalRecords: loadResult.totalRecords,
        filesSummary: loadResult.filesSummary,
        errors: loadResult.errors || [],
        warnings: loadResult.warnings || []
      },
      dataQuality: qualityCheck,
      message: "All file loading tests passed"
    };
    
  } catch (error) {
    console.error("❌ File loading test failed:", error);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Check data quality of loaded payroll files
 */
function checkDataQuality(payrollData) {
  const quality = {
    overall: 'good',
    checks: {},
    issues: [],
    recommendations: []
  };
  
  // Check 1: Record counts
  const totalRecords = payrollData.totalRecords;
  quality.checks.recordCount = {
    status: totalRecords > 0 ? 'pass' : 'fail',
    value: totalRecords,
    message: `${totalRecords} total records loaded`
  };
  
  if (totalRecords === 0) {
    quality.overall = 'poor';
    quality.issues.push("No data records found");
  } else if (totalRecords < 50) {
    quality.overall = 'warning';
    quality.issues.push("Very low record count - verify this is expected");
  }
  
  // Check 2: File balance
  const hasHourly = (payrollData.hourlyGrossToNett?.length || 0) > 0;
  const hasSalary = (payrollData.salaryGrossToNett?.length || 0) > 0;
  
  quality.checks.fileBalance = {
    status: hasHourly && hasSalary ? 'pass' : 'warning',
    hasHourly: hasHourly,
    hasSalary: hasSalary,
    message: `Hourly: ${hasHourly ? 'Yes' : 'No'}, Salary: ${hasSalary ? 'Yes' : 'No'}`
  };
  
  if (!hasHourly || !hasSalary) {
    quality.issues.push(`Missing ${!hasHourly ? 'hourly' : 'salary'} employee data`);
    if (quality.overall === 'good') quality.overall = 'warning';
  }
  
  // Check 3: Errors and warnings
  const errorCount = payrollData.errors?.length || 0;
  const warningCount = payrollData.warnings?.length || 0;
  
  quality.checks.loadErrors = {
    status: errorCount === 0 ? 'pass' : (errorCount > 2 ? 'fail' : 'warning'),
    errors: errorCount,
    warnings: warningCount,
    message: `${errorCount} errors, ${warningCount} warnings`
  };
  
  if (errorCount > 0) {
    quality.issues.push(`${errorCount} file loading errors occurred`);
    if (errorCount > 2) quality.overall = 'poor';
    else if (quality.overall === 'good') quality.overall = 'warning';
  }
  
  // Generate recommendations
  if (quality.overall === 'poor') {
    quality.recommendations.push("🔴 Critical issues found - manual review required");
  } else if (quality.overall === 'warning') {
    quality.recommendations.push("🟡 Some issues detected - review before proceeding");
  } else {
    quality.recommendations.push("🟢 Data quality looks good for validation");
  }
  
  if (errorCount > 0) {
    quality.recommendations.push("Check file formats and permissions");
  }
  
  if (!hasHourly || !hasSalary) {
    quality.recommendations.push("Verify all required payroll files are uploaded");
  }
  
  return quality;
}

/**
 * Get comprehensive loading status for system monitoring
 */
function getPayrollDataLoadingStatus() {
  try {
    console.log("📊 Generating comprehensive payroll data loading status...");
    
    const fileValidation = validatePrePayrollFilesExist();
    
    let loadTest = null;
    try {
      const testData = loadCurrentPayrollFiles();
      const qualityCheck = checkDataQuality(testData);
      
      loadTest = {
        success: true,
        totalRecords: testData.totalRecords,
        filesSummary: testData.filesSummary,
        performanceMetrics: testData.performanceMetrics,
        errors: testData.errors || [],
        warnings: testData.warnings || [],
        dataQuality: qualityCheck
      };
    } catch (error) {
      loadTest = {
        success: false,
        error: error.message
      };
    }
    
    return {
      systemStatus: loadTest?.success ? 'healthy' : 'error',
      fileValidation: fileValidation,
      loadTest: loadTest,
      folderAccess: {
        folderId: CONFIG.PAYROLL_PREVIEW_FILE_ID,
        accessible: true // If we got this far, folder is accessible
      },
      checkedAt: new Date().toISOString(),
      recommendations: generateLoadingRecommendations(fileValidation, loadTest)
    };
    
  } catch (error) {
    console.error("❌ Failed to generate loading status:", error);
    return {
      systemStatus: 'error',
      error: error.message,
      checkedAt: new Date().toISOString()
    };
  }
}

/**
 * Generate recommendations based on loading status
 */
function generateLoadingRecommendations(fileValidation, loadTest) {
  const recommendations = [];
  
  if (!fileValidation.success) {
    recommendations.push("❌ Fix missing files before running validation");
    recommendations.push(...fileValidation.recommendations);
  }
  
  if (fileValidation.hasDuplicates) {
    recommendations.push("⚠️ Archive duplicate files to prevent confusion");
  }
  
  if (loadTest && !loadTest.success) {
    recommendations.push("❌ Fix file loading errors before proceeding");
    recommendations.push("Check file permissions and formats");
  }
  
  if (loadTest && loadTest.success) {
    if (loadTest.totalRecords === 0) {
      recommendations.push("⚠️ No data found - verify files contain data");
    } else if (loadTest.totalRecords < 100) {
      recommendations.push("ℹ️ Low record count - verify this is expected");
    }
    
    if (loadTest.errors && loadTest.errors.length > 0) {
      recommendations.push("⚠️ Some files failed to load - check error details");
    }
    
    if (loadTest.dataQuality && loadTest.dataQuality.overall === 'poor') {
      recommendations.push("🔴 Data quality issues detected - manual review needed");
    }
  }
  
  if (recommendations.length === 0) {
    recommendations.push("✅ System operating normally - ready for validation");
  }
  
  return recommendations;
}
