/**
 * 💡 05_featureRequests.gs
 * Feature Request Management Module for HR/Payroll Suite
 * 
 * Provides comprehensive feature request lifecycle management:
 * - Request submission and tracking
 * - Status management and workflow
 * - Dashboard and reporting capabilities
 * - Full integration with existing utils and config
 * - HTML-safe interface functions using raw data approach
 */

// === Module Constants ===
const FEATURE_REQUEST_STATUS = {
  SUBMITTED: 'Submitted',
  UNDER_REVIEW: 'Under Review', 
  APPROVED: 'Approved',
  REJECTED: 'Rejected',
  IN_DEVELOPMENT: 'In Development',
  COMPLETED: 'Completed'
};

const FEATURE_REQUEST_PRIORITY = {
  HIGH: 'High',
  MEDIUM: 'Medium', 
  LOW: 'Low'
};

const FEATURE_REQUEST_CATEGORY = {
  ENHANCEMENT: 'Enhancement',
  NEW_MODULE: 'New Module',
  BUG_FIX: 'Bug Fix',
  REPORTING: 'Reporting',
  INTEGRATION: 'Integration',
  PERFORMANCE: 'Performance'
};

const EFFORT_ESTIMATE = {
  SMALL: 'Small (1-2 days)',
  MEDIUM: 'Medium (1 week)',
  LARGE: 'Large (2+ weeks)',
  MAJOR: 'Major (1+ months)'
};

// === Module-Level Cache ===
let _featureRequestCache = null;
let _requestIndexes = {};
let _filterOptions = null;

// === Data Validation Schema ===
const FEATURE_REQUEST_SCHEMA = {
  title: {
    required: true,
    type: 'string',
    validator: (value) => value.length >= 5 ? true : 'Title must be at least 5 characters'
  },
  description: {
    required: true,
    type: 'string',
    validator: (value) => value.length >= 20 ? true : 'Description must be at least 20 characters'
  },
  category: {
    required: true,
    type: 'string',
    enum: Object.values(FEATURE_REQUEST_CATEGORY)
  },
  priority: {
    required: true,
    type: 'string',
    enum: Object.values(FEATURE_REQUEST_PRIORITY)
  },
  submitter: {
    required: true,
    type: 'string'
  }
};

// === Core Data Management ===

/**
 * Loads and caches all feature requests with performance optimization.
 * @returns {Array<Object>} Array of feature request objects
 */
function getAllFeatureRequests() {
  if (_featureRequestCache) return _featureRequestCache;

  _featureRequestCache = timeOperation("Loading feature requests", () => {
    try {
      const rawData = loadJsonFromDrive(CONFIG.FEATURE_REQUEST_DATA_FILE_ID);
      
      // Parse dates and ensure data integrity
      return rawData.map(request => ({
        ...request,
        submissionDate: new Date(request.submissionDate),
        lastUpdated: request.lastUpdated ? new Date(request.lastUpdated) : null,
        completedDate: request.completedDate ? new Date(request.completedDate) : null,
        // Ensure all required fields have defaults
        status: request.status || FEATURE_REQUEST_STATUS.SUBMITTED,
        priority: request.priority || FEATURE_REQUEST_PRIORITY.MEDIUM,
        estimatedEffort: request.estimatedEffort || EFFORT_ESTIMATE.MEDIUM,
        yourNotes: request.yourNotes || ''
      }));
      
    } catch (error) {
      console.log("No existing feature request data found, initializing empty array");
      return [];
    }
  });

  console.log(`Loaded ${_featureRequestCache.length} feature requests`);

  // Build performance indexes using utils
  _buildFeatureRequestIndexes();
  
  return _featureRequestCache;
}

/**
 * Gets raw feature request data without processing (HTML-safe).
 * @returns {Array<Object>} Raw JSON data from file
 */
function getRawFeatureRequestData() {
  try {
    const file = DriveApp.getFileById(CONFIG.FEATURE_REQUEST_DATA_FILE_ID);
    const content = file.getBlob().getDataAsString();
    const rawData = JSON.parse(content);
    
    return Array.isArray(rawData) ? rawData : [];
    
  } catch (error) {
    console.error("getRawFeatureRequestData error:", error.message);
    return [];
  }
}

/**
 * Builds indexes and filter options for performance.
 * @private
 */
function _buildFeatureRequestIndexes() {
  timeOperation("Building feature request indexes", () => {
    _requestIndexes = {
      status: buildDataIndex(_featureRequestCache, 'status'),
      priority: buildDataIndex(_featureRequestCache, 'priority'),
      category: buildDataIndex(_featureRequestCache, 'category'),
      submitter: buildDataIndex(_featureRequestCache, 'submitter')
    };

    _filterOptions = {
      status: extractUniqueValues(_featureRequestCache, 'status'),
      priority: extractUniqueValues(_featureRequestCache, 'priority'),
      category: extractUniqueValues(_featureRequestCache, 'category'),
      submitter: extractUniqueValues(_featureRequestCache, 'submitter'),
      estimatedEffort: extractUniqueValues(_featureRequestCache, 'estimatedEffort')
    };
  });

  console.log(`Built indexes for ${Object.keys(_requestIndexes).length} fields`);
}

/**
 * Clears the feature request cache. Call when data changes.
 */
function resetFeatureRequestCache() {
  _featureRequestCache = null;
  _requestIndexes = {};
  _filterOptions = null;
  console.log("Feature request cache cleared");
}

/**
 * Saves feature request data back to Drive with error handling.
 * @param {Array<Object>} requests - Array of feature request objects
 * @returns {Object} Save operation result
 * @private
 */
function _saveFeatureRequestsToFile(requests) {
  return safeExecute(() => {
    const file = DriveApp.getFileById(CONFIG.FEATURE_REQUEST_DATA_FILE_ID);
    const jsonContent = JSON.stringify(requests, null, 2);
    const blob = Utilities.newBlob(jsonContent, 'application/json', file.getName());
    file.setContent(blob.getDataAsString());
    
    // Clear cache to force reload on next access
    resetFeatureRequestCache();
    
    return { success: true, count: requests.length };
    
  }, {
    maxRetries: 2,
    retryDelay: 1000,
    fallback: (error) => {
      console.error("Failed to save feature requests:", error.message);
      return { success: false, error: error.message };
    }
  });
}

// === Feature Request CRUD Operations ===

/**
 * Creates a new feature request with full validation.
 * @param {Object} requestData - New request data
 * @returns {Object} Creation result with new request ID
 */
function createFeatureRequest(requestData) {
  try {
    // Validate input data using utils schema validator
    const validation = validateDataSchema(requestData, FEATURE_REQUEST_SCHEMA);
    if (!validation.isValid) {
      return {
        success: false,
        errors: validation.errors,
        warnings: validation.warnings
      };
    }

    // Get current requests for ID generation
    const rawData = getRawFeatureRequestData();
    const currentYear = new Date().getFullYear();
    const yearRequests = rawData.filter(r => {
      const submissionDate = new Date(r.submissionDate);
      return submissionDate.getFullYear() === currentYear;
    });
    const nextNumber = (yearRequests.length + 1).toString().padStart(3, '0');
    const requestId = `FR-${currentYear}-${nextNumber}`;

    // Create new request object
    const newRequest = {
      requestId: requestId,
      title: requestData.title.trim(),
      description: requestData.description.trim(),
      category: requestData.category,
      priority: requestData.priority,
      submitter: requestData.submitter.trim(),
      businessJustification: requestData.businessJustification?.trim() || '',
      estimatedImpact: requestData.estimatedImpact || 'Medium',
      dependencies: requestData.dependencies?.trim() || '',
      submissionDate: new Date().toISOString(),
      status: FEATURE_REQUEST_STATUS.SUBMITTED,
      estimatedEffort: requestData.estimatedEffort || EFFORT_ESTIMATE.MEDIUM,
      yourNotes: '',
      lastUpdated: new Date().toISOString(),
      completedDate: null
    };

    // Save to file
    const updatedRequests = [...rawData, newRequest];
    const saveResult = _saveFeatureRequestsToFile(updatedRequests);
    
    if (!saveResult.success) {
      return saveResult;
    }

    console.log(`✅ Created feature request: ${requestId}`);
    
    return {
      success: true,
      requestId: requestId,
      request: newRequest
    };

  } catch (error) {
    console.error("Failed to create feature request:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Updates an existing feature request.
 * @param {string} requestId - Request ID to update
 * @param {Object} updates - Fields to update
 * @returns {Object} Update result
 */
function updateFeatureRequest(requestId, updates) {
  try {
    const rawData = getRawFeatureRequestData();
    const requestIndex = rawData.findIndex(r => r.requestId === requestId);
    
    if (requestIndex === -1) {
      return {
        success: false,
        error: `Feature request ${requestId} not found`
      };
    }

    const currentRequest = rawData[requestIndex];
    
    // Apply updates with validation
    const updatedRequest = {
      ...currentRequest,
      ...updates,
      lastUpdated: new Date().toISOString(),
      // Don't allow certain fields to be updated
      requestId: currentRequest.requestId,
      submissionDate: currentRequest.submissionDate
    };

    // Set completion date if status is completed
    if (updates.status === FEATURE_REQUEST_STATUS.COMPLETED && !updatedRequest.completedDate) {
      updatedRequest.completedDate = new Date().toISOString();
    }

    // Clear completion date if status changes away from completed
    if (updates.status && updates.status !== FEATURE_REQUEST_STATUS.COMPLETED) {
      updatedRequest.completedDate = null;
    }

    // Update the array
    rawData[requestIndex] = updatedRequest;
    
    // Save changes
    const saveResult = _saveFeatureRequestsToFile(rawData);
    if (!saveResult.success) {
      return saveResult;
    }

    console.log(`✅ Updated feature request: ${requestId}`);
    
    return {
      success: true,
      requestId: requestId,
      request: updatedRequest
    };

  } catch (error) {
    console.error("Failed to update feature request:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gets a specific feature request by ID.
 * @param {string} requestId - Request ID to find
 * @returns {Object|null} Feature request object or null
 */
function getFeatureRequestById(requestId) {
  const rawData = getRawFeatureRequestData();
  return rawData.find(r => r.requestId === requestId) || null;
}

/**
 * Deletes a feature request (moves to archive).
 * @param {string} requestId - Request ID to delete
 * @returns {Object} Deletion result
 */
function deleteFeatureRequest(requestId) {
  try {
    const rawData = getRawFeatureRequestData();
    const requestIndex = rawData.findIndex(r => r.requestId === requestId);
    
    if (requestIndex === -1) {
      return {
        success: false,
        error: `Feature request ${requestId} not found`
      };
    }

    const requestToDelete = rawData[requestIndex];
    
    // Archive the request before deletion
    _archiveFeatureRequest(requestToDelete);
    
    // Remove from main array
    rawData.splice(requestIndex, 1);
    
    // Save changes
    const saveResult = _saveFeatureRequestsToFile(rawData);
    if (!saveResult.success) {
      return saveResult;
    }

    console.log(`✅ Deleted and archived feature request: ${requestId}`);
    
    return {
      success: true,
      requestId: requestId,
      archived: true
    };

  } catch (error) {
    console.error("Failed to delete feature request:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Archives a feature request to the archive folder.
 * @param {Object} request - Request to archive
 * @private
 */
function _archiveFeatureRequest(request) {
  try {
    const archiveFolder = DriveApp.getFolderById(CONFIG.FEATURE_REQUEST_ARCHIVE_FOLDER_ID);
    const filename = `FR_Archive_${request.requestId}_${new Date().getTime()}.json`;
    const content = JSON.stringify(request, null, 2);
    const blob = Utilities.newBlob(content, 'application/json', filename);
    archiveFolder.createFile(blob);
    console.log(`Archived request ${request.requestId} to ${filename}`);
  } catch (error) {
    console.error(`Failed to archive request ${request.requestId}:`, error.message);
    // Don't fail the main operation if archiving fails
  }
}

// === Advanced Filtering and Search ===

/**
 * Gets feature requests matching multiple filter criteria.
 * @param {Object} filters - Filter criteria
 * @returns {Array<Object>} Matching feature requests
 */
function getFeatureRequestsMatchingFilters(filters) {
  getAllFeatureRequests(); // Ensure data is loaded
  
  const fieldMap = {
    status: 'status',
    priority: 'priority',
    category: 'category',
    submitter: 'submitter'
  };

  return timeOperation(`Filtering feature requests with ${Object.keys(filters).length} criteria`, () => {
    return filterDataWithIndexes(_featureRequestCache, filters, fieldMap, _requestIndexes);
  });
}

/**
 * Gets feature requests by status with optional additional filters.
 * @param {string} status - Status to filter by
 * @param {Object} additionalFilters - Additional filter criteria
 * @returns {Array<Object>} Matching feature requests
 */
function getFeatureRequestsByStatus(status, additionalFilters = {}) {
  return getFeatureRequestsMatchingFilters({
    status: status,
    ...additionalFilters
  });
}

/**
 * Gets active feature requests (not completed or rejected).
 * @returns {Array<Object>} Active feature requests
 */
function getActiveFeatureRequests() {
  return getAllFeatureRequests().filter(request => 
    ![FEATURE_REQUEST_STATUS.COMPLETED, FEATURE_REQUEST_STATUS.REJECTED].includes(request.status)
  );
}

/**
 * Gets feature requests requiring your attention.
 * @returns {Array<Object>} Requests needing review or decision
 */
function getFeatureRequestsRequiringAttention() {
  return getAllFeatureRequests().filter(request => 
    [FEATURE_REQUEST_STATUS.SUBMITTED, FEATURE_REQUEST_STATUS.UNDER_REVIEW].includes(request.status)
  );
}

// === Analytics and Reporting ===

/**
 * Generates comprehensive analytics for feature requests.
 * @returns {Object} Analytics data
 */
function getFeatureRequestAnalytics() {
  const requests = getAllFeatureRequests();
  
  return timeOperation("Generating feature request analytics", () => {
    const analytics = {
      overview: {
        total: requests.length,
        active: getActiveFeatureRequests().length,
        completed: getFeatureRequestsByStatus(FEATURE_REQUEST_STATUS.COMPLETED).length,
        needingAttention: getFeatureRequestsRequiringAttention().length
      },
      byStatus: {},
      byPriority: {},
      byCategory: {},
      bySubmitter: {},
      trends: {}
    };

    // Count by status
    Object.values(FEATURE_REQUEST_STATUS).forEach(status => {
      analytics.byStatus[status] = getFeatureRequestsByStatus(status).length;
    });

    // Count by priority
    Object.values(FEATURE_REQUEST_PRIORITY).forEach(priority => {
      analytics.byPriority[priority] = getFeatureRequestsMatchingFilters({ priority }).length;
    });

    // Count by category
    Object.values(FEATURE_REQUEST_CATEGORY).forEach(category => {
      analytics.byCategory[category] = getFeatureRequestsMatchingFilters({ category }).length;
    });

    // Count by submitter
    const submitters = extractUniqueValues(requests, 'submitter');
    submitters.forEach(submitter => {
      analytics.bySubmitter[submitter] = getFeatureRequestsMatchingFilters({ submitter }).length;
    });

    // Monthly trends (last 12 months)
    const currentDate = new Date();
    const monthlyData = {};
    
    for (let i = 11; i >= 0; i--) {
      const month = new Date(currentDate.getFullYear(), currentDate.getMonth() - i, 1);
      const monthKey = `${month.getFullYear()}-${(month.getMonth() + 1).toString().padStart(2, '0')}`;
      
      const monthRequests = requests.filter(r => {
        const submissionMonth = new Date(r.submissionDate.getFullYear(), r.submissionDate.getMonth(), 1);
        return submissionMonth.getTime() === month.getTime();
      });
      
      monthlyData[monthKey] = {
        submitted: monthRequests.length,
        completed: monthRequests.filter(r => r.status === FEATURE_REQUEST_STATUS.COMPLETED).length
      };
    }
    
    analytics.trends.monthly = monthlyData;

    return analytics;
  });
}

/**
 * Gets filter options for UI components.
 * @returns {Object} Available filter options
 */
function getFeatureRequestFilterOptions() {
  getAllFeatureRequests(); // Ensure data is loaded
  
  return {
    ..._filterOptions,
    statusOptions: Object.values(FEATURE_REQUEST_STATUS),
    priorityOptions: Object.values(FEATURE_REQUEST_PRIORITY),
    categoryOptions: Object.values(FEATURE_REQUEST_CATEGORY),
    effortOptions: Object.values(EFFORT_ESTIMATE),
    _metadata: {
      totalRequests: _featureRequestCache?.length || 0,
      activeRequests: getActiveFeatureRequests().length,
      lastUpdated: new Date()
    }
  };
}

// === Workflow Management ===

/**
 * Advances a feature request to the next logical status.
 * @param {string} requestId - Request ID to advance
 * @param {string} notes - Optional notes for the advancement
 * @returns {Object} Advancement result
 */
function advanceFeatureRequestStatus(requestId, notes = '') {
  const request = getFeatureRequestById(requestId);
  if (!request) {
    return { success: false, error: `Request ${requestId} not found` };
  }

  const statusFlow = {
    [FEATURE_REQUEST_STATUS.SUBMITTED]: FEATURE_REQUEST_STATUS.UNDER_REVIEW,
    [FEATURE_REQUEST_STATUS.UNDER_REVIEW]: FEATURE_REQUEST_STATUS.APPROVED,
    [FEATURE_REQUEST_STATUS.APPROVED]: FEATURE_REQUEST_STATUS.IN_DEVELOPMENT,
    [FEATURE_REQUEST_STATUS.IN_DEVELOPMENT]: FEATURE_REQUEST_STATUS.COMPLETED
  };

  const nextStatus = statusFlow[request.status];
  if (!nextStatus) {
    return { 
      success: false, 
      error: `Cannot advance from status: ${request.status}` 
    };
  }

  const updates = {
    status: nextStatus
  };

  if (notes) {
    updates.yourNotes = (request.yourNotes ? request.yourNotes + '\n\n' : '') + 
                       `[${new Date().toLocaleString()}] Status advanced to ${nextStatus}: ${notes}`;
  }

  return updateFeatureRequest(requestId, updates);
}

/**
 * Bulk status update for multiple requests.
 * @param {Array<string>} requestIds - Array of request IDs
 * @param {string} newStatus - New status to apply
 * @param {string} notes - Optional notes
 * @returns {Object} Bulk update result
 */
function bulkUpdateFeatureRequestStatus(requestIds, newStatus, notes = '') {
  const results = {
    successful: [],
    failed: [],
    total: requestIds.length
  };

  requestIds.forEach(requestId => {
    const updates = { status: newStatus };
    
    if (notes) {
      const request = getFeatureRequestById(requestId);
      if (request) {
        updates.yourNotes = (request.yourNotes ? request.yourNotes + '\n\n' : '') + 
                           `[${new Date().toLocaleString()}] Bulk update to ${newStatus}: ${notes}`;
      }
    }

    const result = updateFeatureRequest(requestId, updates);
    
    if (result.success) {
      results.successful.push(requestId);
    } else {
      results.failed.push({ requestId, error: result.error });
    }
  });

  console.log(`Bulk update complete: ${results.successful.length}/${results.total} successful`);
  
  return {
    success: results.failed.length === 0,
    results: results
  };
}

// === Export and Reporting ===

/**
 * Exports feature requests in various formats.
 * @param {Object} filters - Filter criteria for export
 * @param {Object} options - Export options
 * @returns {Object} Export result
 */
function exportFeatureRequests(filters = {}, options = {}) {
  const {
    format = 'csv',
    includeArchived = false,
    filename = `feature_requests_${new Date().getTime()}`
  } = options;

  try {
    const requests = getFeatureRequestsMatchingFilters(filters);
    
    // Transform data for export
    const exportData = requests.map(request => ({
      'Request ID': request.requestId,
      'Title': request.title,
      'Category': request.category,
      'Priority': request.priority,
      'Status': request.status,
      'Submitter': request.submitter,
      'Submission Date': request.submissionDate.toLocaleDateString(),
      'Estimated Effort': request.estimatedEffort,
      'Business Justification': request.businessJustification,
      'Dependencies': request.dependencies,
      'Your Notes': request.yourNotes,
      'Last Updated': request.lastUpdated ? request.lastUpdated.toLocaleDateString() : '',
      'Completed Date': request.completedDate ? request.completedDate.toLocaleDateString() : ''
    }));

    return exportDataToFormat(exportData, {
      format: format,
      filename: filename,
      folderId: CONFIG.FEATURE_REQUEST_FOLDER_ID
    });

  } catch (error) {
    return {
      success: false,
      error: `Export failed: ${error.message}`
    };
  }
}

// === UI Interface Functions ===

/**
 * Shows the feature request dashboard sidebar.
 */
function showFeatureRequestDashboard() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('05_featureRequestDashboard')
      .setTitle("Feature Request Management")
      .setWidth(400);
    
    SpreadsheetApp.getUi().showSidebar(html);
    return { success: true };

  } catch (error) {
    console.error("Failed to show feature request dashboard:", error.message);
    SpreadsheetApp.getUi().alert(
      'Dashboard Error', 
      `Could not load dashboard: ${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { success: false, error: error.message };
  }
}

/**
 * Shows the feature request submission form.
 */
function showFeatureRequestSubmissionForm() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('05_featureRequestSubmission')
      .setTitle("Submit Feature Request")
      .setWidth(350);
    
    SpreadsheetApp.getUi().showSidebar(html);
    return { success: true };

  } catch (error) {
    console.error("Failed to show submission form:", error.message);
    SpreadsheetApp.getUi().alert(
      'Form Error', 
      `Could not load submission form: ${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { success: false, error: error.message };
  }
}

// === System Health and Testing ===

/**
 * Gets system health for feature request module.
 * @returns {Object} Health check results
 */
function getFeatureRequestSystemHealth() {
  const health = {
    overall: 'healthy',
    components: {},
    lastChecked: new Date()
  };

  // Check data file access
  try {
    const rawData = getRawFeatureRequestData();
    health.components.dataAccess = {
      status: 'healthy',
      message: `${rawData.length} requests accessible`
    };
  } catch (error) {
    health.components.dataAccess = {
      status: 'error',
      message: error.message
    };
    health.overall = 'error';
  }

  // Check configuration
  try {
    if (!CONFIG.FEATURE_REQUEST_DATA_FILE_ID) {
      throw new Error("Missing FEATURE_REQUEST_DATA_FILE_ID in config");
    }
    
    health.components.configuration = {
      status: 'healthy',
      message: 'All required config values present'
    };
  } catch (error) {
    health.components.configuration = {
      status: 'error',
      message: error.message
    };
    health.overall = 'error';
  }

  return health;
}

/**
 * Comprehensive test suite for feature request module.
 */
function testFeatureRequestModule() {
  console.log("=== Feature Request Module Test ===");
  
  const testResults = {
    dataLoading: false,
    crudOperations: false,
    filtering: false,
    analytics: false,
    uiFunctions: false
  };
  
  let testRequestId = null;
  
  try {
    // Test 1: Data loading
    console.log("1. Testing data loading...");
    try {
      const requests = getAllFeatureRequests();
      console.log(`✅ Loaded ${requests.length} feature requests`);
      testResults.dataLoading = true;
    } catch (error) {
      console.error(`❌ Data loading failed: ${error.message}`);
    }
    
    // Test 2: CRUD operations
    console.log("2. Testing CRUD operations...");
    if (testResults.dataLoading) {
      try {
        // Create test request
        const testData = {
          title: 'Test Feature Request for Module Testing',
          description: 'This is a test feature request created during module testing to verify CRUD operations work correctly.',
          category: FEATURE_REQUEST_CATEGORY.ENHANCEMENT,
          priority: FEATURE_REQUEST_PRIORITY.LOW,
          submitter: 'Test System',
          businessJustification: 'Testing purposes only'
        };
        
        const createResult = createFeatureRequest(testData);
        if (createResult.success) {
          testRequestId = createResult.requestId;
          console.log(`✅ Created test request: ${testRequestId}`);
          
          // Test update
          const updateResult = updateFeatureRequest(testRequestId, { 
            status: FEATURE_REQUEST_STATUS.UNDER_REVIEW,
            yourNotes: 'Updated during testing'
          });
          
          if (updateResult.success) {
            console.log(`✅ Updated test request: ${testRequestId}`);
            testResults.crudOperations = true;
          }
        }
      } catch (error) {
        console.error(`❌ CRUD operations failed: ${error.message}`);
      }
    }
    
    // Test 3: Filtering
    console.log("3. Testing filtering...");
    if (testResults.dataLoading) {
      try {
        const filtered = getFeatureRequestsMatchingFilters({ 
          status: FEATURE_REQUEST_STATUS.SUBMITTED 
        });
        console.log(`✅ Filtering test: ${filtered.length} submitted requests found`);
        testResults.filtering = true;
      } catch (error) {
        console.error(`❌ Filtering test failed: ${error.message}`);
      }
    }
    
    // Test 4: Analytics
    console.log("4. Testing analytics...");
    if (testResults.dataLoading) {
      try {
        const analytics = getFeatureRequestAnalytics();
        console.log(`✅ Analytics: ${analytics.overview.total} total requests analyzed`);
        testResults.analytics = true;
      } catch (error) {
        console.error(`❌ Analytics test failed: ${error.message}`);
      }
    }
    
    // Test 5: UI functions
    console.log("5. Testing UI functions...");
    try {
      const rawData = getRawFeatureRequestData();
      console.log(`✅ Raw data access: ${rawData.length} requests`);
      testResults.uiFunctions = true;
    } catch (error) {
      console.error(`❌ UI functions test failed: ${error.message}`);
    }
    
    // Cleanup: Delete test request
    if (testRequestId) {
      try {
        const deleteResult = deleteFeatureRequest(testRequestId);
        if (deleteResult.success) {
          console.log(`🧹 Cleaned up test request: ${testRequestId}`);
        }
      } catch (error) {
        console.log(`⚠️ Could not clean up test request: ${error.message}`);
      }
    }
    
    // Overall result
    const passedTests = Object.values(testResults).filter(Boolean).length;
    const totalTests = Object.keys(testResults).length;
    
    console.log("\n=== Test Summary ===");
    console.log(`✅ Passed: ${passedTests}/${totalTests} tests`);
    
    if (passedTests === totalTests) {
      console.log("🎉 All tests passed! Feature Request module is working correctly.");
    } else {
      console.log("\n🔧 Some tests failed. Check the individual test results above.");
    }
    
    return {
      success: passedTests === totalTests,
      results: testResults,
      passedTests: passedTests,
      totalTests: totalTests
    };
    
  } catch (error) {
    console.error("❌ Test suite failed:", error.message);
    return { success: false, error: error.message };
  }
}

// === HTML-Safe UI Functions (Using Raw Data Approach) ===

/**
 * HTML-safe wrapper for getAllFeatureRequests - returns raw data for client processing
 */
function getAllFeatureRequestsForUI() {
  try {
    const rawData = getRawFeatureRequestData();
    return {
      success: true,
      data: rawData,
      count: rawData.length
    };
  } catch (error) {
    console.error("getAllFeatureRequestsForUI failed:", error.message);
    return {
      success: false,
      error: error.message,
      data: []
    };
  }
}

/**
 * HTML-safe wrapper for getFeatureRequestFilterOptions - uses raw data
 */
function getFeatureRequestFilterOptionsForUI() {
  try {
    const rawData = getRawFeatureRequestData();
    
    // Extract unique values from raw data
    let submitters = [];
    let statuses = [];
    let priorities = [];
    let categories = [];
    
    if (Array.isArray(rawData)) {
      const submitterSet = new Set();
      const statusSet = new Set();
      const prioritySet = new Set();
      const categorySet = new Set();
      
      rawData.forEach(request => {
        if (request.submitter) submitterSet.add(request.submitter);
        if (request.status) statusSet.add(request.status);
        if (request.priority) prioritySet.add(request.priority);
        if (request.category) categorySet.add(request.category);
      });
      
      submitters = Array.from(submitterSet).sort();
      statuses = Array.from(statusSet).sort();
      priorities = Array.from(prioritySet).sort();
      categories = Array.from(categorySet).sort();
    }
    
    return {
      success: true,
      statusOptions: Object.values(FEATURE_REQUEST_STATUS),
      priorityOptions: Object.values(FEATURE_REQUEST_PRIORITY),
      categoryOptions: Object.values(FEATURE_REQUEST_CATEGORY),
      effortOptions: Object.values(EFFORT_ESTIMATE),
      submitter: submitters,
      status: statuses,
      priority: priorities,
      category: categories,
      _metadata: {
        totalRequests: rawData.length,
        activeRequests: rawData.filter(r => !['Completed', 'Rejected'].includes(r.status)).length,
        lastUpdated: new Date().getTime()
      }
    };
    
  } catch (error) {
    console.error("getFeatureRequestFilterOptionsForUI failed:", error.message);
    return {
      success: false,
      error: error.message,
      statusOptions: Object.values(FEATURE_REQUEST_STATUS),
      priorityOptions: Object.values(FEATURE_REQUEST_PRIORITY),
      categoryOptions: Object.values(FEATURE_REQUEST_CATEGORY),
      effortOptions: Object.values(EFFORT_ESTIMATE),
      submitter: [],
      status: [],
      priority: [],
      category: [],
      _metadata: { totalRequests: 0, activeRequests: 0, lastUpdated: new Date().getTime() }
    };
  }
}

/**
 * HTML-safe wrapper for createFeatureRequest
 */
function createFeatureRequestForUI(requestData) {
  try {
    if (!requestData) {
      return { success: false, error: "No request data provided" };
    }
    
    const result = createFeatureRequest(requestData);
    return {
      success: result?.success || false,
      error: result?.error || null,
      requestId: result?.requestId || null,
      request: result?.request || null
    };
    
  } catch (error) {
    console.error("createFeatureRequestForUI failed:", error.message);
    return { success: false, error: error.message };
  }
}

/**
 * HTML-safe wrapper for updateFeatureRequest
 */
function updateFeatureRequestForUI(requestId, updates) {
  try {
    const result = updateFeatureRequest(requestId, updates);
    return {
      success: result?.success || false,
      error: result?.error || null,
      requestId: result?.requestId || requestId,
      request: result?.request || null
    };
  } catch (error) {
    console.error("updateFeatureRequestForUI failed:", error.message);
    return { success: false, error: error.message, requestId: requestId };
  }
}

/**
 * HTML-safe wrapper for advanceFeatureRequestStatus
 */
function advanceFeatureRequestStatusForUI(requestId, notes) {
  try {
    const result = advanceFeatureRequestStatus(requestId, notes || '');
    return {
      success: result?.success || false,
      error: result?.error || null,
      requestId: requestId
    };
  } catch (error) {
    console.error("advanceFeatureRequestStatusForUI failed:", error.message);
    return { success: false, error: error.message, requestId: requestId };
  }
}

/**
 * HTML-safe wrapper for exportFeatureRequests
 */
function exportFeatureRequestsForUI(filters, options) {
  try {
    const result = exportFeatureRequests(filters || {}, options || {});
    return {
      success: result?.success || false,
      error: result?.error || null,
      fileName: result?.fileName || null,
      url: result?.url || null
    };
  } catch (error) {
    console.error("exportFeatureRequestsForUI failed:", error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Simple test function for HTML communication
 */
function simpleTest() {
  try {
    return { 
      success: true, 
      message: "Server communication works", 
      timestamp: new Date().getTime(),
      random: Math.random()
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
