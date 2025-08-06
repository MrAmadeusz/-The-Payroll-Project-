/**Central File IDs */

const CONFIG = {

  ADMIN_EMAILS: [
    '@village-hotels.com',
    '@village-hotels.com',
    '@village-hotels.com'
  ],
  HR_EMAILS: [
    '@village.hotels.com',
    '@village-hotels.com',
    '@village-hotels.com'
  ],
  MANAGER_EMAILS: [
    'general_manager@village-hotels.com'
  ],
  LEISURE_EMAILS: [
    'jan.fish@village-hotels.com' //change this to leisure contacts when known and happy with it 
  ],

  REPORT_EMAIL_CONFIG: {
    defaultGroup: 'admin',
    additionalRecipients: {
      'lengthOfServiceMilestones': ['@village-hotels.com'],
      'hoursDemographic': ['@village-hotels.com']
    }
  },

  EMPLOYEE_MASTER_FOLDER_ID: '1VpT76Qp-Yqk1Jubu0aLn-2XRnW7t_Xnn',
  EMPLOYEE_SNAPSHOTS_FOLDER_ID: '1Mw-oOskuuaIJNEGDtoodiR2tw3nrgWyv',
  PAYROLL_CALENDAR_FILE_ID: '1HiyImk-_uTamMzt-I7eUjCSAIZDnJ1dM',
  LOCATION_MAPPING_FILE_ID: '1v1vJ-lYCl2s--xpJ8DqMlMKkA9IQ3Rmr',
  RATE_CARD_FOLDER_ID: '1La7X6XuwsHRmUtAhmJ6qFODbsrIquzlo',
  NMW_RATE_CARD_FILE_ID: '1Ius7C1UUxeVijayDe9UelLOkoY8h9uXz',
  INVESTMENT_JOURNAL_FOLDER_ID: '1tE97vaOVk7hF4l1TumkPhqj9IhqrY4iv',
  INVESTMENT_JOURNAL_OUTPUT_FOLDER_ID: '1c056o5lDIL5NYh6jBVPeilHIE4--h0b2',
  HOURLY_JOURNAL_FOLDER_ID: '1tE97vaOVk7hF4l1TumkPhqj9IhqrY4iv',
  HOURLY_JOURNAL_OUTPUT_FOLDER_ID: '1c056o5lDIL5NYh6jBVPeilHIE4--h0b2',
  SALARIED_JOURNAL_FOLDER_ID: '1tE97vaOVk7hF4l1TumkPhqj9IhqrY4iv',
  SALARIED_JOURNAL_OUTPUT_FOLDER_ID: '1c056o5lDIL5NYh6jBVPeilHIE4--h0b2',
  HOURLY_JOURNAL_FOLDER_ID: '1tE97vaOVk7hF4l1TumkPhqj9IhqrY4iv',
  HOURLY_JOURNAL_OUTPUT_FOLDER_ID: '1c056o5lDIL5NYh6jBVPeilHIE4--h0b2',
  MASS_MAILER_AUDIT_FOLDER_ID: '1GGpuZ_9s-cB8OxQTj0vXeU3V85xllcQA',
  NMW_CHANGE_LOG_FOLDER_ID: '1rUKgB5Q9XmbRwm7IXTKUSZzNJEHmmmaZ',
  RISK_REGISTER_FOLDER_ID: "1ydhMpVrCar-bFUlbCKGPtP30FWd8Lpec",
  PAYROLL_DATA_FOLDER_ID: "1EvR4JFx5JTQ3i9Qfmo76_Ll4XLrPon0U",
  TASK_LIST_FOLDER_ID: "1O5AGW-HWmIDKrVYOxHQR1mZB2elD76qo",
  TASK_AUDIT_LOG_FOLDER_ID: "1WnKkA_3oJlbciln7d6vsLD52cqYct0Au",
  FEATURE_REQUEST_FOLDER_ID: "1mjYdNdTEjpsYNeb-8eoBnZsd0araRe43",
  FEATURE_REQUEST_DATA_FILE_ID: "121SEu_qWfddDe6euwIP22WFTWyJKoO6n", 
  FEATURE_REQUEST_ARCHIVE_FOLDER_ID: "1_kpNB6nHy-rbZy6-IBae8f7nNMxO1yHg",
  MATERNITY_MANAGEMENT_FOLDER_ID: "1tbzfP2xnZJwkCDxvSR3PaZRxCvz10h0v",
  MATERNITY_DATA_FILE_ID: "1MuiSh0qt4LIdJGGPKATpBR6HsQgILu1A",
  HOURLY_PAYMENT_REPORT_FOLDER_ID: "1eZNMg2F_6O28LZJk65R26tBfaU0iSVIx",
  SALARIED_PAYMENT_REPORT_FOLDER_ID: "1Ts10j3FXznIa4qNhnK1GQ3DyVTWILl1k",
  STARTERSLEAVERS_LOG_FOLDER_ID: "1NomMVpdBk9nKyf1kFS_jO7pPbfjuoXc6",
  TEAM_ABSENCE_FILE_ID: "1UGSyWC1-veUEQczdZaNOk6enRu7ACiVr",
  EV_SCHEME_FOLDER_ID: "1LlZ0Opdein3QIqJdOM9Yje1DDbYEWiZb",
  EV_SCHEME_DATA_FILE_ID: "1uNTHzDfluoyhT1iXqUkiD89Q3Ckj3MdI",
  PAYROLL_PREVIEW_FILE_ID: "1xdAgrRnV0J0T-DVSgXFoAg4qxzNQKE4d",
  PAYROLL_PREVIEW_SNAPSHOT_FILE_ID: "11J2QRGiiUaPDfPE4WEULsHb_r_acoJwW",
  PAYROLL_PREVIEW_OUTPUT_FILE_ID: "1eBLufwWMfx8ABXsBDjSI3mnvstdsG2yS",
  PAYROLL_PREVIEW_AUDIT_FILE_ID: '1k13Qc6n6ErQ3-cpS_MsSwkBfjaHm3odG',
  GROSS_TO_NETT_FOLDER_ID: '1YbUfgnrKUL1n0z6loP0zL6RnBb3oeYfS',
  PAYMENTS_AND_DEDUCTIONS_FOLDER_ID: '1rtDg0aMvwOqS-HDjn-Nl5mcmgoq9riGk',
  PAYROLL_UPLOAD_FILES_FOLDER_ID: '1j3HDf5qsbYHF_QCx0I3WZbYoQjf3gc2b',

  CONTRACT_CHECK_EXCLUSIONS: [
    '30015103', '30012632', '30012675', '30014921', '30015069', '30014995', 
    '30015087', '30015107', '30015090', '30014896', '30015041', '30015093', 
    '00011473', '00252584', '30011711', '30014907', '30015043', '30013900', 
    '30012752', '30013349', '30012563', '30012662', '30012657', '30012767', 
    '30015108', '30014984', '30015097', '30014931', '30012725', '30015105', 
    '30014886', '30015081', '30012361', '30012430', '30014868', '30012709', 
    '30012672', '30012684', '30015078', '30012817', '30015075', '30009568', 
    '30011610', '30015049', '30015077', '30002303', '00084930', '00252857', 
    '00005965', '00086725', '20057997', '00252754', '00240966', '00252172', 
    '00238757', '00239729', '30002829', '30001678', '00252852', '20069133', 
    '00251540', '30010238', '20067314', '00087804', '00081702', '30000890', 
    '00252917', '30007376', '00236171', '00006007', '00231466', '00234840', 
    '00082772', '00252956', '00060440', '00086147', '20067326', '30005910', 
    '00005144', '00237174', '00238379', '00060961', '00096107', '00250516', 
    '00234583', '00158772', '00255994', '00167153', '00179096', '00087507', 
    '30002703', '00179363', '30008005', '00236526', '00234129', '00182400', 
    '00236827', '00252569', '00002442', '00087515', '00230200', '00003688', 
    '00178690', '00260592', '00252761', '00262511', '30004172', '00252728', 
    '00006267', '00178181', '30002616', '00262002', '00252824', '00252639', 
    '30010461', '00252855', '00252845', '00235715', '00094306', '00261946', 
    '20067250', '00253251', '30011473', '00238462', '00238479', '00238458'
  ],

  // ================================================================
  // FILE COMPARISON TOOL CONFIGURATION (Module 13_)
  // ================================================================
  
  FILE_COMPARISON: {
    
    // === PERFORMANCE SETTINGS ===
    CACHE_DURATION: 3600000, // 1 hour (in milliseconds)
    MAX_DIFFERENCES_DEFAULT: 5000, // Default max differences to find
    CHUNK_SIZE: 1000, // Rows to process per chunk for large files
    MAX_PROCESSING_TIME: 300000, // 5 minutes max processing time
    
    // === FILE SIZE THRESHOLDS ===
    LARGE_FILE_WARNING: 10 * 1024 * 1024, // 10MB - show "may be slow" warning
    HUGE_FILE_WARNING: 50 * 1024 * 1024,  // 50MB - show "very large file" warning
    MAX_FILE_SIZE: 100 * 1024 * 1024,     // 100MB - reject files larger than this
    
    // === SUPPORTED FILE FORMATS ===
    SUPPORTED_FORMATS: [
      'csv', 
      'xlsx', 
      'xls', 
      'google-sheets'
    ],
    
    MIME_TYPE_MAPPING: {
      'text/csv': 'csv',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
      'application/vnd.ms-excel': 'xls',
      'application/vnd.google-apps.spreadsheet': 'google-sheets'
    },
    
    // === DEFAULT COMPARISON OPTIONS ===
    DEFAULT_OPTIONS: {
      mode: 'positional', // 'positional' or 'key-based'
      ignoreCase: false,
      ignoreWhitespace: true,
      ignoreColumnOrder: false,
      maxDifferences: 5000,
      enableSmartAnalysis: true
    },
    
    // === ANALYSIS MODULES ===
    ANALYSIS_MODULES: {
      'payroll-intelligence': {
        name: 'Payroll Intelligence',
        icon: '💰',
        description: 'Smart payroll analysis with change detection, new starters, leavers, and pay variance analysis',
        function: 'runPayrollAnalysis',
        enabled: true,
        priority: 1
      },
      'sales-analysis': {
        name: 'Sales Analysis',
        icon: '📈',
        description: 'Sales data comparison and trend analysis',
        function: 'runSalesAnalysis',
        enabled: false, // Future module
        priority: 2
      },
      'hr-analysis': {
        name: 'HR Analysis',
        icon: '👥',
        description: 'Employee data change tracking and compliance monitoring',
        function: 'runHRAnalysis',
        enabled: false, // Future module
        priority: 3
      },
      'inventory-analysis': {
        name: 'Inventory Analysis',
        icon: '📦',
        description: 'Stock level changes and inventory variance analysis',
        function: 'runInventoryAnalysis',
        enabled: false, // Future module
        priority: 4
      }
    },
    
    // === PAYROLL ANALYSIS SPECIFIC SETTINGS ===
    PAYROLL_ANALYSIS: {
      
      // Field detection patterns (case-insensitive, partial matches)
      FIELD_PATTERNS: {
        employeeId: ['employee', 'empno', 'id', 'number', 'payroll', 'staff_id'],
        employeeName: ['name', 'employee_name', 'full_name', 'surname', 'forename', 'first_name', 'last_name'],
        grossPay: ['gross', 'gross_pay', 'total_gross', 'basic_pay', 'salary', 'wages'],
        netPay: ['net', 'net_pay', 'take_home', 'total_net', 'net_salary'],
        hours: ['hours', 'worked', 'total_hours', 'contracted', 'actual_hours'],
        overtimeHours: ['overtime', 'ot_hours', 'extra_hours', 'additional_hours'],
        hourlyRate: ['rate', 'hourly_rate', 'pay_rate', 'basic_rate', 'std_rate'],
        department: ['department', 'dept', 'division', 'cost_centre', 'cost_center'],
        location: ['location', 'site', 'hotel', 'workplace', 'branch'],
        startDate: ['start', 'hire', 'employment_start', 'start_date', 'hired_date'],
        endDate: ['end', 'leave', 'termination', 'employment_end', 'end_date', 'left_date'],
        status: ['status', 'employment_status', 'active', 'current', 'employee_status'],
        payType: ['pay_type', 'salary', 'hourly', 'payment_type', 'pay_frequency'],
        pension: ['pension', 'retirement', 'avc', 'workplace_pension', 'ee_pension'],
        tax: ['tax', 'income_tax', 'paye', 'tax_deduction'],
        ni: ['ni', 'national_insurance', 'ee_ni', 'er_ni', 'nic']
      },
      
      // Significance thresholds
      SIGNIFICANT_CHANGE_THRESHOLDS: {
        absoluteAmount: 500,      // £500+ absolute change
        percentageChange: 20,     // 20%+ percentage change
        hourlyRateChange: 2,      // £2+ per hour change
        hoursChange: 10,          // 10+ hours difference
        veryLargeChange: 50       // 50%+ change (potential error)
      },
      
      // Risk assessment weights
      RISK_WEIGHTS: {
        significantChange: 2,     // Points per significant pay change
        criticalWarning: 10,      // Points per critical warning
        highWarning: 5,           // Points per high priority warning
        newStarter: 1,            // Points per new starter
        leaver: 3,                // Points per leaver
        negativeValue: 15,        // Points for negative pay values
        excessiveHours: 5         // Points for >60 hour weeks
      },
      
      // Status keywords for leaver detection
      LEAVER_STATUS_KEYWORDS: [
        'inactive', 'terminated', 'left', 'leaver', 'end', 'finished', 
        'resigned', 'dismissed', 'redundant', 'retired', 'deceased'
      ],
      
      // Working time limits (for compliance warnings)
      WORKING_TIME_LIMITS: {
        maxWeeklyHours: 48,       // EU Working Time Directive
        maxDailyHours: 11,        // Maximum daily working hours
        warningHours: 60,         // Hours that trigger a warning
        criticalHours: 70         // Hours that trigger critical warning
      }
    },
    
    // === EXPORT SETTINGS ===
    EXPORT: {
      DEFAULT_FORMAT: 'csv',
      SUPPORTED_FORMATS: ['csv', 'xlsx', 'json', 'sheet'],
      MAX_INLINE_DIFFERENCES: 1000, // Show in dashboard if less than this
      FILENAME_TEMPLATE: 'FileComparison_{timestamp}',
      
      EXPORT_FOLDER_MAPPING: {
        'payroll': 'PAYROLL_DATA_FOLDER_ID',
        'general': null // Use root folder
      }
    },
    
    // === NOTIFICATION SETTINGS ===
    NOTIFICATIONS: {
      ENABLED: true,
      
      // When to send notifications
      TRIGGERS: {
        significantChanges: true,     // Send when significant changes detected
        criticalWarnings: true,       // Send for critical warnings
        largeComparisons: false,      // Send for comparisons with >10k differences
        failedComparisons: true       // Send when comparisons fail
      },
      
      // Notification thresholds
      THRESHOLDS: {
        significantChangeCount: 10,   // Notify if more than this many significant changes
        criticalWarningCount: 1,      // Always notify for critical warnings
        totalDifferenceCount: 5000    // Notify if total differences exceed this
      },
      
      // Email templates
      TEMPLATES: {
        significantChanges: {
          subject: '⚠️ Significant Changes Detected in File Comparison',
          priority: 'high'
        },
        criticalWarnings: {
          subject: '🚨 Critical Issues Found in Payroll Comparison',
          priority: 'critical'
        },
        comparisonComplete: {
          subject: '✅ File Comparison Complete',
          priority: 'normal'
        }
      }
    },
    
    // === LOGGING AND MONITORING ===
    LOGGING: {
      LEVEL: 'INFO', // 'DEBUG', 'INFO', 'WARN', 'ERROR'
      LOG_PERFORMANCE: true,        // Log performance metrics
      LOG_ERRORS: true,             // Log detailed errors
      LOG_USER_ACTIONS: true,       // Log user interactions
      
      // Performance monitoring
      PERFORMANCE_THRESHOLDS: {
        slowComparison: 30000,      // 30 seconds
        verySlowComparison: 120000, // 2 minutes
        timeoutWarning: 240000      // 4 minutes
      }
    },
    
    // === FEATURE FLAGS ===
    FEATURES: {
      ENABLE_CACHING: true,           // Enable result caching
      ENABLE_PROGRESS_TRACKING: true, // Show progress for long operations
      ENABLE_SMART_ANALYSIS: true,    // Enable intelligent analysis modules
      ENABLE_EXPORT: true,            // Allow result exports
      ENABLE_NOTIFICATIONS: true,     // Send email notifications
      ENABLE_PERFORMANCE_METRICS: true, // Track and display performance
      
      // Experimental features
      EXPERIMENTAL: {
        FUZZY_MATCHING: false,        // Experimental fuzzy row matching
        ML_CONTENT_DETECTION: false,  // Machine learning content type detection
        BATCH_COMPARISONS: false      // Compare multiple files at once
      }
    },
    
    // === DATA FOLDERS FOR FILE DISCOVERY ===
    SCAN_FOLDERS: [
      'EMPLOYEE_MASTER_FOLDER_ID',
      'EMPLOYEE_SNAPSHOTS_FOLDER_ID', 
      'PAYROLL_DATA_FOLDER_ID',
      'GROSS_TO_NETT_FOLDER_ID',
      'PAYMENTS_AND_DEDUCTIONS_FOLDER_ID',
      'HOURLY_PAYMENT_REPORT_FOLDER_ID',
      'SALARIED_PAYMENT_REPORT_FOLDER_ID',
      'PAYROLL_UPLOAD_FILES_FOLDER_ID'
    ],
    
    // === SECURITY SETTINGS ===
    SECURITY: {
      REQUIRE_ADMIN_FOR_SENSITIVE: true, // Require admin access for payroll comparisons
      LOG_ACCESS_ATTEMPTS: true,         // Log file access attempts
      SANITIZE_OUTPUT: true,             // Remove sensitive data from logs
      
      // File access restrictions
      RESTRICTED_FOLDERS: [
        'MATERNITY_DATA_FILE_ID',  // Sensitive HR data
        'EV_SCHEME_DATA_FILE_ID'   // Personal vehicle data
      ]
    }
  },

  // ================================================================
  // HELPER FUNCTIONS FOR FILE COMPARISON CONFIG
  // ================================================================
  
  /**
   * Get file comparison configuration
   * @param {string} section - Config section to retrieve
   * @returns {Object} Configuration object
   */
  getFileComparisonConfig: function(section = null) {
    if (section) {
      return this.FILE_COMPARISON[section] || {};
    }
    return this.FILE_COMPARISON;
  },
  
  /**
   * Check if a feature is enabled
   * @param {string} feature - Feature name
   * @returns {boolean} Whether feature is enabled
   */
  isFileComparisonFeatureEnabled: function(feature) {
    const features = this.FILE_COMPARISON.FEATURES || {};
    return features[feature] === true;
  },
  
  /**
   * Get analysis module configuration
   * @param {string} moduleId - Module identifier
   * @returns {Object|null} Module configuration or null if not found
   */
  getAnalysisModule: function(moduleId) {
    const modules = this.FILE_COMPARISON.ANALYSIS_MODULES || {};
    return modules[moduleId] || null;
  },
  
  /**
   * Get enabled analysis modules
   * @returns {Array} Array of enabled module configurations
   */
  getEnabledAnalysisModules: function() {
    const modules = this.FILE_COMPARISON.ANALYSIS_MODULES || {};
    return Object.entries(modules)
      .filter(([id, config]) => config.enabled === true)
      .map(([id, config]) => ({ id, ...config }))
      .sort((a, b) => a.priority - b.priority);
  },
  
  /**
   * Get folder IDs to scan for files
   * @returns {Array} Array of folder IDs
   */
  getFileComparisonScanFolders: function() {
    const folderKeys = this.FILE_COMPARISON.SCAN_FOLDERS || [];
    return folderKeys.map(key => this[key]).filter(id => id);
  },
  
  /**
   * Check if user has access to file comparison features
   * @param {string} userEmail - User's email address
   * @param {string} feature - Feature to check access for
   * @returns {boolean} Whether user has access
   */
  hasFileComparisonAccess: function(userEmail, feature = 'basic') {
    const isAdmin = this.ADMIN_EMAILS.includes(userEmail);
    const isHR = this.HR_EMAILS.includes(userEmail);
    
    switch (feature) {
      case 'basic':
        return isAdmin || isHR;
      case 'payroll':
        return isAdmin; // Only admins can access payroll comparisons
      case 'sensitive':
        return isAdmin && this.FILE_COMPARISON.SECURITY.REQUIRE_ADMIN_FOR_SENSITIVE;
      default:
        return isAdmin;
    }
  }
};

// ================================================================
// VALIDATION AND HEALTH CHECK FUNCTIONS
// ================================================================

/**
 * Validate the file comparison configuration
 * @returns {Object} Validation results
 */
function validateFileComparisonConfig() {
  const validation = {
    isValid: true,
    errors: [],
    warnings: [],
    info: []
  };
  
  try {
    const config = CONFIG.FILE_COMPARISON;
    
    // Check required sections exist
    const requiredSections = ['ANALYSIS_MODULES', 'DEFAULT_OPTIONS', 'SCAN_FOLDERS'];
    requiredSections.forEach(section => {
      if (!config[section]) {
        validation.errors.push(`Missing required section: ${section}`);
        validation.isValid = false;
      }
    });
    
    // Check folder IDs exist
    const folderIds = CONFIG.getFileComparisonScanFolders();
    if (folderIds.length === 0) {
      validation.warnings.push('No scan folders configured - file discovery may not work');
    }
    
    // Check analysis modules
    const enabledModules = CONFIG.getEnabledAnalysisModules();
    if (enabledModules.length === 0) {
      validation.warnings.push('No analysis modules enabled - smart analysis will not be available');
    } else {
      validation.info.push(`${enabledModules.length} analysis modules enabled: ${enabledModules.map(m => m.name).join(', ')}`);
    }
    
    // Check performance settings
    if (config.CHUNK_SIZE && config.CHUNK_SIZE < 100) {
      validation.warnings.push('CHUNK_SIZE is very small - may impact performance');
    }
    
    if (config.MAX_PROCESSING_TIME && config.MAX_PROCESSING_TIME < 60000) {
      validation.warnings.push('MAX_PROCESSING_TIME is very short - large files may timeout');
    }
    
    // Check cache settings
    if (config.CACHE_DURATION && config.CACHE_DURATION < 300000) {
      validation.warnings.push('CACHE_DURATION is very short - may not provide performance benefits');
    }
    
    validation.info.push('File comparison configuration validation completed');
    
  } catch (error) {
    validation.errors.push(`Configuration validation failed: ${error.message}`);
    validation.isValid = false;
  }
  
  return validation;
}

/**
 * Get file comparison system health status
 * @returns {Object} Health check results
 */
function getFileComparisonSystemHealth() {
  const health = {
    overall: 'healthy',
    components: {},
    lastChecked: new Date(),
    recommendations: []
  };
  
  try {
    // Check configuration
    const configValidation = validateFileComparisonConfig();
    health.components.configuration = {
      status: configValidation.isValid ? 'healthy' : 'error',
      message: configValidation.isValid ? 'Configuration valid' : `${configValidation.errors.length} config errors`,
      details: configValidation
    };
    
    // Check folder access
    let accessibleFolders = 0;
    const totalFolders = CONFIG.getFileComparisonScanFolders().length;
    
    CONFIG.getFileComparisonScanFolders().forEach(folderId => {
      try {
        DriveApp.getFolderById(folderId);
        accessibleFolders++;
      } catch (error) {
        // Folder not accessible
      }
    });
    
    health.components.folderAccess = {
      status: accessibleFolders === totalFolders ? 'healthy' : accessibleFolders > 0 ? 'warning' : 'error',
      message: `${accessibleFolders}/${totalFolders} folders accessible`,
      accessible: accessibleFolders,
      total: totalFolders
    };
    
    // Check required functions exist
    const requiredFunctions = [
      'compareFiles', 'loadFileAsGrid', 'performCoreComparison', 
      'getAvailableFiles', 'runPayrollAnalysis'
    ];
    
    let availableFunctions = 0;
    requiredFunctions.forEach(funcName => {
      try {
        if (typeof eval(funcName) === 'function') {
          availableFunctions++;
        }
      } catch (error) {
        // Function not available
      }
    });
    
    health.components.functions = {
      status: availableFunctions === requiredFunctions.length ? 'healthy' : 'error',
      message: `${availableFunctions}/${requiredFunctions.length} required functions available`,
      available: availableFunctions,
      total: requiredFunctions.length
    };
    
    // Check cache system
    try {
      const cache = CacheService.getScriptCache();
      cache.put('health_check', 'test', 60);
      const retrieved = cache.get('health_check');
      
      health.components.cache = {
        status: retrieved === 'test' ? 'healthy' : 'warning',
        message: retrieved === 'test' ? 'Cache system working' : 'Cache system may have issues'
      };
    } catch (error) {
      health.components.cache = {
        status: 'error',
        message: `Cache system error: ${error.message}`
      };
    }
    
    // Determine overall health
    const componentStatuses = Object.values(health.components).map(c => c.status);
    if (componentStatuses.includes('error')) {
      health.overall = 'error';
      health.recommendations.push('Fix component errors before using file comparison');
    } else if (componentStatuses.includes('warning')) {
      health.overall = 'warning';
      health.recommendations.push('Some components have warnings - monitor for issues');
    }
    
    // Add specific recommendations
    if (health.components.folderAccess.accessible < health.components.folderAccess.total) {
      health.recommendations.push('Check folder permissions for inaccessible folders');
    }
    
    if (health.components.functions.available < health.components.functions.total) {
      health.recommendations.push('Ensure all file comparison modules are properly deployed');
    }
    
  } catch (error) {
    health.overall = 'error';
    health.components.system = {
      status: 'error',
      message: `System health check failed: ${error.message}`
    };
  }
  
  return health;
}

/**
 * Test basic file comparison functionality
 * @returns {Object} Test results
 */
function testFileComparisonBasicFunctionality() {
  console.log('🧪 Testing File Comparison Basic Functionality...');
  
  try {
    // Test configuration access
    const config = CONFIG.getFileComparisonConfig();
    if (!config || Object.keys(config).length === 0) {
      throw new Error('File comparison configuration not found');
    }
    
    // Test enabled modules
    const enabledModules = CONFIG.getEnabledAnalysisModules();
    console.log(`✅ Configuration loaded: ${enabledModules.length} analysis modules enabled`);
    
    // Test folder access
    const scanFolders = CONFIG.getFileComparisonScanFolders();
    let accessibleFolders = 0;
    
    scanFolders.forEach(folderId => {
      try {
        const folder = DriveApp.getFolderById(folderId);
        accessibleFolders++;
      } catch (error) {
        console.warn(`⚠️ Cannot access folder: ${folderId}`);
      }
    });
    
    console.log(`✅ Folder access: ${accessibleFolders}/${scanFolders.length} folders accessible`);
    
    // Test user access functions
    const testEmail = Session.getEffectiveUser().getEmail();
    const hasBasicAccess = CONFIG.hasFileComparisonAccess(testEmail, 'basic');
    const hasPayrollAccess = CONFIG.hasFileComparisonAccess(testEmail, 'payroll');
    
    console.log(`✅ User access: Basic=${hasBasicAccess}, Payroll=${hasPayrollAccess}`);
    
    // Test feature flags
    const cacheEnabled = CONFIG.isFileComparisonFeatureEnabled('ENABLE_CACHING');
    const analysisEnabled = CONFIG.isFileComparisonFeatureEnabled('ENABLE_SMART_ANALYSIS');
    
    console.log(`✅ Features: Caching=${cacheEnabled}, Analysis=${analysisEnabled}`);
    
    return {
      success: true,
      configLoaded: true,
      enabledModules: enabledModules.length,
      accessibleFolders: accessibleFolders,
      totalFolders: scanFolders.length,
      userAccess: {
        basic: hasBasicAccess,
        payroll: hasPayrollAccess
      },
      features: {
        caching: cacheEnabled,
        analysis: analysisEnabled
      }
    };
    
  } catch (error) {
    console.error('❌ File comparison basic functionality test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Print file comparison configuration summary
 */
function printFileComparisonConfigSummary() {
  console.log('📋 FILE COMPARISON CONFIGURATION SUMMARY');
  console.log('========================================');
  
  const config = CONFIG.FILE_COMPARISON;
  
  console.log(`\n🔧 CORE SETTINGS:`);
  console.log(`   Cache Duration: ${config.CACHE_DURATION / 1000}s`);
  console.log(`   Max Differences: ${config.MAX_DIFFERENCES_DEFAULT.toLocaleString()}`);
  console.log(`   Chunk Size: ${config.CHUNK_SIZE.toLocaleString()} rows`);
  console.log(`   Max Processing Time: ${config.MAX_PROCESSING_TIME / 1000}s`);
  
  console.log(`\n📁 FILE SETTINGS:`);
  console.log(`   Large File Warning: ${(config.LARGE_FILE_WARNING / 1024 / 1024).toFixed(1)}MB`);
  console.log(`   Huge File Warning: ${(config.HUGE_FILE_WARNING / 1024 / 1024).toFixed(1)}MB`);
  console.log(`   Supported Formats: ${config.SUPPORTED_FORMATS.join(', ')}`);
  console.log(`   Scan Folders: ${config.SCAN_FOLDERS.length} configured`);
  
  console.log(`\n🧠 ANALYSIS MODULES:`);
  const modules = config.ANALYSIS_MODULES;
  Object.entries(modules).forEach(([id, module]) => {
    const status = module.enabled ? '✅' : '❌';
    console.log(`   ${status} ${module.icon} ${module.name} (Priority: ${module.priority})`);
  });
  
  console.log(`\n💰 PAYROLL ANALYSIS:`);
  const payroll = config.PAYROLL_ANALYSIS;
  console.log(`   Field Patterns: ${Object.keys(payroll.FIELD_PATTERNS).length} types`);
  console.log(`   Significant Change Threshold: £${payroll.SIGNIFICANT_CHANGE_THRESHOLDS.absoluteAmount} or ${payroll.SIGNIFICANT_CHANGE_THRESHOLDS.percentageChange}%`);
  console.log(`   Max Weekly Hours Warning: ${payroll.WORKING_TIME_LIMITS.warningHours}h`);
  
  console.log(`\n🚀 FEATURES:`);
  const features = config.FEATURES;
  Object.entries(features).forEach(([feature, enabled]) => {
    if (typeof enabled === 'boolean') {
      const status = enabled ? '✅' : '❌';
      console.log(`   ${status} ${feature.replace(/_/g, ' ').toLowerCase()}`);
    }
  });
  
  console.log(`\n📧 NOTIFICATIONS:`);
  const notifications = config.NOTIFICATIONS;
  console.log(`   Enabled: ${notifications.ENABLED ? '✅' : '❌'}`);
  console.log(`   Significant Changes Threshold: ${notifications.THRESHOLDS.significantChangeCount}`);
  console.log(`   Critical Warning Threshold: ${notifications.THRESHOLDS.criticalWarningCount}`);
  
  console.log(`\n🔒 SECURITY:`);
  const security = config.SECURITY;
  console.log(`   Admin Required for Sensitive: ${security.REQUIRE_ADMIN_FOR_SENSITIVE ? '✅' : '❌'}`);
  console.log(`   Log Access Attempts: ${security.LOG_ACCESS_ATTEMPTS ? '✅' : '❌'}`);
  console.log(`   Restricted Folders: ${security.RESTRICTED_FOLDERS.length}`);
  
  // Health check
  console.log(`\n🏥 SYSTEM HEALTH CHECK:`);
  const health = getFileComparisonSystemHealth();
  console.log(`   Overall Status: ${health.overall.toUpperCase()}`);
  Object.entries(health.components).forEach(([component, info]) => {
    const statusIcon = info.status === 'healthy' ? '✅' : info.status === 'warning' ? '⚠️' : '❌';
    console.log(`   ${statusIcon} ${component}: ${info.message}`);
  });
  
  if (health.recommendations.length > 0) {
    console.log(`\n💡 RECOMMENDATIONS:`);
    health.recommendations.forEach(rec => console.log(`   • ${rec}`));
  }
  
  console.log('\n========================================');
}

// ================================================================
// LEGACY COMPATIBILITY AND MIGRATION HELPERS
// ================================================================

/**
 * Migrate from old configuration format (if needed)
 * Call this once after updating to ensure compatibility
 */
function migrateFileComparisonConfig() {
  console.log('🔄 Checking for file comparison configuration migration...');
  
  try {
    // Check if old format exists and new doesn't
    if (!CONFIG.FILE_COMPARISON) {
      console.log('⚠️ FILE_COMPARISON config not found - this appears to be a fresh installation');
      return {
        success: true,
        migrated: false,
        message: 'No migration needed - fresh installation'
      };
    }
    
    // Validate current configuration
    const validation = validateFileComparisonConfig();
    
    if (!validation.isValid) {
      console.log('❌ Current configuration has errors:');
      validation.errors.forEach(error => console.log(`   • ${error}`));
      
      return {
        success: false,
        migrated: false,
        errors: validation.errors
      };
    }
    
    console.log('✅ File comparison configuration is up to date');
    
    return {
      success: true,
      migrated: false,
      message: 'Configuration is current'
    };
    
  } catch (error) {
    console.error('❌ Migration check failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Reset file comparison configuration to defaults
 * Use with caution - this will overwrite custom settings
 */
function resetFileComparisonConfigToDefaults() {
  console.log('⚠️ WARNING: This will reset all file comparison settings to defaults');
  console.log('Custom configurations will be lost. Continue? (This is just a warning - function will not actually reset)');
  
  // This is a safety measure - actual reset would require manual confirmation
  return {
    success: false,
    message: 'Reset cancelled for safety. Modify this function to enable actual reset.'
  };
}

// ================================================================
// EXPORT AND BACKUP UTILITIES
// ================================================================

/**
 * Export current file comparison configuration
 * Useful for backup or sharing settings
 */
function exportFileComparisonConfig() {
  try {
    const exportData = {
      exportedAt: new Date().toISOString(),
      version: '1.0',
      configuration: CONFIG.FILE_COMPARISON,
      metadata: {
        totalScanFolders: CONFIG.getFileComparisonScanFolders().length,
        enabledModules: CONFIG.getEnabledAnalysisModules().length,
        userEmail: Session.getEffectiveUser().getEmail()
      }
    };
    
    const jsonString = JSON.stringify(exportData, null, 2);
    
    // Create backup file
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const filename = `FileComparisonConfig_Backup_${timestamp}.json`;
    
    const blob = Utilities.newBlob(jsonString, 'application/json', filename);
    const file = DriveApp.createFile(blob);
    
    console.log(`✅ Configuration exported to: ${filename}`);
    console.log(`📁 File ID: ${file.getId()}`);
    
    return {
      success: true,
      fileId: file.getId(),
      fileName: filename,
      url: file.getUrl()
    };
    
  } catch (error) {
    console.error('❌ Configuration export failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ================================================================
// INITIALIZATION AND SETUP
// ================================================================

/**
 * Initialize file comparison system
 * Run this once after installation to verify everything is working
 */
function initializeFileComparisonSystem() {
  console.log('🚀 INITIALIZING FILE COMPARISON SYSTEM');
  console.log('=====================================');
  
  const results = {
    steps: [],
    success: true,
    errors: [],
    warnings: []
  };
  
  try {
    // Step 1: Validate configuration
    console.log('\n1️⃣ Validating configuration...');
    const configValidation = validateFileComparisonConfig();
    results.steps.push({
      step: 'Configuration Validation',
      success: configValidation.isValid,
      details: configValidation
    });
    
    if (!configValidation.isValid) {
      results.success = false;
      results.errors.push(...configValidation.errors);
    }
    results.warnings.push(...configValidation.warnings);
    
    // Step 2: Check system health
    console.log('\n2️⃣ Checking system health...');
    const health = getFileComparisonSystemHealth();
    results.steps.push({
      step: 'System Health Check',
      success: health.overall !== 'error',
      details: health
    });
    
    if (health.overall === 'error') {
      results.success = false;
      results.errors.push('System health check failed');
    }
    
    // Step 3: Test basic functionality
    console.log('\n3️⃣ Testing basic functionality...');
    const functionalityTest = testFileComparisonBasicFunctionality();
    results.steps.push({
      step: 'Basic Functionality Test',
      success: functionalityTest.success,
      details: functionalityTest
    });
    
    if (!functionalityTest.success) {
      results.success = false;
      results.errors.push(functionalityTest.error);
    }
    
    // Step 4: Check file access
    console.log('\n4️⃣ Checking file access...');
    const scanFolders = CONFIG.getFileComparisonScanFolders();
    let accessibleFolders = 0;
    
    scanFolders.forEach(folderId => {
      try {
        DriveApp.getFolderById(folderId).getName();
        accessibleFolders++;
      } catch (error) {
        results.warnings.push(`Cannot access folder: ${folderId}`);
      }
    });
    
    results.steps.push({
      step: 'File Access Check',
      success: accessibleFolders > 0,
      details: { accessible: accessibleFolders, total: scanFolders.length }
    });
    
    if (accessibleFolders === 0) {
      results.success = false;
      results.errors.push('No scan folders are accessible');
    }
    
    // Final summary
    console.log('\n📊 INITIALIZATION SUMMARY');
    console.log('========================');
    console.log(`Overall Status: ${results.success ? '✅ SUCCESS' : '❌ FAILED'}`);
    console.log(`Steps Completed: ${results.steps.filter(s => s.success).length}/${results.steps.length}`);
    
    if (results.errors.length > 0) {
      console.log(`\n❌ ERRORS (${results.errors.length}):`);
      results.errors.forEach(error => console.log(`   • ${error}`));
    }
    
    if (results.warnings.length > 0) {
      console.log(`\n⚠️ WARNINGS (${results.warnings.length}):`);
      results.warnings.forEach(warning => console.log(`   • ${warning}`));
    }
    
    if (results.success) {
      console.log('\n🎉 File Comparison System is ready for use!');
      console.log('Access it via the menu: Admin Tools → File Comparison Tool');
    } else {
      console.log('\n⚠️ Please fix the errors above before using the File Comparison Tool');
    }
    
    return results;
    
  } catch (error) {
    console.error('❌ Initialization failed:', error);
    results.success = false;
    results.errors.push(`Initialization error: ${error.message}`);
    return results;
  }
}
