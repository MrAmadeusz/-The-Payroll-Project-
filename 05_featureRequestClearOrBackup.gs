/**
 * RESET FUNCTION - Clears all feature request data
 * Run this once to reset to clean MVP state
 */
function resetAllFeatureRequestData() {
  try {
    console.log("🧹 Resetting all feature request data...");
    
    // Get the data file
    const file = DriveApp.getFileById(CONFIG.FEATURE_REQUEST_DATA_FILE_ID);
    
    // Create empty array
    const emptyData = [];
    const jsonContent = JSON.stringify(emptyData, null, 2);
    
    // Replace file content with empty array
    const blob = Utilities.newBlob(jsonContent, 'application/json', file.getName());
    file.setContent(blob.getDataAsString());
    
    // Clear cache
    resetFeatureRequestCache();
    
    console.log("✅ All feature request data cleared successfully");
    console.log("📁 File reset to empty array");
    console.log("🔄 Cache cleared");
    
    return {
      success: true,
      message: "All feature request data has been reset",
      previousCount: "All test data removed",
      newCount: 0
    };
    
  } catch (error) {
    console.error("❌ Failed to reset data:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Alternative: Archive existing data before reset (safer)
 */
function resetFeatureRequestDataWithBackup() {
  try {
    console.log("🧹 Resetting data with backup...");
    
    // First, get current data for backup
    const currentData = getRawFeatureRequestData();
    
    if (currentData.length > 0) {
      // Create backup file
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const backupName = `feature_requests_backup_${timestamp}.json`;
      const backupContent = JSON.stringify(currentData, null, 2);
      const backupBlob = Utilities.newBlob(backupContent, 'application/json', backupName);
      
      // Save backup to the same folder
      const folder = DriveApp.getFolderById(CONFIG.FEATURE_REQUEST_FOLDER_ID);
      const backupFile = folder.createFile(backupBlob);
      
      console.log(`💾 Created backup: ${backupFile.getName()}`);
      console.log(`📁 Backup URL: ${backupFile.getUrl()}`);
    }
    
    // Now reset the main file
    const result = resetAllFeatureRequestData();
    
    if (result.success) {
      result.backup = currentData.length > 0 ? `Created backup with ${currentData.length} requests` : 'No backup needed (no data)';
    }
    
    return result;
    
  } catch (error) {
    console.error("❌ Failed to reset with backup:", error.message);
    return {
      success: false,
      error: error.message
    };
  }
}
