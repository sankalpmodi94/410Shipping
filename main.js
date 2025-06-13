/**
 * Main entry point for the shipping system
 */

/**
 * Runs the complete shipping data pipeline
 */
function runShippingPipeline() {
  Logger.log("🚀 Starting shipping data pipeline...");
  
  try {
    // Step 1: Ingest new data
    Logger.log("\n=== Step 1: Data Ingestion ===");
    INGESTION_SERVICE.ingestNewData();

    // Step 2: Clean and process data
    Logger.log("\n=== Step 2: Data Cleaning ===");
    CLEANING_SERVICE.cleanData();

    // Step 3: Generate and send reports
    Logger.log("\n=== Step 3: Report Generation ===");
    REPORTING_SERVICE.generateAndSendReports();

    Logger.log("\n✅ Pipeline completed successfully!");
  } catch (error) {
    Logger.error('Pipeline error:', error);
    throw error;
  }
}

/**
 * Runs only the ingestion step
 */
function runIngestion() {
  Logger.log("🚀 Running ingestion step...");
  INGESTION_SERVICE.ingestNewData();
}

/**
 * Runs only the cleaning step
 */
function runCleaning() {
  Logger.log("🚀 Running cleaning step...");
  CLEANING_SERVICE.cleanData();
}

/**
 * Runs only the reporting step
 */
function runReporting() {
  Logger.log("🚀 Running reporting step...");
  REPORTING_SERVICE.generateAndSendReports();
}

/**
 * Cleans up processed labels from Gmail
 * @param {number} [maxThreads] - Maximum number of threads to process
 */
function cleanupLabels(maxThreads) {
  Logger.log("🧹 Cleaning up processed labels...");
  GMAIL_OPS.cleanupProcessedLabels(maxThreads);
}

/**
 * Removes duplicates from the raw data sheet
 */
function removeDuplicates() {
  Logger.log("🧹 Removing duplicates from raw data...");
  SHEETS_OPS.removeDuplicatesFromSheet(CONFIG.SPREADSHEET.SHEETS.RAW_DATA, 5);
}

/**
 * Tests the system configuration
 */
function testConfiguration() {
  Logger.log("🔍 Testing system configuration...");
  
  try {
    // Test spreadsheet access
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
    Logger.log('✓ Spreadsheet access: OK');
    
    // Test sheet access
    Object.values(CONFIG.SPREADSHEET.SHEETS).forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found`);
      }
      Logger.log(`✓ Sheet "${sheetName}": OK`);
    });
    
    // Test Gmail access
    const label = GmailApp.getUserLabelByName(CONFIG.GMAIL.PROCESSED_LABEL);
    Logger.log('✓ Gmail access: OK');
    
    Logger.log('✅ Configuration test passed!');
  } catch (error) {
    Logger.error('Configuration test failed:', error);
    throw error;
  }
}

// Export main functions
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    runShippingPipeline,
    runIngestion,
    runCleaning,
    runReporting,
    cleanupLabels,
    removeDuplicates,
    testConfiguration
  };
} else {
  // For Google Apps Script environment
  global.SHIPPING_SYSTEM = {
    runShippingPipeline,
    runIngestion,
    runCleaning,
    runReporting,
    cleanupLabels,
    removeDuplicates,
    testConfiguration
  };
} 