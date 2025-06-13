// =================================================================
//                      CONFIGURATION
// =================================================================
const CONFIG = {
    SPREADSHEET_ID: '13FavpJzu9ZP6R3I2svEMToYu29CzhjQBGUmskIKQPJs', // Your Google Sheet ID
    RAW_SHEET_NAME: 'Raw Data',
    CLEAN_SHEET_NAME: 'Clean Data',
    
    // -- UPDATED: GMAIL SEARCH PARAMETERS --
    GMAIL_SEARCH_QUERY: 'has:attachment filename:csv', // The base query to find emails.
    DATE_RANGE_DAYS: 2,                               // Number of days back to search. Set to 0 to search all time.
    PROCESSED_LABEL: 'CSV_Processed',                 // The label applied to emails after processing to prevent re-reading them.
    // ------------------------------------
  
    TIER_COLUMN_NAME: 'Tier', 
    VENDOR_TIER_MAP: { 
      'TIER 1': 'A', 'TIER 2': 'B', 'TIER 3': 'C',
      'TIER 4': 'C', 'TIER 5': 'D', 'TIER 6': 'D'
    }
  };
  
  // =================================================================
  //                      MAIN FUNCTION FOR STEP 1
  // =================================================================
  function step1_ingestAndCleanData() {
    Logger.log("🚀 Starting Step 1: Ingestion & Cleaning...");
    ingestRawCsvFromGmail();
    cleanAndProcessRawData();
    Logger.log("✅ Step 1 Finished Successfully!");
  }
  
  // =================================================================
  //                 PHASE 1: INGEST RAW DATA FROM GMAIL
  // =================================================================
  function ingestRawCsvFromGmail() {
    Logger.log("--- Starting Phase 1: Ingesting from Gmail ---");
    
    // --- NEW: DYNAMIC SEARCH QUERY BUILDING ---
    let searchQuery = CONFIG.GMAIL_SEARCH_QUERY;
  
    // Add the exclusion label to the query to avoid processing emails again
    if (CONFIG.PROCESSED_LABEL) {
      searchQuery += ` -label:${CONFIG.PROCESSED_LABEL}`;
    }
  
    // Add the date range to the query to limit the search scope
    if (CONFIG.DATE_RANGE_DAYS && CONFIG.DATE_RANGE_DAYS > 0) {
      searchQuery += ` newer_than:${CONFIG.DATE_RANGE_DAYS}d`;
    }
  
    Logger.log(`Executing Gmail search with query: "${searchQuery}"`);
    // --------------------------------------------
  
    const threads = GmailApp.search(searchQuery);
  
    if (threads.length === 0) {
      Logger.log('No new emails matching the criteria were found.');
      return;
    }
    Logger.log(`Found ${threads.length} email threads to process.`);
  
    const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const rawSheet = spreadsheet.getSheetByName(CONFIG.RAW_SHEET_NAME);
    if (!rawSheet) throw new Error(`Sheet "${CONFIG.RAW_SHEET_NAME}" not found`);
  
    threads.forEach(thread => {
      thread.getMessages().forEach(message => {
        message.getAttachments().forEach(attachment => {
          if (isCSVFile(attachment)) {
            Logger.log(`Processing CSV: ${attachment.getName()}`);
            const csvContent = attachment.getDataAsString();
            const parsedData = Utilities.parseCsv(csvContent);
  
            if (parsedData.length < 2) return;
  
            const headers = parsedData[0];
            const dataRows = parsedData.slice(1);
            const emailData = { sender: message.getFrom(), subject: message.getSubject(), date: message.getDate(), filename: attachment.getName() };
            
            appendRawData(rawSheet, headers, dataRows, emailData);
          }
        });
      });
      // Mark the entire thread as processed
      thread.addLabel(processedLabel);
    });
  }
  
  /**
   * Appends raw data with metadata to the 'Raw Data' sheet.
   * Adds a "Status" column to track processing.
   */
  function appendRawData(sheet, csvHeaders, dataRows, emailData) {
    const metadataHeaders = ['Ingest Date', 'Sender', 'Subject', 'Filename', 'Original Row'];
    // Prepend an empty string for the "Status" column
    const newRows = dataRows.map((row, index) => ['', emailData.date, emailData.sender, emailData.subject, emailData.filename, index + 2, ...row]); 
  
    if (newRows.length === 0) return;
  
    if (sheet.getLastRow() === 0) {
      const fullHeaders = ['Status', ...metadataHeaders, ...csvHeaders];
      sheet.appendRow(fullHeaders);
    }
  
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    Logger.log(`Appended ${newRows.length} raw rows from ${emailData.filename}`);
  }
  
  
  // =================================================================
  //                 PHASE 2: CLEAN AND PROCESS RAW DATA
  // =================================================================
  function cleanAndProcessRawData() {
    Logger.log("--- Starting Phase 2: Cleaning and Processing Data ---");
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const rawSheet = spreadsheet.getSheetByName(CONFIG.RAW_SHEET_NAME);
    const cleanSheet = spreadsheet.getSheetByName(CONFIG.CLEAN_SHEET_NAME);
    
    const rawDataRange = rawSheet.getDataRange();
    const rawData = rawDataRange.getValues();
    if (rawData.length < 2) {
      Logger.log("No raw data to clean.");
      return;
    }
    
    const rawHeaders = rawData.shift();
    const statusColIndex = rawHeaders.indexOf('Status');
    const tierColIndex = rawHeaders.indexOf(CONFIG.TIER_COLUMN_NAME);
  
    const cleanData = cleanSheet.getDataRange().getValues();
    cleanData.shift(); // Remove headers
    const existingSignatures = new Set(cleanData.map(row => JSON.stringify(row.slice(6).map(normalizeValueForComparison))));
    
    const rowsToProcess = [];
    const rowIndicesToUpdate = [];
  
    rawData.forEach((row, index) => {
      if (row[statusColIndex] === '') { 
        rowsToProcess.push(row);
        rowIndicesToUpdate.push(index + 2);
      }
    });
  
    if (rowsToProcess.length === 0) {
      Logger.log("No new raw data to process.");
      return;
    }
  
    const uniqueNewRows = [];
    rowsToProcess.forEach(row => {
      const dataPortion = row.slice(6); 
      const signature = JSON.stringify(dataPortion.map(normalizeValueForComparison));
      
      if (!existingSignatures.has(signature)) {
        const tierValue = row[tierColIndex] || '';
        const vendorTier = CONFIG.VENDOR_TIER_MAP[tierValue.trim()] || 'N/A';
        const finalRow = [...row, vendorTier];
        uniqueNewRows.push(finalRow);
        existingSignatures.add(signature);
      }
    });
  
    Logger.log(`Found ${uniqueNewRows.length} unique new rows to add.`);
    
    if (uniqueNewRows.length > 0) {
      if (cleanSheet.getLastRow() === 0) {
        cleanSheet.appendRow([...rawHeaders, 'Vendor Tier']);
      }
      cleanSheet.getRange(cleanSheet.getLastRow() + 1, 1, uniqueNewRows.length, uniqueNewRows[0].length).setValues(uniqueNewRows);
    }
    
    const timestamp = new Date();
    rowIndicesToUpdate.forEach(rowIndex => {
      rawSheet.getRange(rowIndex, statusColIndex + 1).setValue(`Processed on ${timestamp.toLocaleString()}`);
    });
    Logger.log(`Updated status for ${rowIndicesToUpdate.length} rows in '${CONFIG.RAW_SHEET_NAME}'.`);
  }
  
  // --- Helper Functions (getOrCreateLabel, isCSVFile, normalizeValueForComparison) ---
  function isCSVFile(attachment) { return attachment.getName().toLowerCase().endsWith('.csv') || attachment.getContentType() === 'text/csv'; }
  function getOrCreateLabel(labelName) { let label = GmailApp.getUserLabelByName(labelName); if (!label) { label = GmailApp.createLabel(labelName); } return label; }
  function normalizeValueForComparison(value) { if (value === null || value === undefined) return ''; if (value instanceof Date) return value.toISOString(); return value.toString().trim(); }