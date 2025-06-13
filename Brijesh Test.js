/**
 * Manual cleanup function to remove processed labels
 */
function cleanupProcessedLabels() {
    try {
      const label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
      if (label) {
        const threads = label.getThreads();
        threads.forEach(thread => thread.removeLabel(label));
        console.log(`Removed processed label from ${threads.length} threads`);
      }
    } catch (error) {
      console.error('Error in cleanup:', error);
    }
  }
  
  /**
   * Manual function to remove all duplicate rows from the sheet
   */
  function removeDuplicatesFromSheet() {
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
      const existingData = sheet.getDataRange().getValues();
      
      if (existingData.length <= 1) {
        console.log('No data to process');
        return;
      }
      
      const headers = existingData[0];
      const dataRows = existingData.slice(1);
      const uniqueRows = [];
      const metadataColumnCount = 5;
      
      console.log(`Processing ${dataRows.length} rows for duplicates...`);
      
      for (let i = 0; i < dataRows.length; i++) {
        const currentRow = dataRows[i];
        let isDuplicate = false;
        
        // Check against already processed unique rows
        for (const uniqueRow of uniqueRows) {
          let isMatch = true;
          
          // Compare data columns only (skip first 5 metadata columns)
          for (let col = metadataColumnCount; col < Math.min(currentRow.length, uniqueRow.length); col++) {
            const val1 = normalizeValueForComparison(currentRow[col]);
            const val2 = normalizeValueForComparison(uniqueRow[col]);
            
            if (val1 !== val2) {
              isMatch = false;
              break;
            }
          }
          
          if (isMatch) {
            isDuplicate = true;
            console.log(`Duplicate found at row ${i + 2}: ${currentRow.slice(metadataColumnCount).join(' | ')}`);
            break;
          }
        }
        
        if (!isDuplicate) {
          uniqueRows.push(currentRow);
        }
      }
      
      console.log(`Removed ${dataRows.length - uniqueRows.length} duplicate rows`);
      console.log(`Keeping ${uniqueRows.length} unique rows`);
      
      // Clear sheet and write unique data
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (uniqueRows.length > 0) {
        sheet.getRange(2, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
      }
      
      console.log('Sheet updated with unique rows only');
      
    } catch (error) {
      console.error('Error removing duplicates:', error);
    }
  }/**
   * Gmail CSV to Google Sheets Script
   * Scans Gmail for emails with CSV attachments and appends data to a Google Sheet
   */
  
  // Configuration - Update these values
  const CONFIG = {
    SPREADSHEET_ID: '13FavpJzu9ZP6R3I2svEMToYu29CzhjQBGUmskIKQPJs', // Replace with your Google Sheet ID
    SHEET_NAME: 'Raw Import', // Replace with your sheet name
    GMAIL_SEARCH_QUERY: 'has:attachment filename:csv', // Modify search criteria as needed
    MAX_EMAILS_TO_PROCESS: 10, // Limit emails processed per run
    DATE_RANGE_DAYS: 1, // Only process emails from last N days
    PROCESSED_LABEL: 'CSV_Processed', // Label to mark processed emails
    HEADER_VALIDATION: true, // Validate CSV headers match sheet headers
    DUPLICATE_CHECK: true, // Check for duplicate rows before adding
    DUPLICATE_CHECK_COLUMNS: [] // Specify columns to check for duplicates (empty = all columns)
  };
  
  /**
   * Main function to scan emails and process CSV attachments
   */
  function scanEmailsAndProcessCSV() {
    try {
      console.log('Starting email scan...');
      
      // Get or create the processed label
      const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
      
      // Search for emails with CSV attachments that haven't been processed
      const searchQuery = `${CONFIG.GMAIL_SEARCH_QUERY} -label:${CONFIG.PROCESSED_LABEL} newer_than:${CONFIG.DATE_RANGE_DAYS}d`;
      const threads = GmailApp.search(searchQuery, 0, CONFIG.MAX_EMAILS_TO_PROCESS);
      
      console.log(`Found ${threads.length} email threads to process`);
      
      if (threads.length === 0) {
        console.log('No new emails with CSV attachments found');
        return;
      }
      
      // Get the target spreadsheet
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
      
      if (!sheet) {
        throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found`);
      }
      
      let totalProcessed = 0;
      
      // Process each thread
      for (const thread of threads) {
        const messages = thread.getMessages();
        
        for (const message of messages) {
          const processed = processEmailMessage(message, sheet);
          if (processed) {
            totalProcessed++;
          }
        }
        
        // Mark thread as processed
        thread.addLabel(processedLabel);
      }
      
      console.log(`Processing complete. Processed ${totalProcessed} CSV files`);
      
    } catch (error) {
      console.error('Error in scanEmailsAndProcessCSV:', error);
      throw error;
    }
  }
  
  /**
   * Process a single email message for CSV attachments
   */
  function processEmailMessage(message, sheet) {
    try {
      const attachments = message.getAttachments();
      let processed = false;
      
      for (const attachment of attachments) {
        if (isCSVFile(attachment)) {
          console.log(`Processing CSV: ${attachment.getName()}`);
          
          // Parse CSV data
          const csvData = parseCSVAttachment(attachment);
          
          if (csvData && csvData.data && csvData.data.length > 0) {
            // Add metadata columns
            const emailData = {
              sender: message.getFrom(),
              subject: message.getSubject(),
              date: message.getDate(),
              filename: attachment.getName()
            };
            
            // Append data to sheet
            appendCSVDataToSheet(sheet, csvData, emailData);
            processed = true;
            
            console.log(`Successfully processed ${csvData.data.length} rows from ${attachment.getName()}`);
          }
        }
      }
      
      return processed;
      
    } catch (error) {
      console.error('Error processing email message:', error);
      return false;
    }
  }
  
  /**
   * Check if attachment is a CSV file
   */
  function isCSVFile(attachment) {
    const fileName = attachment.getName().toLowerCase();
    const contentType = attachment.getContentType();
    
    return fileName.endsWith('.csv') || 
           contentType === 'text/csv' || 
           contentType === 'application/csv';
  }
  
  /**
   * Parse CSV attachment data
   */
  function parseCSVAttachment(attachment) {
    try {
      const csvContent = attachment.getDataAsString();
      const rows = csvContent.split('\n').filter(row => row.trim() !== '');
      
      if (rows.length === 0) {
        return null;
      }
      
      // Parse CSV rows
      const parsedData = rows.map(row => parseCSVRow(row));
      
      // Return object with headers and data
      return {
        headers: parsedData[0] || [],
        data: parsedData.slice(1) // Skip header row
      };
      
    } catch (error) {
      console.error('Error parsing CSV:', error);
      return null;
    }
  }
  
  /**
   * Parse a single CSV row handling quoted fields
   */
  function parseCSVRow(row) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < row.length; i++) {
      const char = row[i];
      
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    
    result.push(current.trim());
    return result;
  }
  
  /**
   * Append CSV data to the Google Sheet
   */
  function appendCSVDataToSheet(sheet, csvData, emailData) {
    try {
      // Get current data and headers
      const existingData = sheet.getDataRange().getValues();
      const isFirstData = existingData.length <= 1 && existingData[0].every(cell => cell === '');
      
      let sheetHeaders = [];
      let dataStartRow = 1;
      
      if (isFirstData) {
        // Create headers if sheet is empty
        const metadataHeaders = ['Email Date', 'Sender', 'Subject', 'Filename', 'Row Number'];
        sheetHeaders = [...metadataHeaders, ...csvData.headers];
        
        sheet.clear();
        sheet.getRange(1, 1, 1, sheetHeaders.length).setValues([sheetHeaders]);
        dataStartRow = 2;
        
        console.log('Created new headers:', sheetHeaders);
      } else {
        // Get existing headers
        sheetHeaders = existingData[0];
        dataStartRow = existingData.length + 1;
        
        // Validate CSV headers match sheet headers (excluding metadata columns)
        if (CONFIG.HEADER_VALIDATION) {
          validateHeaders(csvData.headers, sheetHeaders);
        }
      }
      
      // Get ALL existing data for comprehensive duplicate checking
      const allExistingData = sheet.getDataRange().getValues();
      console.log(`=== DUPLICATE CHECK INFO ===`);
      console.log(`Total existing rows in sheet: ${allExistingData.length} (including header)`);
      console.log(`Data rows to check against: ${Math.max(0, allExistingData.length - 1)}`);
      console.log(`New CSV rows to process: ${csvData.data.length}`);
      
      // Prepare data with metadata
      const newRows = [];
      const duplicateCount = { count: 0 };
      
      csvData.data.forEach((row, index) => {
        const newRow = [
          emailData.date,
          emailData.sender,
          emailData.subject,
          emailData.filename,
          index + 1, // Row number within CSV
          ...row
        ];
        
        console.log(`\n--- Processing CSV row ${index + 1} of ${csvData.data.length} ---`);
        
        // Check for duplicates if enabled
        if (CONFIG.DUPLICATE_CHECK) {
          // Use allExistingData to ensure we check against ALL rows in the sheet
          if (!isDuplicateRow(newRow, allExistingData, sheetHeaders, duplicateCount)) {
            newRows.push(newRow);
            console.log(`✅ Row ${index + 1} added (unique)`);
          } else {
            console.log(`❌ Row ${index + 1} skipped (duplicate)`);
          }
        } else {
          newRows.push(newRow);
          console.log(`➕ Row ${index + 1} added (duplicate check disabled)`);
        }
      });
      
      if (newRows.length === 0) {
        console.log(`All ${csvData.data.length} rows were duplicates. No data added.`);
        if (duplicateCount.count > 0) {
          console.log(`Found ${duplicateCount.count} duplicate rows that were skipped.`);
        }
        return;
      }
      
      // Add new rows to sheet
      if (newRows.length > 0) {
        sheet.getRange(dataStartRow, 1, newRows.length, newRows[0].length).setValues(newRows);
        console.log(`Added ${newRows.length} new rows to sheet.`);
        
        if (duplicateCount.count > 0) {
          console.log(`Skipped ${duplicateCount.count} duplicate rows.`);
        }
      }
      
      // Format the new data
      formatNewData(sheet, newRows.length);
      
    } catch (error) {
      console.error('Error appending data to sheet:', error);
      throw error;
    }
  }
  
  /**
   * Validate that CSV headers match Google Sheet headers
   */
  function validateHeaders(csvHeaders, sheetHeaders) {
    // Extract data headers from sheet (excluding metadata columns)
    const metadataColumnCount = 5; // Email Date, Sender, Subject, Filename, Row Number
    const sheetDataHeaders = sheetHeaders.slice(metadataColumnCount);
    
    // Compare headers
    if (csvHeaders.length !== sheetDataHeaders.length) {
      throw new Error(`Header count mismatch: CSV has ${csvHeaders.length} columns, Sheet expects ${sheetDataHeaders.length} columns`);
    }
    
    for (let i = 0; i < csvHeaders.length; i++) {
      const csvHeader = (csvHeaders[i] || '').toString().trim();
      const sheetHeader = (sheetDataHeaders[i] || '').toString().trim();
      
      if (csvHeader !== sheetHeader) {
        throw new Error(`Header mismatch at column ${i + 1}: CSV has "${csvHeader}", Sheet expects "${sheetHeader}"`);
      }
    }
    
    console.log('✓ CSV headers match sheet headers');
  }
  
  /**
   * Check if a row is a duplicate of existing data
   * This function checks against ALL existing rows in the Google Sheet
   */
  function isDuplicateRow(newRow, allExistingData, sheetHeaders, duplicateCount) {
    // Skip header row in existing data - get ALL data rows from the sheet
    const dataRows = allExistingData.slice(1);
    
    // Always skip the first 5 metadata columns for duplicate checking
    const metadataColumnCount = 5; // Email Date, Sender, Subject, Filename, Row Number
    
    // Determine which columns to check for duplicates (starting from column 6)
    let columnsToCheck = [];
    if (CONFIG.DUPLICATE_CHECK_COLUMNS.length > 0) {
      // Use specified columns (but ensure they're beyond the metadata columns)
      columnsToCheck = CONFIG.DUPLICATE_CHECK_COLUMNS.map(colName => {
        const index = sheetHeaders.indexOf(colName);
        if (index === -1) {
          console.warn(`Duplicate check column "${colName}" not found in headers`);
          return -1;
        }
        if (index < metadataColumnCount) {
          console.warn(`Duplicate check column "${colName}" is a metadata column and will be skipped`);
          return -1;
        }
        return index;
      }).filter(index => index !== -1);
    } else {
      // Check all data columns starting from column 6 (index 5)
      columnsToCheck = Array.from(
        {length: Math.max(0, newRow.length - metadataColumnCount)}, 
        (_, i) => i + metadataColumnCount
      );
    }
    
    if (columnsToCheck.length === 0) {
      console.warn('No valid columns found for duplicate checking');
      return false;
    }
    
    // Debug logging
    console.log(`🔍 Checking against ALL ${dataRows.length} existing rows in the sheet`);
    console.log(`New row data columns (6+): ${newRow.slice(metadataColumnCount).join(' | ')}`);
    console.log(`Columns to check (indices): ${columnsToCheck.join(', ')}`);
    
    // Check against EVERY existing row in the sheet
    for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
      const existingRow = dataRows[rowIndex];
      let isMatch = true;
      let matchDetails = [];
      
      for (const colIndex of columnsToCheck) {
        // Ensure both rows have data at this column index
        if (colIndex >= newRow.length || colIndex >= existingRow.length) {
          isMatch = false;
          matchDetails.push(`Col${colIndex}: LENGTH_MISMATCH`);
          break;
        }
        
        // Normalize values for comparison
        const newValue = normalizeValueForComparison(newRow[colIndex]);
        const existingValue = normalizeValueForComparison(existingRow[colIndex]);
        
        if (newValue !== existingValue) {
          isMatch = false;
          matchDetails.push(`Col${colIndex}: "${newValue}" ≠ "${existingValue}"`);
          break;
        } else {
          matchDetails.push(`Col${colIndex}: "${newValue}" = "${existingValue}"`);
        }
      }
      
      if (isMatch) {
        duplicateCount.count++;
        console.log(`🚨 DUPLICATE FOUND at existing sheet row ${rowIndex + 2}`);
        console.log(`Match details: ${matchDetails.join(', ')}`);
        console.log(`New row data: ${newRow.slice(metadataColumnCount).join(' | ')}`);
        console.log(`Existing row data: ${existingRow.slice(metadataColumnCount).join(' | ')}`);
        console.log(`--- DUPLICATE CONFIRMED - SKIPPING NEW ROW ---`);
        return true;
      } else {
        // Log detailed comparison for first few rows to help debug
        if (rowIndex < 2) {
          console.log(`Row ${rowIndex + 2} comparison: ${matchDetails.slice(0, 3).join(', ')}`);
        }
      }
    }
    
    console.log(`✅ No duplicate found among all ${dataRows.length} existing rows`);
    return false;
  }
  
  /**
   * Normalize values for consistent comparison
   */
  function normalizeValueForComparison(value) {
    if (value === null || value === undefined) {
      return '';
    }
    
    // Convert to string and normalize
    let normalized = value.toString().trim();
    
    // Handle dates - convert to ISO string for consistent comparison
    if (value instanceof Date) {
      normalized = value.toISOString();
    }
    
    // Handle numbers - ensure consistent formatting
    if (typeof value === 'number' || (!isNaN(value) && !isNaN(parseFloat(value)) && value.toString().trim() !== '')) {
      const num = parseFloat(value);
      if (!isNaN(num)) {
        normalized = num.toString();
      }
    }
    
    return normalized;
  }
  
  /**
   * Format newly added data
   */
  function formatNewData(sheet, numRows) {
    try {
      const lastRow = sheet.getLastRow();
      const startRow = lastRow - numRows + 1;
      const numCols = sheet.getLastColumn();
      
      // Format date column
      if (numCols > 0) {
        sheet.getRange(startRow, 1, numRows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }
      
      // Add alternating row colors for better readability
      const range = sheet.getRange(startRow, 1, numRows, numCols);
      range.setBorder(true, true, true, true, true, true);
      
    } catch (error) {
      console.error('Error formatting data:', error);
    }
  }
  
  /**
   * Get or create a Gmail label
   */
  function getOrCreateLabel(labelName) {
    let label = GmailApp.getUserLabelByName(labelName);
    
    if (!label) {
      label = GmailApp.createLabel(labelName);
      console.log(`Created new label: ${labelName}`);
    }
    
    return label;
  }
  
  /**
   * Setup function to create triggers and initialize
   */
  function setupScript() {
    try {
      // Delete existing triggers
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'scanEmailsAndProcessCSV') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      
      // Create new time-based trigger (runs every hour)
      ScriptApp.newTrigger('scanEmailsAndProcessCSV')
        .timeBased()
        .everyHours(1)
        .create();
      
      console.log('Setup complete. Script will run every hour.');
      console.log('You can also run scanEmailsAndProcessCSV() manually.');
      
    } catch (error) {
      console.error('Error in setup:', error);
      throw error;
    }
  }
  
  /**
   * Test function to verify configuration
   */
  function testConfiguration() {
    try {
      console.log('Testing configuration...');
      
      // Test spreadsheet access
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      console.log(`✓ Spreadsheet found: ${spreadsheet.getName()}`);
      
      // Test sheet access
      const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
      if (!sheet) {
        throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found`);
      }
      console.log(`✓ Sheet found: ${CONFIG.SHEET_NAME}`);
      
      // Test Gmail access
      const testThreads = GmailApp.search('in:inbox', 0, 1);
      console.log(`✓ Gmail access confirmed`);
      
      console.log('✓ Configuration test passed!');
      
    } catch (error) {
      console.error('Configuration test failed:', error);
      throw error;
    }
  }
  
  /**
   * Test function to debug duplicate detection
   */
  function debugDuplicateDetection() {
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
      const existingData = sheet.getDataRange().getValues();
      
      console.log('=== DUPLICATE DETECTION DEBUG ===');
      console.log(`Sheet has ${existingData.length} total rows (including header)`);
      console.log(`Headers: ${existingData[0].join(' | ')}`);
      
      if (existingData.length > 1) {
        console.log('\n--- First few data rows ---');
        for (let i = 1; i < Math.min(4, existingData.length); i++) {
          console.log(`Row ${i + 1}: ${existingData[i].slice(5).join(' | ')}`);
        }
        
        // Check for duplicates within existing data
        console.log('\n--- Checking existing data for duplicates ---');
        const duplicatesFound = [];
        for (let i = 1; i < existingData.length; i++) {
          for (let j = i + 1; j < existingData.length; j++) {
            const row1Data = existingData[i].slice(5);
            const row2Data = existingData[j].slice(5);
            
            let isMatch = true;
            for (let k = 0; k < Math.min(row1Data.length, row2Data.length); k++) {
              const val1 = normalizeValueForComparison(row1Data[k]);
              const val2 = normalizeValueForComparison(row2Data[k]);
              if (val1 !== val2) {
                isMatch = false;
                break;
              }
            }
            
            if (isMatch) {
              duplicatesFound.push(`Rows ${i + 1} and ${j + 1} are duplicates`);
            }
          }
        }
        
        if (duplicatesFound.length > 0) {
          console.log('🚨 EXISTING DUPLICATES FOUND:');
          duplicatesFound.forEach(dup => console.log(dup));
        } else {
          console.log('✅ No duplicates found in existing data');
        }
      }
      
    } catch (error) {
      console.error('Debug error:', error);
    }
  }