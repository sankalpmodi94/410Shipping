/**
 * Cleaning service for the shipping system
 */

/**
 * Main cleaning function that processes raw data and creates clean data
 */
function cleanData() {
  Logger.log("🚀 Starting data cleaning...");
  
  try {
    const { headers, rows } = SHEETS_OPS.getSheetData(CONFIG.SPREADSHEET.SHEETS.RAW_DATA);
    if (rows.length === 0) {
      Logger.log('No raw data to clean');
      return;
    }

    // Find important column indices
    const tierColIndex = headers.indexOf(CONFIG.DATA.TIER_COLUMN_NAME);
    if (tierColIndex === -1) {
      throw new Error(`Tier column "${CONFIG.DATA.TIER_COLUMN_NAME}" not found`);
    }

    // Process each row
    const cleanedRows = rows.map(row => {
      const tierValue = row[tierColIndex] || '';
      const vendorTier = CONFIG.DATA.VENDOR_TIER_MAP[tierValue.trim()] || 'N/A';
      return [...row, vendorTier];
    });

    // Get existing clean data to check for duplicates
    const existingData = SHEETS_OPS.getSheetData(CONFIG.SPREADSHEET.SHEETS.CLEAN_DATA);
    const existingSignatures = new Set(
      existingData.rows.map(row => 
        JSON.stringify(row.slice(6).map(UTILS.normalizeValueForComparison))
      )
    );

    // Filter out duplicates
    const uniqueNewRows = cleanedRows.filter(row => {
      const signature = JSON.stringify(row.slice(6).map(UTILS.normalizeValueForComparison));
      if (!existingSignatures.has(signature)) {
        existingSignatures.add(signature);
        return true;
      }
      return false;
    });

    Logger.log(`Found ${uniqueNewRows.length} unique new rows to add`);

    if (uniqueNewRows.length > 0) {
      // Add headers if clean sheet is empty
      if (existingData.headers.length === 0) {
        SHEETS_OPS.appendDataToSheet(
          CONFIG.SPREADSHEET.SHEETS.CLEAN_DATA,
          [...headers, 'Vendor Tier'],
          uniqueNewRows
        );
      } else {
        SHEETS_OPS.appendDataToSheet(
          CONFIG.SPREADSHEET.SHEETS.CLEAN_DATA,
          [],
          uniqueNewRows
        );
      }
    }

    // Update status in raw data sheet
    const statusColIndex = headers.indexOf('Status');
    if (statusColIndex !== -1) {
      const timestamp = new Date();
      rows.forEach((_, index) => {
        SHEETS_OPS.updateCell(
          CONFIG.SPREADSHEET.SHEETS.RAW_DATA,
          index + 2,
          statusColIndex + 1,
          `Processed on ${UTILS.formatDate(timestamp)}`
        );
      });
    }

    Logger.log("✅ Data cleaning completed successfully!");
  } catch (error) {
    Logger.error('Error in data cleaning:', error);
    throw error;
  }
}

/**
 * Validates and cleans a single data row
 * @param {Array} row - The row to clean
 * @param {Array} headers - The headers for the row
 * @returns {Array} The cleaned row
 */
function cleanRow(row, headers) {
  return row.map((value, index) => {
    const header = headers[index];
    
    // Apply specific cleaning rules based on header
    switch (header) {
      case CONFIG.DATA.TIER_COLUMN_NAME:
        return (value || '').trim().toUpperCase();
      
      case 'Email':
        return (value || '').trim().toLowerCase();
      
      case 'Phone':
        // Remove non-numeric characters
        return (value || '').replace(/\D/g, '');
      
      case 'Address':
        // Standardize address format
        return (value || '')
          .trim()
          .replace(/\s+/g, ' ')
          .replace(/,/g, ', ');
      
      default:
        return value;
    }
  });
}

// Export cleaning service
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    cleanData,
    cleanRow
  };
} else {
  // For Google Apps Script environment
  global.CLEANING_SERVICE = {
    cleanData,
    cleanRow
  };
} 