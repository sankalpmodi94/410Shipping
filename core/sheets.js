/**
 * Google Sheets operations for the shipping system
 */

/**
 * Gets a sheet by name from the configured spreadsheet
 * @param {string} sheetName - Name of the sheet to get
 * @returns {Sheet} The requested sheet
 */
function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return sheet;
}

/**
 * Appends data to a sheet with headers if needed
 * @param {string} sheetName - Name of the sheet
 * @param {Array} headers - Array of header strings
 * @param {Array<Array>} rows - 2D array of data rows
 * @param {Object} metadata - Additional metadata to append
 */
function appendDataToSheet(sheetName, headers, rows, metadata = {}) {
  const sheet = getSheet(sheetName);
  const metadataHeaders = Object.keys(metadata);
  
  if (sheet.getLastRow() === 0) {
    // Add headers if sheet is empty
    const fullHeaders = [...metadataHeaders, ...headers];
    sheet.appendRow(fullHeaders);
  }

  // Prepare rows with metadata
  const newRows = rows.map(row => {
    const metadataValues = metadataHeaders.map(key => metadata[key]);
    return [...metadataValues, ...row];
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
         .setValues(newRows);
  }
}

/**
 * Removes duplicate rows from a sheet
 * @param {string} sheetName - Name of the sheet
 * @param {number} metadataColumnCount - Number of metadata columns to exclude from comparison
 */
function removeDuplicatesFromSheet(sheetName, metadataColumnCount = 0) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('No data to process');
    return;
  }

  const headers = data[0];
  const dataRows = data.slice(1);
  const uniqueRows = [];
  
  Logger.log(`Processing ${dataRows.length} rows for duplicates...`);

  for (const currentRow of dataRows) {
    let isDuplicate = false;
    
    for (const uniqueRow of uniqueRows) {
      let isMatch = true;
      
      // Compare data columns only (skip metadata columns)
      for (let col = metadataColumnCount; col < Math.min(currentRow.length, uniqueRow.length); col++) {
        const val1 = UTILS.normalizeValueForComparison(currentRow[col]);
        const val2 = UTILS.normalizeValueForComparison(uniqueRow[col]);
        
        if (val1 !== val2) {
          isMatch = false;
          break;
        }
      }
      
      if (isMatch) {
        isDuplicate = true;
        break;
      }
    }
    
    if (!isDuplicate) {
      uniqueRows.push(currentRow);
    }
  }

  Logger.log(`Removed ${dataRows.length - uniqueRows.length} duplicate rows`);
  Logger.log(`Keeping ${uniqueRows.length} unique rows`);

  // Update sheet with unique rows
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (uniqueRows.length > 0) {
    sheet.getRange(2, 1, uniqueRows.length, uniqueRows[0].length)
         .setValues(uniqueRows);
  }
}

/**
 * Gets data from a sheet as a 2D array
 * @param {string} sheetName - Name of the sheet
 * @returns {Object} Object containing headers and data rows
 */
function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { headers: [], rows: [] };
  }

  return {
    headers: data[0],
    rows: data.slice(1)
  };
}

/**
 * Updates a specific cell in a sheet
 * @param {string} sheetName - Name of the sheet
 * @param {number} row - Row number (1-based)
 * @param {number} col - Column number (1-based)
 * @param {*} value - Value to set
 */
function updateCell(sheetName, row, col, value) {
  const sheet = getSheet(sheetName);
  sheet.getRange(row, col).setValue(value);
}

// Export Sheets operations
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getSheet,
    appendDataToSheet,
    removeDuplicatesFromSheet,
    getSheetData,
    updateCell
  };
} else {
  // For Google Apps Script environment
  global.SHEETS_OPS = {
    getSheet,
    appendDataToSheet,
    removeDuplicatesFromSheet,
    getSheetData,
    updateCell
  };
} 