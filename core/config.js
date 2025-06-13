/**
 * Centralized configuration for the shipping system
 */

const CONFIG = {
  // Spreadsheet Configuration
  SPREADSHEET: {
    ID: '13FavpJzu9ZP6R3I2svEMToYu29CzhjQBGUmskIKQPJs',
    SHEETS: {
      RAW_DATA: 'Raw Data',
      CLEAN_DATA: 'Clean Data',
      COLS_TO_SEND: 'Cols to Send',
      MAIL_LOG: 'Mail Log',
      RAW_IMPORT: 'Raw Import'
    }
  },

  // Gmail Configuration
  GMAIL: {
    SEARCH_QUERY: 'has:attachment filename:csv',
    DATE_RANGE_DAYS: 2,
    PROCESSED_LABEL: 'CSV_Processed',
    MAX_EMAILS_TO_PROCESS: 10
  },

  // Data Processing Configuration
  DATA: {
    TIER_COLUMN_NAME: 'Tier',
    VENDOR_TIER_MAP: {
      'TIER 1': 'A', 'TIER 2': 'B', 'TIER 3': 'C',
      'TIER 4': 'C', 'TIER 5': 'D', 'TIER 6': 'D'
    },
    HEADER_VALIDATION: true,
    DUPLICATE_CHECK: true,
    DUPLICATE_CHECK_COLUMNS: []
  },

  // Email Configuration
  EMAIL: {
    RECIPIENT: 'sankalpmodi5@gmail.com',
    SENDER_NAME: 'Data Export System'
  },

  // PDF Processing Configuration
  PDF: {
    TARGET_SHEET_NAME: 'Sheet 2',
    OCR_LANGUAGE: 'en'
  }
};

// Export configuration
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { CONFIG };
} else {
  // For Google Apps Script environment
  global.CONFIG = CONFIG;
} 