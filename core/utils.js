/**
 * Shared utility functions for the shipping system
 */

/**
 * Normalizes a value for comparison by converting to string and trimming
 * @param {*} value - The value to normalize
 * @returns {string} Normalized string value
 */
function normalizeValueForComparison(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString();
  return value.toString().trim();
}

/**
 * Creates or retrieves a Gmail label
 * @param {string} labelName - Name of the label
 * @returns {GmailLabel} The Gmail label object
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

/**
 * Checks if an attachment is a CSV file
 * @param {GmailAttachment} attachment - The attachment to check
 * @returns {boolean} True if the attachment is a CSV file
 */
function isCSVFile(attachment) {
  return attachment.getName().toLowerCase().endsWith('.csv') || 
         attachment.getContentType() === 'text/csv';
}

/**
 * Creates a CSV string from headers and rows
 * @param {Array} headers - Array of header strings
 * @param {Array<Array>} rows - 2D array of data rows
 * @returns {string} CSV formatted string
 */
function createCSV(headers, rows) {
  function escapeCSVValue(value) {
    if (value === null || value === undefined) return '';
    const stringValue = value.toString();
    if (stringValue.includes(',') || stringValue.includes('\n') || stringValue.includes('"')) {
      return `"${stringValue.replace(/"/g, '""')}"`;
    }
    return stringValue;
  }

  const csvLines = [
    headers.map(escapeCSVValue).join(','),
    ...rows.map(row => row.map(escapeCSVValue).join(','))
  ];
  
  return csvLines.join('\n');
}

/**
 * Formats a date for display
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Creates a sanitized filename
 * @param {string} baseName - The base name for the file
 * @param {string} extension - The file extension
 * @returns {string} Sanitized filename with timestamp
 */
function createSanitizedFilename(baseName, extension) {
  const sanitizedBase = baseName.replace(/[^a-zA-Z0-9-_]/g, '_');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  return `${sanitizedBase}_${timestamp}.${extension}`;
}

/**
 * Validates if headers match expected format
 * @param {Array} csvHeaders - Headers from CSV
 * @param {Array} expectedHeaders - Expected headers
 * @returns {boolean} True if headers match
 */
function validateHeaders(csvHeaders, expectedHeaders) {
  if (!Array.isArray(csvHeaders) || !Array.isArray(expectedHeaders)) return false;
  if (csvHeaders.length !== expectedHeaders.length) return false;
  
  return csvHeaders.every((header, index) => 
    normalizeValueForComparison(header) === normalizeValueForComparison(expectedHeaders[index])
  );
}

// Export utilities
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    normalizeValueForComparison,
    getOrCreateLabel,
    isCSVFile,
    createCSV,
    formatDate,
    createSanitizedFilename,
    validateHeaders
  };
} else {
  // For Google Apps Script environment
  global.UTILS = {
    normalizeValueForComparison,
    getOrCreateLabel,
    isCSVFile,
    createCSV,
    formatDate,
    createSanitizedFilename,
    validateHeaders
  };
} 