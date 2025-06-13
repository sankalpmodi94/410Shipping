/**
 * Reporting service for the shipping system
 */

/**
 * Main reporting function that generates and sends reports
 */
function generateAndSendReports() {
  Logger.log("🚀 Starting report generation...");
  
  try {
    // Get columns to include in reports
    const colsToSend = SHEETS_OPS.getSheetData(CONFIG.SPREADSHEET.SHEETS.COLS_TO_SEND);
    if (colsToSend.rows.length === 0) {
      Logger.log('No columns specified for reports');
      return;
    }

    const columnsToSend = colsToSend.rows.flat().filter(String);
    if (columnsToSend.length === 0) {
      Logger.log('No valid columns to send');
      return;
    }

    // Get clean data
    const { headers, rows } = SHEETS_OPS.getSheetData(CONFIG.SPREADSHEET.SHEETS.CLEAN_DATA);
    if (rows.length === 0) {
      Logger.log('No data to report');
      return;
    }

    // Find sender column index
    const senderColIndex = headers.indexOf('Sender');
    if (senderColIndex === -1) {
      throw new Error('Sender column not found');
    }

    // Find indices of columns to send
    const indicesToSend = columnsToSend.map(colName => headers.indexOf(colName))
                                     .filter(index => index !== -1);
    const newHeaders = indicesToSend.map(index => headers[index]);

    // Group rows by sender
    const senderGroups = rows.reduce((acc, row) => {
      const sender = row[senderColIndex];
      if (sender) {
        const senderKey = sender.toString().trim();
        if (!acc[senderKey]) acc[senderKey] = [];
        acc[senderKey].push(row);
      }
      return acc;
    }, {});

    // Generate and send reports for each sender
    for (const [sender, senderRows] of Object.entries(senderGroups)) {
      try {
        // Filter rows to include only desired columns
        const filteredRows = senderRows.map(row => 
          indicesToSend.map(index => row[index])
        );

        // Create and send report
        const report = createReport(sender, newHeaders, filteredRows);
        sendReport(sender, report);
        
        // Log the sent report
        logReportSent(sender, filteredRows.length, report.filename);
        
        Logger.log(`✓ Generated and sent report for: ${sender}`);
      } catch (error) {
        Logger.error(`Error processing sender "${sender}":`, error);
      }
    }

    Logger.log("✅ Report generation completed successfully!");
  } catch (error) {
    Logger.error('Error in report generation:', error);
    throw error;
  }
}

/**
 * Creates a report for a sender
 * @param {string} sender - The sender name
 * @param {Array} headers - Report headers
 * @param {Array<Array>} rows - Report data rows
 * @returns {Object} Report object with content and metadata
 */
function createReport(sender, headers, rows) {
  const csvContent = UTILS.createCSV(headers, rows);
  const filename = UTILS.createSanitizedFilename(sender, 'csv');
  
  return {
    content: csvContent,
    filename,
    rowCount: rows.length,
    timestamp: new Date()
  };
}

/**
 * Sends a report via email
 * @param {string} sender - The sender name
 * @param {Object} report - The report object
 */
function sendReport(sender, report) {
  const subject = `Shipping Label for Brijesh's customer: ${sender}`;
  const body = `
Hello,

Please find attached CSV with details of the shipping label for the customer: ${sender}
Please generate ${report.rowCount} labels accordingly.

Best regards,
Narad Muni

---
Auto-generated Message
- Sender: ${sender}
- Number of Labels: ${report.rowCount}
- Export date: ${UTILS.formatDate(report.timestamp)}
  `.trim();

  const blob = Utilities.newBlob(report.content, 'text/csv', report.filename);
  
  GmailApp.sendEmail(
    CONFIG.EMAIL.RECIPIENT,
    subject,
    body,
    {
      attachments: [blob],
      name: CONFIG.EMAIL.SENDER_NAME
    }
  );
}

/**
 * Logs a sent report to the mail log sheet
 * @param {string} sender - The sender name
 * @param {number} rowCount - Number of rows in the report
 * @param {string} filename - Report filename
 */
function logReportSent(sender, rowCount, filename) {
  SHEETS_OPS.appendDataToSheet(
    CONFIG.SPREADSHEET.SHEETS.MAIL_LOG,
    [],
    [[
      new Date(),
      CONFIG.EMAIL.RECIPIENT,
      `Shipping Label for Brijesh's customer: ${sender}`,
      filename,
      rowCount
    ]]
  );
}

// Export reporting service
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    generateAndSendReports,
    createReport,
    sendReport,
    logReportSent
  };
} else {
  // For Google Apps Script environment
  global.REPORTING_SERVICE = {
    generateAndSendReports,
    createReport,
    sendReport,
    logReportSent
  };
} 