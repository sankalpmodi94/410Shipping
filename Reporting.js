// =================================================================
//                      CONFIGURATION
// =================================================================
const SENDER_CONFIG = {
    SPREADSHEET_ID: '13FavpJzu9ZP6R3I2svEMToYu29CzhjQBGUmskIKQPJs', // Same Sheet ID
    SHEET_NAME: 'Clean Data',
    RECIPIENT_EMAIL: 'sankalpmodi5@gmail.com',
    SENDER_COLUMN_NAME: 'Sender',
    COLS_TO_SEND_SHEET_NAME: 'Cols to Send', // NEW: Sheet with columns to include in email
    LOG_SHEET_NAME: 'Mail Log'                // NEW: Sheet for logging sent emails
  };
  
  // =================================================================
  //                      MAIN FUNCTION FOR STEP 2
  // =================================================================
  function step2_groupAndEmailReports() {
    Logger.log("🚀 Starting Step 2: Grouping and Emailing...");
    try {
      const spreadsheet = SpreadsheetApp.openById(SENDER_CONFIG.SPREADSHEET_ID);
      const dataSheet = spreadsheet.getSheetByName(SENDER_CONFIG.SHEET_NAME);
      const colsToSendSheet = spreadsheet.getSheetByName(SENDER_CONFIG.COLS_TO_SEND_SHEET_NAME);
      
      if (!dataSheet || !colsToSendSheet) throw new Error("Ensure 'Clean Data' and 'Cols to Send' sheets exist.");
  
      // Get the list of columns to send from the dedicated sheet
      const columnsToSend = colsToSendSheet.getRange('A1:A').getValues().flat().filter(String);
      if (columnsToSend.length === 0) {
        Logger.log("No columns specified in 'Cols to Send' sheet. Halting.");
        return;
      }
  
      const data = dataSheet.getDataRange().getValues();
      if (data.length < 2) {
        Logger.log('No clean data to process.');
        return;
      }
  
      const headers = data.shift();
      const senderColumnIndex = headers.indexOf(SENDER_CONFIG.SENDER_COLUMN_NAME);
      if (senderColumnIndex === -1) throw new Error(`Sender column "${SENDER_CONFIG.SENDER_COLUMN_NAME}" not found`);
  
      // Group rows by sender
      const senderGroups = data.reduce((acc, row) => {
        const sender = row[senderColumnIndex];
        if (sender) {
          const senderKey = sender.toString().trim();
          if (!acc[senderKey]) acc[senderKey] = [];
          acc[senderKey].push(row);
        }
        return acc;
      }, {});
      
      // Find the indices of the columns we want to send
      const indicesToSend = columnsToSend.map(colName => headers.indexOf(colName)).filter(index => index !== -1);
      const newHeaders = indicesToSend.map(index => headers[index]);
  
      for (const [sender, senderRows] of Object.entries(senderGroups)) {
        // Filter each row to include only the desired columns
        const filteredRows = senderRows.map(row => indicesToSend.map(index => row[index]));
  
        const csvContent = createSenderCSV(newHeaders, filteredRows);
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
        const filename = `${sender.replace(/[^a-zA-Z0-9-_]/g, '_')}_data_${timestamp}.csv`;
        const blob = Utilities.newBlob(csvContent, 'text/csv', filename);
        
        sendEmailWithCSV(sender, filteredRows.length, blob);
        logEmailSent(SENDER_CONFIG.LOG_SHEET_NAME, {
          recipient: SENDER_CONFIG.RECIPIENT_EMAIL,
          subject: `Shipping Label for Brijesh's customer: ${sender}`,
          attachmentName: filename,
          rowCount: filteredRows.length
        });
        Logger.log(`✓ Emailed & logged report for: ${sender}`);
      }
  
      Logger.log("✅ Step 2 Finished Successfully!");
    } catch (error) {
      Logger.log(`Error in Step 2: ${error.toString()}`);
    }
  }
  
  // =================================================================
  //                      HELPER FUNCTIONS
  // =================================================================
  
  /**
   * NEW: Logs email details to the specified sheet.
   */
  function logEmailSent(sheetName, details) {
    try {
      const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!logSheet) return; // Fail silently if log sheet doesn't exist
  
      if (logSheet.getLastRow() === 0) {
        logSheet.appendRow(['Timestamp', 'Recipient', 'Subject', 'Attachment Name', 'Row Count']);
      }
      logSheet.appendRow([
        new Date(),
        details.recipient,
        details.subject,
        details.attachmentName,
        details.rowCount
      ]);
    } catch (e) {
      Logger.log(`Could not write to mail log: ${e.toString()}`);
    }
  }
  
  function createSenderCSV(headers, rows) {
    // ... (This function is unchanged)
    const allRows = [headers, ...rows];
    return allRows.map(row => 
      row.map(cell => {
        const value = (cell === null || cell === undefined) ? '' : cell.toString();
        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
          return `"${value.replace(/"/g, '""')}"`;
        }
        return value;
      }).join(',')
    ).join('\n');
  }
  
  function sendEmailWithCSV(sender, rowCount, csvBlob) {
      // ... (This function is unchanged)
      const subject = `Shipping Label for Brijesh's customer: ${sender}`;
      const body = `Hello,\n\nPlease find the attached CSV with details of the shipping label for the customer: ${sender}\nPlease generate ${rowCount} labels accordingly.\n\nBest regards,\nNarad Muni\n---\nAuto-generated Message`.trim();
      GmailApp.sendEmail(SENDER_CONFIG.RECIPIENT_EMAIL, subject, body, { attachments: [csvBlob], name: 'Data Export System' });
  }