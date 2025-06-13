/**
 * Ingestion service for the shipping system
 */

/**
 * Main ingestion function that processes new data from Gmail
 */
function ingestNewData() {
  Logger.log("🚀 Starting data ingestion...");
  
  try {
    const threads = GMAIL_OPS.searchGmailThreads();
    if (threads.length === 0) {
      Logger.log('No new emails to process');
      return;
    }

    Logger.log(`Found ${threads.length} email threads to process`);

    for (const thread of threads) {
      const attachments = GMAIL_OPS.processGmailThread(thread);
      
      for (const { attachment, sender, subject, date, filename } of attachments) {
        if (UTILS.isCSVFile(attachment)) {
          processCSVAttachment(attachment, { sender, subject, date, filename });
        }
      }

      GMAIL_OPS.markThreadAsProcessed(thread);
    }

    Logger.log("✅ Data ingestion completed successfully!");
  } catch (error) {
    Logger.error('Error in data ingestion:', error);
    throw error;
  }
}

/**
 * Processes a CSV attachment and stores it in the raw data sheet
 * @param {GmailAttachment} attachment - The CSV attachment
 * @param {Object} metadata - Email metadata
 */
function processCSVAttachment(attachment, metadata) {
  try {
    Logger.log(`Processing CSV: ${attachment.getName()}`);
    
    const csvContent = attachment.getDataAsString();
    const parsedData = Utilities.parseCsv(csvContent);

    if (parsedData.length < 2) {
      Logger.log('CSV file has no data rows');
      return;
    }

    const headers = parsedData[0];
    const dataRows = parsedData.slice(1);

    // Add metadata
    const emailMetadata = {
      'Ingest Date': metadata.date,
      'Sender': metadata.sender,
      'Subject': metadata.subject,
      'Filename': metadata.filename,
      'Original Row': null // Will be set for each row
    };

    // Append to raw data sheet
    SHEETS_OPS.appendDataToSheet(
      CONFIG.SPREADSHEET.SHEETS.RAW_DATA,
      headers,
      dataRows.map((row, index) => {
        emailMetadata['Original Row'] = index + 2;
        return row;
      }),
      emailMetadata
    );

    Logger.log(`Successfully processed ${dataRows.length} rows from ${metadata.filename}`);
  } catch (error) {
    Logger.error(`Error processing CSV ${metadata.filename}:`, error);
  }
}

/**
 * Processes a PDF attachment and extracts its text
 * @param {GmailAttachment} attachment - The PDF attachment
 * @param {Object} metadata - Email metadata
 */
function processPdfAttachment(attachment, metadata) {
  try {
    Logger.log(`Processing PDF: ${attachment.getName()}`);
    
    const extractedText = GMAIL_OPS.extractTextFromPdf(attachment);
    
    // Store in PDF data sheet
    SHEETS_OPS.appendDataToSheet(
      CONFIG.SPREADSHEET.SHEETS.PDF_DATA,
      ['Extracted Text'],
      [[extractedText]],
      {
        'Email Subject': metadata.subject,
        'Sender': metadata.sender,
        'Email Date': metadata.date,
        'PDF Filename': metadata.filename
      }
    );

    Logger.log(`Successfully processed PDF ${metadata.filename}`);
  } catch (error) {
    Logger.error(`Error processing PDF ${metadata.filename}:`, error);
  }
}

// Export ingestion service
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    ingestNewData,
    processCSVAttachment,
    processPdfAttachment
  };
} else {
  // For Google Apps Script environment
  global.INGESTION_SERVICE = {
    ingestNewData,
    processCSVAttachment,
    processPdfAttachment
  };
} 