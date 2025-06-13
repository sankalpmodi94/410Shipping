/**
 * Scans recent Gmail messages for PDF attachments, extracts their text using Google Drive's OCR,
 * and pushes the content to "Sheet 2" of the active Google Sheet.
 * Configuration allows limiting the number of emails processed and the date range.
 *
 * IMPORTANT: This function requires the 'Google Drive Advanced Service' to be enabled
 * in your Apps Script project.
 *
 * It no longer relies on the external 'pdfToText' library.
 */
function myFunction() {
    // --- Configuration Settings ---
    // Maximum number of emails to process per run.
    const MAX_EMAILS_TO_PROCESS = 10;
    // Only process emails from the last N days.
    const DATE_RANGE_DAYS = 1; // Example: 1 means today and yesterday
    // Name of the Google Sheet tab (sheet) where data will be pushed.
    const TARGET_SHEET_NAME = "Sheet 2";
  
    Logger.log(`Starting PDF extraction and pushing to Google Sheet:`);
    Logger.log(`- Method: Google Drive OCR`);
    Logger.log(`- Target Sheet: "${TARGET_SHEET_NAME}"`);
    Logger.log(`- Max emails to process: ${MAX_EMAILS_TO_PROCESS}`);
    Logger.log(`- Date range: Last ${DATE_RANGE_DAYS} day(s)`);
    Logger.log('---');
  
    // --- Google Sheet Setup ---
    let sheet;
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
  
      // If 'Sheet 2' doesn't exist, create it.
      if (!sheet) {
        sheet = spreadsheet.insertSheet(TARGET_SHEET_NAME);
        Logger.log(`Created new sheet: "${TARGET_SHEET_NAME}"`);
      }
  
      // Add headers if the sheet is empty
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Email Subject', 'Sender', 'Email Date', 'PDF Filename', 'Extracted Text']);
        Logger.log('Added headers to the sheet.');
      }
    } catch (e) {
      Logger.log(`Error setting up Google Sheet: ${e.message}`);
      Logger.log("Please ensure this script is bound to a Google Sheet or you have the necessary permissions.");
      return; // Exit if sheet setup fails.
    }
    Logger.log(`Successfully configured Google Sheet "${TARGET_SHEET_NAME}".`);
    Logger.log('---');
  
  
    // --- Gmail Search Setup ---
    // Calculate the date range for the Gmail search query.
    const today = new Date();
    const dateBefore = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
    // Calculate the date N days ago for the 'after' part of the query.
    const dateAfterObj = new Date();
    dateAfterObj.setDate(today.getDate() - DATE_RANGE_DAYS);
    const dateAfter = Utilities.formatDate(dateAfterObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
    // Construct the Gmail search query.
    // It looks for emails with attachments within the specified date range.
    const searchQuery = `has:attachment after:${dateAfter} before:${dateBefore}`;
    Logger.log(`Gmail search query: "${searchQuery}"`);
  
    // Search Gmail for threads matching the query.
    // GmailApp.search returns threads in reverse chronological order (newest first).
    const threads = GmailApp.search(searchQuery);
  
    // Check if any threads were found.
    if (threads.length === 0) {
      Logger.log("No emails found with attachments in the specified date range.");
      return; // Exit if no emails are found.
    }
  
    let emailsProcessedCount = 0;
  
    // --- Process Emails and Extract PDFs ---
    // Iterate through each thread found.
    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      // Get messages from the current thread.
      // Messages within a thread are also in chronological order (oldest first).
      const messages = thread.getMessages();
  
      // Iterate through messages in reverse order (newest first within the thread)
      // to prioritize more recent emails.
      for (let j = messages.length - 1; j >= 0; j--) {
        // Check if we've reached the maximum number of emails to process.
        if (emailsProcessedCount >= MAX_EMAILS_TO_PROCESS) {
          Logger.log(`Reached maximum limit of ${MAX_EMAILS_TO_PROCESS} emails to process. Stopping.`);
          return; // Exit the function as limit is reached.
        }
  
        const message = messages[j];
        const attachments = message.getAttachments();
  
        let pdfBlob = null;
        let detectedPdfFilename = null;
  
        // Loop through each attachment of the current message to find the first PDF.
        for (let k = 0; k < attachments.length; k++) {
          const attachment = attachments[k];
  
          // Check if the attachment's content type is PDF.
          if (attachment.getContentType() === MimeType.PDF) {
            pdfBlob = attachment.getAs(MimeType.PDF);
            detectedPdfFilename = attachment.getName();
            break; // Found a PDF, stop searching attachments for this message.
          }
        }
  
        // If a PDF attachment was found in the current message.
        if (pdfBlob !== null) {
          Logger.log(`--- Processing Email ${emailsProcessedCount + 1} ---`);
          Logger.log(`Subject: ${message.getSubject()}`);
          Logger.log(`From: ${message.getFrom()}`);
          Logger.log(`Date: ${message.getDate()}`);
          Logger.log(`Found PDF: '${detectedPdfFilename}'. Attempting to extract text using Drive OCR.`);
  
          // Get the text from the PDF blob using the Google Drive OCR method.
          let extractedText = '';
          try {
            extractedText = extractTextFromPdfUsingDriveOcr(pdfBlob, detectedPdfFilename);
            Logger.log(`Successfully extracted text from '${detectedPdfFilename}'.`);
            // Log only the first 200 characters to prevent excessive log size.
            Logger.log(extractedText.substring(0, 200) + (extractedText.length > 200 ? '...' : ''));
  
            // --- Push to Google Sheet ---
            sheet.appendRow([
              message.getSubject(),
              message.getFrom(),
              Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
              detectedPdfFilename,
              extractedText
            ]);
            Logger.log(`Pushed data to "${TARGET_SHEET_NAME}".`);
  
          } catch (error) {
            Logger.log(`Error extracting text from '${detectedPdfFilename}' using Drive OCR: ${error.message}`);
            Logger.log("Please ensure 'Drive API' Advanced Service is enabled and permissions are granted.");
            extractedText = `ERROR: Could not extract text using Drive OCR (${error.message})`; // Store error message
          }
          emailsProcessedCount++; // Increment count only after attempting to process a PDF.
        } else {
          Logger.log(`No PDF attachment found in message (Subject: ${message.getSubject()}).`);
        }
      }
    }
  
    if (emailsProcessedCount === 0) {
      Logger.log("No PDF attachments were found and processed within the specified criteria.");
    } else {
      Logger.log(`--- Finished processing. Total PDFs extracted and pushed: ${emailsProcessedCount} ---`);
    }
  }
  
  /**
   * Extracts text from a PDF Blob using Google Drive's OCR functionality.
   * This function uploads the PDF to Drive, converts it to a Google Doc with OCR,
   * reads the text, and then deletes the temporary files.
   *
   * @param {GoogleAppsScript.Base.Blob} pdfBlob The PDF file as a Blob.
   * @param {string} originalFilename The original filename of the PDF.
   * @returns {string} The extracted text from the PDF.
   * @throws {Error} If OCR fails or temporary files cannot be deleted.
   */
  function extractTextFromPdfUsingDriveOcr(pdfBlob, originalFilename) {
    let tempPdfFileId = null;
    let tempDocFileId = null;
    let extractedText = '';
  
    try {
      // 1. Upload the PDF Blob to Google Drive.
      // Use a temporary name to avoid conflicts and make it identifiable.
      const tempPdfName = `temp_ocr_pdf_${Date.now()}_${originalFilename}`;
      const pdfFile = Drive.Files.insert({
        title: tempPdfName,
        mimeType: MimeType.PDF
      }, pdfBlob);
      tempPdfFileId = pdfFile.id;
      Logger.log(`Uploaded temporary PDF to Drive: ${tempPdfFileId}`);
  
      // 2. Convert the uploaded PDF to a Google Document with OCR enabled.
      // We use the ID of the newly uploaded PDF as the source.
      const tempDocName = `ocr_doc_from_${originalFilename}_${Date.now()}`;
      const docFile = Drive.Files.insert({
        title: tempDocName,
        mimeType: MimeType.GOOGLE_DOCS, // Target MIME type for Google Docs
        ocr: true,                     // Enable Optical Character Recognition
        ocrLanguage: 'en'              // Optional: Specify language for better OCR results (e.g., 'en', 'fr', 'es')
      }, pdfFile.getBlob()); // Pass the blob of the uploaded PDF for conversion
      tempDocFileId = docFile.id;
      Logger.log(`Converted PDF to temporary Google Doc with OCR: ${tempDocFileId}`);
  
      // 3. Open the Google Document and extract its text content.
      const doc = DocumentApp.openById(tempDocFileId);
      extractedText = doc.getBody().getText();
      Logger.log('Text extracted from Google Doc.');
  
    } catch (error) {
      Logger.log(`Error during OCR process: ${error.message}`);
      throw new Error(`Failed to extract text using Drive OCR: ${error.message}`);
    } finally {
      // 4. Clean up: Delete the temporary PDF and Google Doc files from Drive.
      // This ensures your Drive doesn't get cluttered with temporary files.
      if (tempDocFileId) {
        try {
          Drive.Files.remove(tempDocFileId);
          Logger.log(`Deleted temporary Google Doc: ${tempDocFileId}`);
        } catch (e) {
          Logger.log(`Warning: Could not delete temporary Google Doc ${tempDocFileId}: ${e.message}`);
        }
      }
      if (tempPdfFileId) {
        try {
          Drive.Files.remove(tempPdfFileId);
          Logger.log(`Deleted temporary PDF file: ${tempPdfFileId}`);
        } catch (e) {
          Logger.log(`Warning: Could not delete temporary PDF file ${tempPdfFileId}: ${e.message}`);
        }
      }
    }
    return extractedText;
  }
  
  /*
   * ==============================================================================
   * ADDITIONAL SETUP FOR GOOGLE SHEETS INTEGRATION:
   * ==============================================================================
   * This script is designed to run from a Google Apps Script project that is
   * BOUND to a Google Sheet (i.e., you opened the script editor from within a Sheet).
   * If you're running this from a standalone script, you'll need to specify
   * the target spreadsheet by ID, like:
   * `const spreadsheet = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID');`
   *
   * Ensure your script has permissions to access Google Sheets, Gmail, and Google Drive.
   *
   * ==============================================================================
   * IMPORTANT: ENABLE GOOGLE DRIVE ADVANCED SERVICE
   * ==============================================================================
   * 1. Open your Apps Script project.
   * 2. In the left sidebar, click on 'Services' (it looks like a '+' icon).
   * 3. Find and select 'Drive API'.
   * 4. Click 'Add'.
   * You will be prompted to grant permissions for Google Drive the first time you run this script.
   */
  