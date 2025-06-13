/**
 * Gmail operations for the shipping system
 */

/**
 * Searches Gmail for threads matching the configured criteria
 * @returns {GmailThread[]} Array of matching Gmail threads
 */
function searchGmailThreads() {
  const { GMAIL } = CONFIG;
  let searchQuery = GMAIL.SEARCH_QUERY;

  // Add exclusion for processed label
  if (GMAIL.PROCESSED_LABEL) {
    searchQuery += ` -label:${GMAIL.PROCESSED_LABEL}`;
  }

  // Add date range
  if (GMAIL.DATE_RANGE_DAYS && GMAIL.DATE_RANGE_DAYS > 0) {
    searchQuery += ` newer_than:${GMAIL.DATE_RANGE_DAYS}d`;
  }

  Logger.log(`Executing Gmail search with query: "${searchQuery}"`);
  return GmailApp.search(searchQuery, 0, GMAIL.MAX_EMAILS_TO_PROCESS);
}

/**
 * Processes a Gmail thread and extracts attachments
 * @param {GmailThread} thread - The Gmail thread to process
 * @returns {Array} Array of processed attachments with metadata
 */
function processGmailThread(thread) {
  const processedAttachments = [];
  const messages = thread.getMessages();

  for (const message of messages) {
    const attachments = message.getAttachments();
    const messageData = {
      sender: message.getFrom(),
      subject: message.getSubject(),
      date: message.getDate()
    };

    for (const attachment of attachments) {
      processedAttachments.push({
        ...messageData,
        attachment,
        filename: attachment.getName()
      });
    }
  }

  return processedAttachments;
}

/**
 * Marks a Gmail thread as processed
 * @param {GmailThread} thread - The thread to mark
 */
function markThreadAsProcessed(thread) {
  const label = UTILS.getOrCreateLabel(CONFIG.GMAIL.PROCESSED_LABEL);
  thread.addLabel(label);
}

/**
 * Removes processed label from threads
 * @param {number} [maxThreads] - Maximum number of threads to process
 */
function cleanupProcessedLabels(maxThreads) {
  try {
    const label = GmailApp.getUserLabelByName(CONFIG.GMAIL.PROCESSED_LABEL);
    if (!label) return;

    const threads = label.getThreads();
    const threadsToProcess = maxThreads ? threads.slice(0, maxThreads) : threads;

    threadsToProcess.forEach(thread => thread.removeLabel(label));
    Logger.log(`Removed processed label from ${threadsToProcess.length} threads`);
  } catch (error) {
    Logger.error('Error in cleanupProcessedLabels:', error);
  }
}

/**
 * Extracts text from a PDF attachment using Google Drive OCR
 * @param {GmailAttachment} attachment - The PDF attachment
 * @returns {string} Extracted text from the PDF
 */
function extractTextFromPdf(attachment) {
  let tempPdfFileId = null;
  let tempDocFileId = null;

  try {
    // Upload PDF to Drive
    const tempPdfName = `temp_ocr_pdf_${Date.now()}_${attachment.getName()}`;
    const pdfFile = Drive.Files.insert({
      title: tempPdfName,
      mimeType: MimeType.PDF
    }, attachment.getAs(MimeType.PDF));
    tempPdfFileId = pdfFile.id;

    // Convert to Google Doc with OCR
    const tempDocName = `ocr_doc_from_${attachment.getName()}_${Date.now()}`;
    const docFile = Drive.Files.insert({
      title: tempDocName,
      mimeType: MimeType.GOOGLE_DOCS,
      ocr: true,
      ocrLanguage: CONFIG.PDF.OCR_LANGUAGE
    }, pdfFile.getBlob());
    tempDocFileId = docFile.id;

    // Extract text
    const doc = DocumentApp.openById(tempDocFileId);
    const text = doc.getBody().getText();

    return text;
  } finally {
    // Cleanup temporary files
    if (tempPdfFileId) Drive.Files.remove(tempPdfFileId);
    if (tempDocFileId) Drive.Files.remove(tempDocFileId);
  }
}

// Export Gmail operations
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    searchGmailThreads,
    processGmailThread,
    markThreadAsProcessed,
    cleanupProcessedLabels,
    extractTextFromPdf
  };
} else {
  // For Google Apps Script environment
  global.GMAIL_OPS = {
    searchGmailThreads,
    processGmailThread,
    markThreadAsProcessed,
    cleanupProcessedLabels,
    extractTextFromPdf
  };
} 