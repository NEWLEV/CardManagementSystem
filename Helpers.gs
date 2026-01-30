// --- FILE: Helpers.gs ---
// Contains internal helper and utility functions used across the application.

/**
 * Helper function to send email alerts to admins.
 */
function _sendAdminAlert(subject, body) {
  if (!ADMIN_EMAILS || ADMIN_EMAILS.length === 0 || ADMIN_EMAILS[0] === '') {
      console.error(`Admin emails not configured. Cannot send alert: ${subject}`);
      return;
  }
  try {
    const emailQuota = MailApp.getRemainingDailyQuota();
    if (emailQuota > 0) {
      MailApp.sendEmail({
        to: ADMIN_EMAILS.join(','),
        subject: `Card Management App Alert: ${subject}`,
        body: `A critical error or warning was detected:\n\n${body}\n\nTimestamp: ${new Date().toISOString()}`
      });
      console.log(`Admin alert sent: ${subject}`);
    } else {
      console.warn(`Could not send admin alert. Email quota exceeded. Subject: ${subject}`);
    }
  } catch (e) {
    console.error(`Failed to send admin alert: ${e.message}`);
  }
}

/**
 * Normalization function for card numbers.
 */
function _normalizeCardNumber(cardNum) {
  if (cardNum === null || cardNum === undefined) {
    return '';
  }
  return String(cardNum).trim();
}

/**
 * Function to check if the provided approval code is valid.
 */
function _isValidApprovalCode(submittedCode) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const codeSheet = ss.getSheetByName(SHEETS.APPROVAL_CODES);
        
        if (!codeSheet) {
            console.error(`ApprovalCodes sheet ("${SHEETS.APPROVAL_CODES}") not found!`);
            _sendAdminAlert('Approval Code Check Failed', `Sheet "${SHEETS.APPROVAL_CODES}" not found.`);
            return false;
        }
        
        const lastRow = codeSheet.getLastRow();
        if (lastRow === 0) {
            console.warn(`No approval codes found in "${SHEETS.APPROVAL_CODES}" sheet.`);
             _sendAdminAlert('Approval Code Check Warning', `No codes found in "${SHEETS.APPROVAL_CODES}". Submissions using "Other" will fail.`);
            return false;
        }

        const validCodes = codeSheet.getRange(1, 1, lastRow, 1)
                                    .getValues()
                                    .map(row => String(row[0]).trim())
                                    .filter(code => code.length > 0);

        if (validCodes.length === 0) {
            console.warn(`No valid approval codes found in column A of "${SHEETS.APPROVAL_CODES}" sheet.`);
            _sendAdminAlert('Approval Code Check Warning', `No codes found in column A of "${SHEETS.APPROVAL_CODES}". Submissions using "Other" will fail.`);
            return false;
        }

        const submittedCodeClean = String(submittedCode).trim();
        return validCodes.includes(submittedCodeClean);

    } catch (error) {
        console.error(`Error validating approval code: ${error.message}`);
        _sendAdminAlert('Approval Code Validation Error', `An error occurred while checking the code.\n\nError: ${error.message}`);
        return false;
    }
}

/**
 * NEW: Gets the current user's email address.
 */
function getCurrentUserEmail() {
    try {
        let email = Session.getActiveUser().getEmail();
        if (!email) {
            email = Session.getEffectiveUser().getEmail();
        }
        return email || 'Unknown User';
    } catch (e) {
        console.error("Error getting user email:", e);
        return 'Error User';
    }
}

/**
 * Batch logs card usage and invalidates cache once.
 */
function logCardUseBatch(entries) {
  if (!entries || entries.length === 0) return;
  console.log(`Batch logging ${entries.length} card uses.`);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.AUDIT_LOG);
      sheet.appendRow(['Timestamp', 'Type', 'Card Number', 'Used By', 'Mode']);
      console.log(`Created audit log sheet: ${SHEETS.AUDIT_LOG}`);
    }
    
    const timestamp = new Date();
    const rowsToAppend = entries.map(entry => [
      timestamp,
      entry.type.toString().trim(),
      _normalizeCardNumber(entry.number),
      entry.user || 'Unknown',
      entry.dropOff ? 'Drop-Off' : ''
    ]);
    
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    
    CacheService.getScriptCache().remove(CACHE_KEYS.USED_CARDS);
    console.log("USED_CARDS_SET Cache invalidated after batch audit log write.");
    
  } catch (error) {
    console.error('logCardUseBatch failed:', error.message);
     _sendAdminAlert('Batch Audit Log Failed', `Could not log card uses.\n\nError: ${error.message}`);
  }
}

/**
 * Background operations (external export for Drop-Offs).
 */
function runBackgroundOperations(cardData, commonData, currentUser, rowIndex) {
  if (commonData.isDropOff) {
    try {
      const externalSpreadsheet = SpreadsheetApp.openById(EXPORT_SHEET_ID);
      const externalSheet = externalSpreadsheet.getSheets()[0];
      if (!externalSheet) {
          throw new Error(`First sheet not found in external spreadsheet ID: ${EXPORT_SHEET_ID}`);
      }
      externalSheet.appendRow([
        commonData.date,
        commonData.name,
        cardData.type,
        cardData.cardNumber,
        '',
        commonData.note
      ]);
      console.log(`Exported drop-off for ${commonData.name} to external sheet.`);
    } catch (exportError) {
      console.error('Export to external sheet failed:', exportError.message);
      _sendAdminAlert('External Export Failure', `Failed to export drop-off data.\nSpreadsheet ID: ${EXPORT_SHEET_ID}\nError: ${exportError.message}`);
    }
  }
}

/**
 * Maps a raw spreadsheet row array to a structured object for the client.
 * Includes fixes for broken signature images and column shifting.
 */
function mapRowToSignatureData(row) {
  if (!row || row.length === 0) return null;

  // Helper to safely get data at a specific column index (1-based config -> 0-based array)
  const getCol = (colIndex) => {
    const idx = colIndex - 1;
    return (idx >= 0 && idx < row.length) ? row[idx] : '';
  };

  // 1. Safe Timestamp Parsing
  let timestampISO = null;
  const rawTs = getCol(COLUMNS.DATA_TIMESTAMP);
  if (rawTs instanceof Date) {
    timestampISO = rawTs.toISOString();
  } else if (rawTs) {
    try { timestampISO = new Date(rawTs).toISOString(); } catch (e) {}
  }

  // 2. SIGNATURE IMAGE FIX
  // The app saves signatures with the "data:image/png;base64," prefix.
  // We must strip this prefix here because the Frontend and PDF generator 
  // both add it back manually. If we don't, we get double prefixes (broken image).
  let cleanSignature = '';
  const rawSig = getCol(COLUMNS.BASE64);
  
  if (rawSig && typeof rawSig === 'string') {
    if (rawSig.startsWith('data:image')) {
      // Split at the comma and take the second part (the raw base64)
      const parts = rawSig.split(',');
      cleanSignature = parts.length > 1 ? parts[1] : parts[0]; 
    } else {
      cleanSignature = rawSig;
    }
  }

  return {
    date: getCol(COLUMNS.DATA_DATE),
    name: getCol(COLUMNS.DATA_NAME),
    prn: getCol(COLUMNS.DATA_PRN),
    cardType: getCol(COLUMNS.DATA_TYPE),
    cardNumber: getCol(COLUMNS.DATA_NUMBER),
    signatureStatus: getCol(COLUMNS.SIGNATURE), // e.g., "DROP-OFF"
    reason: getCol(COLUMNS.DATA_REASON),
    note: getCol(COLUMNS.DATA_NOTE),
    timestamp: timestampISO,
    signatureBase64: cleanSignature, // Now perfectly clean for PDF/Web
    entryId: getCol(COLUMNS.DATA_ENTRY_ID)
  };
}

/**
 * Normalizes card numbers (removes spaces, uppercase) for consistent comparison.
 */
function _normalizeCardNumber(raw) {
  if (!raw) return '';
  return String(raw).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Validates approval codes (Basic check: not empty, min length).
 */
function _isValidApprovalCode(code) {
  if (!code) return false;
  const str = String(code).trim();
  return str.length > 2; 
}

/**
 * Sends email alerts to admins (Utility wrapper).
 */
function _sendAdminAlert(subject, body) {
  if (typeof ADMIN_EMAILS === 'undefined' || !ADMIN_EMAILS || ADMIN_EMAILS.length === 0) return;
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAILS.join(','),
      subject: `[CardSystem Alert] ${subject}`,
      body: body
    });
  } catch (e) {
    console.error("Failed to send admin alert:", e);
  }
}

 
/** Simple HTML escaping helper */
function escapeHtml(unsafe) {
   if (!unsafe) return '';
   return unsafe.toString()
         .replace(/&/g, "&amp;")
         .replace(/</g, "&lt;")
         .replace(/>/g, "&gt;")
         .replace(/"/g, "&quot;")
         .replace(/'/g, "&#039;");
}
