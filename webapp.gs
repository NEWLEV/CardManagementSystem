// --- FILE: WebApp.gs ---
// Contains all functions directly callable from the client-side JavaScript.

/**
 * ENHANCEMENT: Bundles initial app data into a single server call.
 * UPDATED: Includes user email.
 */
function getAppData() {
  try {
    const availableCards = getPublicInventoryList();
    const history = getHistory();
    const counts = getCardCounts();
    const userEmail = getCurrentUserEmail();

    return {
      availableCards: availableCards,
      history: history,
      counts: counts,
      userEmail: userEmail
    };
  } catch (e) {
    console.error(`getAppData failed: ${e.message}`, e.stack);
    _sendAdminAlert('getAppData CRITICAL Failure', `Error: ${e.message}\nStack: ${e.stack}`);
    return { error: `Server error during initial load: ${e.message}` };
  }
}

/**
 * Handles submission using batch processing. - UPDATED
 */
function addSignatureEntry(data) {
  console.log('Starting BATCH multi-card entry');
  const startTime = Date.now();
  try {
    // Added PRN check
    if (!data || !data.name || !data.prn || !data.date || !data.signature || !data.cards || data.cards.length === 0) {
      throw new Error('Missing required data (Name, PRN, Date, Signature, or Cards)');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.DATA);
    if (!sheet) {
        _sendAdminAlert('addSignatureEntry Failure', `Sheet "${SHEETS.DATA}" not found.`);
        throw new Error(`Sheet "${SHEETS.DATA}" not found`);
    }

    const currentUser = getCurrentUserEmail();
    const timestamp = new Date();

    const rowsToAdd = [];
    const auditEntriesToAdd = [];
    let cardProcessingErrors = [];

    const commonData = {
      date: data.date,
      name: data.name.toString().trim().toUpperCase(),
      prn: data.prn.toString().trim(), // Added PRN
      note: (data.note || '').toString().trim(),
      signatureBase64: data.signature,
      isDropOff: Boolean(data.isDropOff),
      isSafeHarbor: Boolean(data.isSafeHarbor)
    };

    let signaturePlaceholder = '';
    if (commonData.isDropOff) signaturePlaceholder = 'DROP-OFF';
    else if (commonData.isSafeHarbor) signaturePlaceholder = 'SAFE HARBOR';

    const usedCardsSet = _getUsedCardSet();

    for (const card of data.cards) {
      try {
        if (!card.type || !card.cardNumber || !card.reason) {
          throw new Error('Card entry is missing Type, Number, or Reason.');
        }
        if (!CARD_TYPES.includes(card.type)) {
          throw new Error(`Invalid card type: ${card.type}`);
        }

        const cleanCardNumber = _normalizeCardNumber(card.cardNumber);
        const cleanReason = (card.reason || '').toString().trim();
        const cleanApprovalCode = (card.approvalCode || '').toString().trim();
        const entryId = Utilities.getUuid();

        if (cleanReason === 'Other') {
            if (!cleanApprovalCode) {
                throw new Error(`Approval Code required for card ${cleanCardNumber} (Reason: Other).`);
            }
            if (!_isValidApprovalCode(cleanApprovalCode)) {
                throw new Error(`Invalid Approval Code for card ${cleanCardNumber}.`);
            }
        }

        const cardKey = `${card.type}|${cleanCardNumber}`;
        if (usedCardsSet.has(cardKey)) {
          throw new Error(`Card ${cleanCardNumber} (${card.type}) is no longer available.`);
        }

        // Updated array size and indices
        const newRowData = new Array(COLUMNS.DATA_ENTRY_ID).fill('');
        newRowData[COLUMNS.DATA_DATE - 1] = commonData.date;
        newRowData[COLUMNS.DATA_NAME - 1] = commonData.name;
        newRowData[COLUMNS.DATA_PRN - 1] = commonData.prn; // Save PRN
        newRowData[COLUMNS.DATA_TYPE - 1] = card.type.trim();
        newRowData[COLUMNS.DATA_NUMBER - 1] = cleanCardNumber;
        newRowData[COLUMNS.SIGNATURE - 1] = signaturePlaceholder;
        newRowData[COLUMNS.DATA_REASON - 1] = cleanReason;
        newRowData[COLUMNS.DATA_NOTE - 1] = commonData.note;
        newRowData[COLUMNS.DATA_TIMESTAMP - 1] = timestamp;
        newRowData[COLUMNS.BASE64 - 1] = (signaturePlaceholder === '') ? commonData.signatureBase64 : '';
        newRowData[COLUMNS.DATA_ENTRY_ID - 1] = entryId;

        rowsToAdd.push(newRowData);
        auditEntriesToAdd.push({ type: card.type, number: cleanCardNumber, user: currentUser, dropOff: commonData.isDropOff });

        usedCardsSet.add(cardKey);

        if (commonData.isDropOff) {
          runBackgroundOperations(card, commonData, currentUser, null);
        }

      } catch (cardError) {
        console.warn(`Card processing failed: ${cardError.message}`);
        cardProcessingErrors.push(`Card ${card.cardNumber || 'N/A'}: ${cardError.message}`);
      }
    }

    let newRowIndices = [];
    if (rowsToAdd.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      const numCols = COLUMNS.DATA_ENTRY_ID; // Use updated constant

      sheet.getRange(startRow, 1, rowsToAdd.length, numCols).setValues(rowsToAdd);
      sheet.setRowHeights(startRow, rowsToAdd.length, ROW_HEIGHT_CONFIG.STANDARD);

      newRowIndices = Array.from({length: rowsToAdd.length}, (_, i) => startRow + i);

      logCardUseBatch(auditEntriesToAdd);
    }

    SpreadsheetApp.flush();

    if (cardProcessingErrors.length > 0) {
      const message = `Submission processed with errors: ${cardProcessingErrors.join('; ')}`;
      if (rowsToAdd.length === 0) {
        throw new Error(message);
      } else {
        console.warn(message);
        return {
            status: "partial_success",
            message: `Successfully added ${rowsToAdd.length} card(s). Some cards failed.`,
            errors: cardProcessingErrors,
            rows: newRowIndices,
            processingTime: Date.now() - startTime
          };
      }
    }

    return {
      status: "success",
      message: `Successfully added ${rowsToAdd.length} card(s).`,
      rows: newRowIndices,
      processingTime: Date.now() - startTime
    };

  } catch (e) {
    console.error(`addSignatureEntry error: ${e.message}`, e.stack);
    _sendAdminAlert('addSignatureEntry CRITICAL Failure', `Error: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Server error: ${e.message}`);
  }
}

/**
 * Gets history entries from the Data sheet. Limits rows for performance. - UPDATED
 */
function getHistory() {
  const MAX_HISTORY_ROWS = 250;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.DATA);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - MAX_HISTORY_ROWS + 1);
    const numRows = lastRow - startRow + 1;
    if (numRows <= 0) return [];

    // Fetch columns up to the last defined column (Entry ID)
    const firstCol = COLUMNS.DATA_NAME; // Start reading from Name
    const lastCol = COLUMNS.DATA_ENTRY_ID; // Read up to Entry ID
    const numColsToFetch = lastCol - firstCol + 1;

    const range = sheet.getRange(startRow, firstCol, numRows, numColsToFetch);
    const data = range.getValues();

    const history = data.map(row => {
        // Calculate indices relative to the starting column (DATA_NAME)
        const nameIndex = COLUMNS.DATA_NAME - firstCol;
        const prnIndex = COLUMNS.DATA_PRN - firstCol; // Index for PRN
        const typeIndex = COLUMNS.DATA_TYPE - firstCol;
        const numberIndex = COLUMNS.DATA_NUMBER - firstCol;
        const statusIndex = COLUMNS.SIGNATURE - firstCol;
        const reasonIndex = COLUMNS.DATA_REASON - firstCol;
        const noteIndex = COLUMNS.DATA_NOTE - firstCol;
        const timestampIndex = COLUMNS.DATA_TIMESTAMP - firstCol;
        const entryIdIndex = COLUMNS.DATA_ENTRY_ID - firstCol;

        const timestampValue = row[timestampIndex];
        let timestampISO = null;
        if (timestampValue instanceof Date && !isNaN(timestampValue.getTime())) {
            timestampISO = timestampValue.toISOString();
        } else if (timestampValue) {
         try {
             const parsedDate = new Date(timestampValue);
             if (!isNaN(parsedDate.getTime())) {
                 timestampISO = parsedDate.toISOString();
             }
         } catch(e) { /* ignore parse error */ }
        }

        // Return object including PRN - Check if index is valid before accessing
        return {
          name: row[nameIndex] || 'N/A',
          prn: prnIndex >= 0 && prnIndex < row.length ? (row[prnIndex] || 'N/A') : 'N/A', // Read PRN
          type: typeIndex >= 0 && typeIndex < row.length ? (row[typeIndex] || 'N/A') : 'N/A',
          cardNumber: numberIndex >= 0 && numberIndex < row.length ? (row[numberIndex] || 'N/A') : 'N/A',
          reason: reasonIndex >= 0 && reasonIndex < row.length ? (row[reasonIndex] || 'N/A') : 'N/A',
          note: noteIndex >= 0 && noteIndex < row.length ? (row[noteIndex] || '') : '',
          timestamp: timestampISO,
          entryId: entryIdIndex >= 0 && entryIdIndex < row.length ? (row[entryIdIndex] || null) : null,
          signatureStatus: statusIndex >= 0 && statusIndex < row.length ? (row[statusIndex] || '') : ''
        };
    }).filter(item => item.entryId && item.timestamp); // Keep existing filter

    console.log(`Successfully fetched ${history.length} history items.`);
    return history;

  } catch (e) {
    console.error(`getHistory failed: ${e.message}`, e.stack);
    _sendAdminAlert('History Fetch Failed', `Error: ${e.message}`);
    return [];
  }
}


/**
 * Gets all data for a specific signature entry using its unique ID.
 * ADDED LOGGING FOR DEBUGGING "Entry not found".
 */
function getSignatureData(entryId) {
  console.log(`getSignatureData called for entryId: ${entryId}`); // Log input
  try {
    if (!entryId) {
      console.error("getSignatureData: No Entry ID provided.");
      throw new Error("No Entry ID provided.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(SHEETS.DATA);
    const archiveSheet = ss.getSheetByName(SHEETS.ARCHIVED_DATA); // Get archive sheet handle

    // --- Verify Column Index ---
    const idColumn = COLUMNS.DATA_ENTRY_ID;
    console.log(`Searching for ID in column: ${idColumn}`);
    if (!idColumn || idColumn < 1) {
        console.error(`Invalid DATA_ENTRY_ID column configured: ${idColumn}`);
        _sendAdminAlert("getSignatureData Config Error", `Invalid DATA_ENTRY_ID column: ${idColumn}`);
        throw new Error("Server configuration error (ID Column).");
    }
    // --- End Verify ---

    let foundCell = null;
    let searchSheetName = null; // Track where it was found

    // --- Search Active Sheet ---
    if (dataSheet && dataSheet.getLastRow() > 1) {
      console.log(`Searching active sheet "${SHEETS.DATA}"...`);
      const range = dataSheet.getRange(2, idColumn, dataSheet.getLastRow() - 1, 1);
      const textFinder = range.createTextFinder(entryId).matchEntireCell(true).matchCase(false); // Case-insensitive might be safer
      try {
           foundCell = textFinder.findNext();
           if (foundCell) {
                searchSheetName = SHEETS.DATA;
                console.log(`Entry ID ${entryId} found in active sheet at row ${foundCell.getRow()}.`);
           } else {
                console.log(`Entry ID ${entryId} not found in active sheet "${SHEETS.DATA}".`);
           }
      } catch (e) {
          console.error(`Error searching active sheet: ${e.message}`);
          // Decide if you want to stop or continue to archive search
          // For now, let's continue to archive search
      }
    } else {
        console.log(`Active data sheet "${SHEETS.DATA}" missing or empty, skipping search.`);
    }
    // --- End Search Active ---

    // --- Search Archive Sheet (if not found in active) ---
    if (!foundCell) {
        if (archiveSheet && archiveSheet.getLastRow() > 1) {
            console.log(`Searching archive sheet "${SHEETS.ARCHIVED_DATA}"...`);
            // Make sure archive sheet has enough columns
            if (archiveSheet.getMaxColumns() < idColumn) {
                 console.error(`Archive sheet "${SHEETS.ARCHIVED_DATA}" has fewer columns (${archiveSheet.getMaxColumns()}) than required ID column (${idColumn}). Cannot search.`);
                 _sendAdminAlert("getSignatureData Archive Error", `Archive sheet "${SHEETS.ARCHIVED_DATA}" columns (${archiveSheet.getMaxColumns()}) < ID column (${idColumn}).`);
            } else {
                const archiveRange = archiveSheet.getRange(2, idColumn, archiveSheet.getLastRow() - 1, 1);
                const archiveFinder = archiveRange.createTextFinder(entryId).matchEntireCell(true).matchCase(false);
                 try {
                    foundCell = archiveFinder.findNext();
                    if (foundCell) {
                        searchSheetName = SHEETS.ARCHIVED_DATA;
                        console.log(`Entry ID ${entryId} found in archive sheet at row ${foundCell.getRow()}.`);
                    } else {
                        console.log(`Entry ID ${entryId} not found in archive sheet "${SHEETS.ARCHIVED_DATA}".`);
                    }
                 } catch (e) {
                     console.error(`Error searching archive sheet: ${e.message}`);
                     // If archive search fails, the entry is truly not found
                 }
            }
        } else {
            console.log(`Archive data sheet "${SHEETS.ARCHIVED_DATA}" missing or empty, skipping search.`);
        }
    }
    // --- End Search Archive ---

    // --- Process Result ---
    if (!foundCell || !searchSheetName) {
      console.warn(`Entry ID ${entryId} was ultimately not found in any searched sheet.`);
      // Return null instead of throwing error, client expects null for not found
      return null;
      // throw new Error("Signature record not found."); // Old behavior
    }

    // Get the sheet where the ID was found
    const sourceSheet = ss.getSheetByName(searchSheetName);
    if (!sourceSheet) { // Should not happen if foundCell is valid, but good check
        console.error(`Critical error: Found cell for ${entryId} but source sheet "${searchSheetName}" is missing!`);
        throw new Error("Internal server error: Could not retrieve record sheet.");
    }

    const rowNum = foundCell.getRow();
    // Ensure we read enough columns, use max of configured ID column and actual sheet width
    const numColsToRead = Math.min(sourceSheet.getLastColumn(), Math.max(COLUMNS.DATA_ENTRY_ID, COLUMNS.BASE64));
    console.log(`Reading row ${rowNum} from sheet "${searchSheetName}", columns 1 to ${numColsToRead}.`);
    const rowData = sourceSheet.getRange(rowNum, 1, 1, numColsToRead).getValues()[0];

    // Map the row data to the object structure expected by the client
    const mappedData = mapRowToSignatureData(rowData);
    console.log(`Mapped data for entryId ${entryId}:`, mappedData); // Log the final mapped data
    return mappedData;

  } catch (e) {
    // Log unexpected errors during the process
    console.error(`getSignatureData unexpected error for ID ${entryId}: ${e.message}`, e.stack);
    // Send alert for unexpected errors
     _sendAdminAlert('getSignatureData CRITICAL Failure', `Entry ID: ${entryId}\nError: ${e.message}\nStack: ${e.stack}`);
    // Throw error to be caught by withFailureHandler on client
    throw new Error(`Server error retrieving record: ${e.message}`);
  }
}


/**
 * Generates a PDF for a specific entry and returns it as a Base64 string. - UPDATED
 */
function exportSignatureToPdf(entryId) {
  try {
    const data = getSignatureData(entryId); // This now includes PRN
    if (!data) {
        throw new Error(`Record with ID ${entryId} could not be found.`);
    }

    const timestamp = data.timestamp ? (data.timestamp instanceof Date ? data.timestamp.toLocaleString() : new Date(data.timestamp).toLocaleString()) : 'N/A';
    const entryDate = data.date ? (data.date instanceof Date ? data.date.toLocaleDateString() : new Date(data.date).toLocaleDateString()) : 'N/A';

    let signatureHtml = '';
    // ... (signature HTML generation logic remains the same) ...
     if (data.signatureStatus && data.signatureStatus !== '') {
      signatureHtml = `<p style="font-size: 16px; border: 1px solid #ccc; padding: 20px; text-align: center; background-color: #f9f9f9;"><strong>Status: ${escapeHtml(data.signatureStatus)}</strong><br>(No signature provided)</p>`;
    }
    else if (data.signatureBase64 && data.signatureBase64.length > 100) {
      signatureHtml = `<div style="border: 1px solid #000; padding: 5px; background-color: #f4f4f4; display: inline-block;">
        <img src="data:image/png;base64,${data.signatureBase64}" alt="Signature" style="width: 300px; height: auto; display: block;"/>
      </div>`;
    }
    else {
      signatureHtml = `<p style="font-size: 16px; border: 1px solid #ccc; padding: 20px; text-align: center; background-color: #f9f9f9;"><strong>(No Signature on File)</strong></p>`;
    }


    const htmlContent = `
      <html>
        <head>
          <style>
            body { font-family: 'Helvetica', 'Arial', sans-serif; font-size: 12px; margin: 25px; }
            h1 { font-size: 20px; color: #333; border-bottom: 2px solid #333; padding-bottom: 5px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #ddd; padding: 10px; text-align: left; vertical-align: top; word-break: break-word; }
            th { background-color: #f2f2f2; width: 30%; font-weight: bold; }
            .sig-container { margin-top: 25px; }
            .sig-label { font-size: 14px; font-weight: bold; margin-bottom: 10px; }
          </style>
        </head>
        <body>
          <h1>Card Distribution Record</h1>
          <p><strong>Entry ID:</strong> ${escapeHtml(data.entryId)}</p>
          <p><strong>Timestamp:</strong> ${escapeHtml(timestamp)}</p>
          <table>
            <tr><th>Full Name</th><td>${escapeHtml(data.name) || 'N/A'}</td></tr>
            <tr><th>PRN</th><td>${escapeHtml(data.prn) || 'N/A'}</td></tr> <!-- Added PRN -->
            <tr><th>Date of Entry</th><td>${escapeHtml(entryDate)}</td></tr>
            <tr><th>Card Type</th><td>${escapeHtml(data.cardType) || 'N/A'}</td></tr>
            <tr><th>Card Number</th><td>${escapeHtml(data.cardNumber) || 'N/A'}</td></tr>
            <tr><th>Reason</th><td>${escapeHtml(data.reason) || 'N/A'}</td></tr>
            <tr><th>Notes</th><td>${escapeHtml(data.note) || 'N/A'}</td></tr>
          </table>
          <div class="sig-container">
            <div class="sig-label">Signature / Status:</div>
            ${signatureHtml}
          </div>
        </body>
      </html>
    `;

    const htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, `signature_${data.entryId}.html`);
    const pdfBlob = htmlBlob.getAs(MimeType.PDF);

    return Utilities.base64Encode(pdfBlob.getBytes());

  } catch (e) {
    console.error(`exportSignatureToPdf error for ID ${entryId}: ${e.message}`);
    _sendAdminAlert('PDF Export Failed', `Failed to generate PDF for entry ID ${entryId}.\nError: ${e.message}`);
    throw new Error(`Server error generating PDF: ${e.message}`);
  }
}

/**
 * Client-callable function to validate an admin approval code.
 */
function checkAdminCode(submittedCode) {
  try {
    if (!submittedCode || String(submittedCode).trim() === '') {
      return false;
    }
    return _isValidApprovalCode(submittedCode);
  } catch (e) {
    console.error(`checkAdminCode error: ${e.message}`);
    return false;
  }
}

/**
 * Admin function to add new cards to the inventory sheet.
 */
function addCardsToInventory(type, numbersString) {
  if (!isUserAdmin()) {
    throw new Error('Permission denied. Admin access required.');
  }
  if (!type || !CARD_TYPES.includes(type)) {
    throw new Error('Invalid card type specified.');
  }
  if (!numbersString || numbersString.trim() === '') {
    throw new Error('No card numbers provided.');
  }
  
  try {
    const newCards = [...new Set(
        numbersString.split(/[\s,]+/)
        .map(_normalizeCardNumber)
        .filter(v => v !== '')
    )];
        
    if (newCards.length === 0) {
      throw new Error('No valid card numbers were parsed.');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cardSheet = ss.getSheetByName(SHEETS.CARDS);
    if (!cardSheet) {
      throw new Error(`Sheet "${SHEETS.CARDS}" not found.`);
    }
    
    const typeMap = CARD_TYPES.reduce((map, t, i) => { map[t] = i + 1; return map; }, {});
    const col = typeMap[type];
    
    const lastDataRow = cardSheet.getLastRow();
    let existingCardSet = new Set();
    if (lastDataRow > 1) {
      const existingData = cardSheet.getRange(2, col, lastDataRow - 1, 1).getValues();
      existingData.forEach(row => {
          const card = _normalizeCardNumber(row[0]);
          if (card) existingCardSet.add(card);
      });
    }
    
    const validNewCards = newCards.filter(card => !existingCardSet.has(card));
    const duplicateCount = newCards.length - validNewCards.length;
    
    if (validNewCards.length === 0) {
      return { success: true, added: 0, duplicates: duplicateCount, message: 'All provided cards already exist in the inventory.' };
    }
    
    const colValues = cardSheet.getRange(1, col, cardSheet.getMaxRows()).getValues();
    let nextEmptyRowInCol = 1;
    while (nextEmptyRowInCol < colValues.length && _normalizeCardNumber(colValues[nextEmptyRowInCol][0]) !== '') {
      nextEmptyRowInCol++;
    }
    const appendRowIndex = nextEmptyRowInCol + 1;

    const dataToAppend = validNewCards.map(card => [card]);
    
    cardSheet.getRange(appendRowIndex, col, dataToAppend.length, 1).setValues(dataToAppend);
    
    console.log(`Admin ${getCurrentUserEmail()} added ${validNewCards.length} new '${type}' cards.`);
    
    CacheService.getScriptCache().remove(CACHE_KEYS.MASTER_INVENTORY);
    console.log("MASTER_CARD_INVENTORY_CACHE invalidated after adding cards.");
    
    return { success: true, added: validNewCards.length, duplicates: duplicateCount };
    
  } catch (error) {
    console.error(`addCardsToInventory failed: ${error.message}`);
    _sendAdminAlert('Admin Card Add Failure', `Error adding cards for ${type}.\n\nError: ${error.message}`);
    throw error;
  }
}

/**
 * Admin function to remove cards from the inventory sheet.
 */
function removeCardsFromInventory(type, numbersString) {
  if (!isUserAdmin()) {
    throw new Error('Permission denied. Admin access required.');
  }
  if (!type || !CARD_TYPES.includes(type)) {
    throw new Error('Invalid card type specified.');
  }
  if (!numbersString || numbersString.trim() === '') {
    throw new Error('No card numbers to remove provided.');
  }
  
  try {
    const cardsToRemoveSet = new Set(
      numbersString.split(/[\s,]+/)
        .map(_normalizeCardNumber)
        .filter(v => v !== '')
    );
        
    if (cardsToRemoveSet.size === 0) {
      throw new Error('No valid card numbers were parsed for removal.');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cardSheet = ss.getSheetByName(SHEETS.CARDS);
    if (!cardSheet) {
      throw new Error(`Sheet "${SHEETS.CARDS}" not found.`);
    }
    
    const typeMap = CARD_TYPES.reduce((map, t, i) => { map[t] = i + 1; return map; }, {});
    const col = typeMap[type];
    
    const lastDataRow = cardSheet.getLastRow();
    if (lastDataRow < 2) {
      return { success: true, removed: 0, notFound: cardsToRemoveSet.size, message: 'Inventory for this type is empty.' };
    }
    
    const cardRange = cardSheet.getRange(2, col, lastDataRow - 1, 1);
    const allCardsData = cardRange.getValues();
    
    let removedCount = 0;
    const cardsToKeep = [];
    const rowsToDelete = [];

    allCardsData.forEach((row, index) => {
        const card = _normalizeCardNumber(row[0]);
        if (card === '') return;

        if (cardsToRemoveSet.has(card)) {
            removedCount++;
            cardsToRemoveSet.delete(card);
            rowsToDelete.push(index + 2);
        } else {
            cardsToKeep.push([card]);
        }
    });
    
    const notFoundCount = cardsToRemoveSet.size;
    
    if (removedCount === 0) {
      return { success: true, removed: 0, notFound: notFoundCount, message: 'None of the provided cards were found in the inventory.' };
    }
    
    cardSheet.getRange(2, col, cardSheet.getMaxRows() - 1, 1).clearContent();
    
    if (cardsToKeep.length > 0) {
      cardSheet.getRange(2, col, cardsToKeep.length, 1).setValues(cardsToKeep);
    }
    
    console.log(`Admin ${getCurrentUserEmail()} removed ${removedCount} '${type}' cards.`);
    
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEYS.USED_CARDS);
    cache.remove(CACHE_KEYS.MASTER_INVENTORY);
    console.log("Both caches (USED_CARDS_SET, MASTER_CARD_INVENTORY_CACHE) invalidated after removing cards.");
    
    return { success: true, removed: removedCount, notFound: notFoundCount };
    
  } catch (error) {
    console.error(`removeCardsFromInventory failed: ${error.message}`);
    _sendAdminAlert('Admin Card Remove Failure', `Error removing cards for ${type}.\n\nError: ${error.message}`);
    throw error;
  }
}

/**
 * ENHANCED: Gets available card numbers using server-side caches.
 * This is called by the Admin Panel's "List Inventory" button.
 */
function getAvailableCardNumbers(type) {
  // No admin check here, as it's called by the client.
  // The *visibility* of the button is controlled client-side,
  // and this data is already public via getPublicInventoryList.
  try {
    if (!CARD_TYPES.includes(type)) {
        console.warn(`getAvailableCardNumbers called with invalid type: ${type}`);
        return [];
    }
    
    const masterInventoryMap = _getAllCardsFromInventory();
    const usedCardsSet = _getUsedCardSet();

    const allCardsForType = masterInventoryMap.get(type);
    if (!allCardsForType || allCardsForType.size === 0) {
      return [];
    }
    
    const availableCards = [];
    for (const card of allCardsForType) {
      if (!usedCardsSet.has(`${type}|${card}`)) {
        availableCards.push(card);
      }
    }
    
    return availableCards.sort();
    
  } catch (error) {
    console.error(`getAvailableCardNumbers failed for type ${type}: ${error.message}`);
    _sendAdminAlert('Available Card Check Failed', `Error getting cards for type ${type}: ${error.message}`);
    return [];
  }
}
