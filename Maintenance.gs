// --- FILE: Maintenance.gs ---
// Contains functions for diagnostics, scheduled maintenance, and archiving.

/**
 * Setup maintenance triggers via Admin menu
 */
function setupMaintenanceTriggers() {
  try {
    console.log("Setting up maintenance triggers...");

    // Delete existing triggers for these functions to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      const handlerFunction = trigger.getHandlerFunction();
      if (['scheduledQuickDiagnosis', 'scheduledDeepMaintenance', 'archiveOldData'].includes(handlerFunction)) {
        ScriptApp.deleteTrigger(trigger);
        console.log(`Deleted existing trigger for: ${handlerFunction}`);
      }
    });

    // Create new triggers
    // Daily quick diagnosis at 2 AM
    ScriptApp.newTrigger('scheduledQuickDiagnosis')
      .timeBased()
      .atHour(2)
      .everyDays(1)
      .create();
    console.log("Created trigger for scheduledQuickDiagnosis.");

    // Weekly deep maintenance on Sundays at 3 AM
    ScriptApp.newTrigger('scheduledDeepMaintenance')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(3)
      .create();
    console.log("Created trigger for scheduledDeepMaintenance.");

    // Monthly archiving on the 1st at 1 AM
    ScriptApp.newTrigger('archiveOldData')
      .timeBased()
      .onMonthDay(1)
      .atHour(1)
      .create();
    console.log("Created trigger for archiveOldData.");

    console.log("Maintenance triggers setup complete.");
    SpreadsheetApp.getUi().alert("Maintenance triggers have been set up successfully!");
    return "Triggers setup complete";

  } catch (error) {
    console.error("Error setting up triggers:", error.message, error.stack);
    _sendAdminAlert('Trigger Setup Failed', error.message);
    SpreadsheetApp.getUi().alert(`Error setting up triggers: ${error.message}`);
    throw error;
  }
}

/**
 * Quick diagnostic function to check basic health.
 */
function quickDiagnosis() {
  try {
    console.log("Running quick diagnosis...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = [SHEETS.DATA, SHEETS.CARDS, SHEETS.AUDIT_LOG, SHEETS.APPROVAL_CODES];
    const missingSheets = [];

    requiredSheets.forEach(sheetName => {
      if (!ss.getSheetByName(sheetName)) {
        missingSheets.push(sheetName);
      }
    });

    if (missingSheets.length > 0) {
      const message = `Missing required sheets: ${missingSheets.join(', ')}`;
      console.error(message);
      _sendAdminAlert('Quick Diagnosis Failed', message);
      return message;
    }

    console.log("Quick diagnosis passed: All required sheets present.");
    return "Quick diagnosis passed: All required sheets present.";
  } catch (error) {
    console.error("Quick diagnosis error:", error.message, error.stack);
    _sendAdminAlert('Quick Diagnosis Error', error.message);
    return `Error: ${error.message}`;
  }
}

/** Scheduled quick health check */
function scheduledQuickDiagnosis() {
  try {
    console.log("Running scheduled quick diagnosis...");
    const result = quickDiagnosis();
    console.log(`Scheduled diagnosis result: ${result}`);
  } catch (error) {
    console.error("Scheduled quick diagnosis failed:", error.message, error.stack);
    _sendAdminAlert('Scheduled Diagnosis Failed', error.message);
  }
}

/**
 * Scheduled deep maintenance.
 */
function scheduledDeepMaintenance() {
  try {
    console.log('Running scheduled deep maintenance.');
    // Currently no deep maintenance tasks are defined besides archiving,
    // which runs on its own schedule. Audit log trimming was previously considered
    // but is complex and potentially risky if done automatically without careful review.
    console.log('Deep maintenance: No specific tasks currently configured.');
    console.log('Deep maintenance completed.');
  } catch (error) {
    console.error('Scheduled deep maintenance failed:', error.message, error.stack);
    _sendAdminAlert('Deep Maintenance Failed', `Error: ${error.message}`);
  }
}

// --- ARCHIVING FUNCTIONS ---

/**
 * V4 FIX + Logging: Archives old data from BOTH Data and Card Audit Log sheets
 * and explicitly clears the used cards cache.
 */
function archiveOldData() {
  const archiveDateThreshold = new Date();
  archiveDateThreshold.setDate(archiveDateThreshold.getDate() - ARCHIVE_CONFIG.ARCHIVE_DAYS_AGO);
  let results = [];
  const overallStartTime = Date.now();

  console.log(`Starting archiving. Threshold Date: ${archiveDateThreshold.toISOString()}, Max Rows: ${ARCHIVE_CONFIG.MAX_ACTIVE_ROWS}, Trim Rows: ${ARCHIVE_CONFIG.ROWS_TO_TRIM_ON_HIGH_WATER}`);

  try {
    // Archive Data sheet
    console.log(`Starting archive process for sheet: ${SHEETS.DATA}`);
    const dataStartTime = Date.now();
    const dataResult = _archiveSheet_Optimized(
        SHEETS.DATA,
        SHEETS.ARCHIVED_DATA,
        COLUMNS.DATA_DATE - 1, // Date is in the first column (index 0 for array)
        archiveDateThreshold,
        ARCHIVE_CONFIG.MAX_ACTIVE_ROWS,
        ARCHIVE_CONFIG.ROWS_TO_TRIM_ON_HIGH_WATER
    );
    results.push(dataResult);
    console.log(`Finished archiving ${SHEETS.DATA}. Time: ${Date.now() - dataStartTime} ms. Result: ${dataResult}`);

    // Archive Card Audit Log sheet
    console.log(`Starting archive process for sheet: ${SHEETS.AUDIT_LOG}`);
    const auditStartTime = Date.now();
    const auditLogResult = _archiveSheet_Optimized(
        SHEETS.AUDIT_LOG,
        SHEETS.ARCHIVED_AUDIT_LOG,
        0, // Timestamp is the first column (index 0) in Audit Log
        archiveDateThreshold,
        ARCHIVE_CONFIG.MAX_ACTIVE_ROWS,
        ARCHIVE_CONFIG.ROWS_TO_TRIM_ON_HIGH_WATER
    );
    results.push(auditLogResult);
     console.log(`Finished archiving ${SHEETS.AUDIT_LOG}. Time: ${Date.now() - auditStartTime} ms. Result: ${auditLogResult}`);

    // ---> ADDED LOGGING AROUND CACHE CLEAR <---
    console.log(`Attempting to clear cache key: ${CACHE_KEYS.USED_CARDS}`);
    try {
        CacheService.getScriptCache().remove(CACHE_KEYS.USED_CARDS);
        console.log(`Cache key ${CACHE_KEYS.USED_CARDS} removed successfully.`);
    } catch (cacheError) {
        console.error(`Error removing cache key ${CACHE_KEYS.USED_CARDS}: ${cacheError.message}`, cacheError.stack);
         _sendAdminAlert('Archiving Cache Clear Failed', `Failed to remove USED_CARDS cache after archiving.\nError: ${cacheError.message}`);
         results.push("Error clearing used cards cache."); // Add error to results
    }
    // ---> END LOGGING <---

    const totalTime = Date.now() - overallStartTime;
    console.log(`Archiving complete. Total Time: ${totalTime} ms. Results: ${results.join(' | ')}`);
    return `Archiving complete. ${results.join(' | ')}`;

  } catch (error) {
    console.error('CRITICAL Error during data archiving:', error.message, error.stack);
    _sendAdminAlert('Archiving Failure', `The automated archiving process failed critically.\n\nError: ${error.message}\nStack: ${error.stack}`);
    // Re-throw the error so Apps Script logs it as a failure
    throw new Error(`Archiving failed: ${error.message}`);
  }
}

/**
 * OPTIMIZED Helper for archiving a sheet. Includes more logging.
 */
function _archiveSheet_Optimized(sourceSheetName, archiveSheetName, dateColumnIndex, archiveDate, maxRows, rowsToTrim) {
    console.log(`Optimized Archiving - Source: ${sourceSheetName}, Archive: ${archiveSheetName}, Date Col Index: ${dateColumnIndex}, Max Rows: ${maxRows}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(sourceSheetName);

    if (!sourceSheet) {
      console.warn(`Source sheet "${sourceSheetName}" not found. Skipping archive.`);
        return `${sourceSheetName}: Source sheet not found.`;
    }
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) { // Need at least 1 header row + 1 data row
        console.log(`${sourceSheetName}: No data rows (lastRow=${lastRow}). Skipping archive.`);
        return `${sourceSheetName}: No data to archive.`;
    }

    let archiveSheet = ss.getSheetByName(archiveSheetName);
    if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveSheetName);
        const sourceHeaderRange = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
        sourceHeaderRange.copyTo(archiveSheet.getRange(1, 1));
        archiveSheet.setFrozenRows(1);
        // Attempt to copy column widths
        try {
            for(let i = 1; i <= sourceSheet.getLastColumn(); i++) {
                archiveSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
            }
        } catch(e) { console.warn(`Could not copy column widths for ${archiveSheetName}: ${e.message}`); }
        console.log(`Created archive sheet: ${archiveSheetName}`);
    } else {
        // Ensure archive sheet has enough columns if source was modified
        if (archiveSheet.getMaxColumns() < sourceSheet.getMaxColumns()) {
            archiveSheet.insertColumnsAfter(archiveSheet.getMaxColumns(), sourceSheet.getMaxColumns() - archiveSheet.getMaxColumns());
            // Re-copy header if columns were added
            const sourceHeaderRange = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
            sourceHeaderRange.copyTo(archiveSheet.getRange(1, 1));
             console.log(`Added columns to ${archiveSheetName} to match ${sourceSheetName}.`);
        }
    }


    const numColumns = sourceSheet.getLastColumn();
    // Check if there's actually data to read
    const dataRange = sourceSheet.getRange(2, 1, lastRow - 1, numColumns);
    console.log(`Reading data from ${sourceSheetName} (Range: ${dataRange.getA1Notation()})...`);
    const allData = dataRange.getValues();
    console.log(`Finished reading ${allData.length} data rows from ${sourceSheetName}.`);

    const rowsToArchiveData = [];
    const sourceRowIndicesToDelete = new Set(); // Using Set for efficient check later

    // Identify rows older than the threshold date
    allData.forEach((row, index) => {
        let rowDate;
        // Check if the date column index is valid for the row
        if (dateColumnIndex >= 0 && dateColumnIndex < row.length) {
            const dateValue = row[dateColumnIndex];
            if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
                rowDate = dateValue;
            } else if (dateValue) { // Try parsing if it's not already a Date object
                try {
                    const parsedDate = new Date(dateValue);
                    if (!isNaN(parsedDate.getTime())) rowDate = parsedDate;
                } catch (e) { /* Ignore parsing errors, treat as invalid date */ }
            }
        } else {
            // Log if the date column index is out of bounds for this row
            // This might happen with inconsistent data entry or sheet structure changes
             console.warn(`Date column index ${dateColumnIndex} is out of bounds for row ${index + 2} in ${sourceSheetName}. Row length: ${row.length}. Skipping date check for this row.`);
        }


        if (rowDate && rowDate <= archiveDate) {
            // Ensure the row has the correct number of columns before adding
            let rowToArchive = row.slice(0, numColumns); // Take only expected columns
            while (rowToArchive.length < numColumns) rowToArchive.push(''); // Pad if necessary
            rowsToArchiveData.push(rowToArchive);
            sourceRowIndicesToDelete.add(index + 2); // Add the actual row number (1-based)
        }
    });
    console.log(`Found ${rowsToArchiveData.length} rows older than ${archiveDate.toLocaleDateString()} based on column index ${dateColumnIndex}.`);

    let highWaterRowsArchived = 0;
    const currentDataRowCount = lastRow - 1; // Number of data rows currently in the source sheet

    // Identify oldest rows if maxRows limit is exceeded
    if (maxRows > 0 && currentDataRowCount > maxRows) {
        const rowsOverLimit = currentDataRowCount - maxRows;
        // Determine how many *additional* rows need archiving to meet the target
        const neededTrimCount = Math.min(rowsOverLimit, rowsToTrim > 0 ? rowsToTrim : rowsOverLimit);
        let actualTrimCount = 0; // Count how many we actually mark

        console.log(`${sourceSheetName}: ${currentDataRowCount} rows exceeds max ${maxRows}. Target trim count: ${neededTrimCount}.`);

        // Iterate through *all* data rows *from oldest to newest* (index 0 upwards)
        for (let i = 0; i < allData.length && actualTrimCount < neededTrimCount; i++) {
            const rowIndexInSheet = i + 2; // Actual row number (1-based)
            // Check if this row is *not already* marked for deletion by the date threshold
            if (!sourceRowIndicesToDelete.has(rowIndexInSheet)) {
                let rowToArchive = allData[i].slice(0, numColumns);
                while (rowToArchive.length < numColumns) rowToArchive.push('');
                rowsToArchiveData.push(rowToArchive);
                sourceRowIndicesToDelete.add(rowIndexInSheet); // Mark this oldest row for deletion too
                actualTrimCount++;
                highWaterRowsArchived++; // Increment counter for high-water specifically
            }
        }
        console.log(`Marked ${highWaterRowsArchived} additional oldest rows for high-water mark archiving.`);
    }

    if (sourceRowIndicesToDelete.size === 0) { // Check the Set size
        console.log(`${sourceSheetName}: No rows met archive criteria (date or high-water).`);
        return `${sourceSheetName}: No rows to archive.`;
    }

    // Append identified rows to the archive sheet
    if (rowsToArchiveData.length > 0) {
        archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchiveData.length, numColumns)
              .setValues(rowsToArchiveData);
        console.log(`Appended ${rowsToArchiveData.length} rows to ${archiveSheetName}.`);
    } else {
         console.log(`${sourceSheetName}: Although rows were marked for deletion, no data was generated to append (this might indicate an issue).`);
    }


    // Delete rows from the source sheet, starting from the bottom up
    const sortedIndices = Array.from(sourceRowIndicesToDelete).sort((a, b) => b - a); // Sort descending
    console.log(`Preparing to delete ${sortedIndices.length} rows from ${sourceSheetName}...`);

    // Group contiguous rows for efficient deletion
    let deleteCount = 0;
    for (let i = 0; i < sortedIndices.length; ) {
        let startRowToDelete = sortedIndices[i];
        let numRowsToDelete = 1;
        // Check subsequent indices to see if they are contiguous
        while (i + numRowsToDelete < sortedIndices.length && sortedIndices[i + numRowsToDelete] === startRowToDelete - numRowsToDelete) {
            numRowsToDelete++;
        }
        console.log(`Deleting ${numRowsToDelete} rows starting from row ${startRowToDelete - numRowsToDelete + 1}`);
        try {
            sourceSheet.deleteRows(startRowToDelete - numRowsToDelete + 1, numRowsToDelete);
            deleteCount += numRowsToDelete;
        } catch (e) {
             console.error(`Error deleting rows ${startRowToDelete - numRowsToDelete + 1} to ${startRowToDelete}: ${e.message}`, e.stack);
             // Optionally try deleting one by one as a fallback? Might be too slow.
             _sendAdminAlert('Archiving Delete Failed', `Error deleting rows from ${sourceSheetName}.\nRange: ${startRowToDelete - numRowsToDelete + 1} for ${numRowsToDelete} rows.\nError: ${e.message}`);
             // Continue to next non-contiguous block if possible
        }

        i += numRowsToDelete; // Move index past the processed block
    }


    SpreadsheetApp.flush(); // Ensure changes are written
    console.log(`${sourceSheetName}: Deleted ${deleteCount} rows. Archived total ${rowsToArchiveData.length} rows.`);
    return `${sourceSheetName}: Archived ${rowsToArchiveData.length} rows (${highWaterRowsArchived} due to high-water). Deleted ${deleteCount} rows.`;
}
