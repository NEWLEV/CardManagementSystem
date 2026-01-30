// --- FILE: Data.gs ---
// Contains core logic for data retrieval, caching, and processing.

/**
 * ENHANCEMENT: Caches the master list of all cards from the 'Cards' sheet.
 */
function _getAllCardsFromInventory() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEYS.MASTER_INVENTORY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      const inventoryMap = new Map();
      // Ensure data is parsed correctly back into Sets
      for (const [type, numbers] of Object.entries(parsed)) {
          if (Array.isArray(numbers)) { // Check if it's an array before creating Set
            inventoryMap.set(type, new Set(numbers));
          } else {
             console.warn(`_getAllCardsFromInventory: Cached data for type "${type}" is not an array. Rebuilding cache might be needed.`);
             inventoryMap.set(type, new Set()); // Initialize as empty set if format is wrong
          }
      }
      // console.log("Master inventory loaded from cache."); // Optional: uncomment for debugging cache hits
      return inventoryMap;
    } catch (e) {
      console.error("Error parsing cached master inventory. Rebuilding.", e.message, e.stack);
      cache.remove(CACHE_KEYS.MASTER_INVENTORY); // Clear bad cache entry
    }
  }

  console.log('Cache miss/rebuild: Building master card inventory from sheet...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cardSheet = ss.getSheetByName(SHEETS.CARDS);
  const inventoryMap = new Map();
  // Ensure CARD_TYPES is defined (should be from Config.gs)
   if (typeof CARD_TYPES === 'undefined' || !Array.isArray(CARD_TYPES)) {
       console.error("CARD_TYPES constant is not defined or not an array. Cannot build inventory.");
       _sendAdminAlert("Inventory Build Failed", "CARD_TYPES constant missing or invalid in Config.gs");
       return inventoryMap; // Return empty map
   }

  const typeMap = CARD_TYPES.reduce((map, type, index) => {
      map[type] = index + 1; // 1-based column index
      return map;
  }, {});

  if (cardSheet) {
    const lastRow = cardSheet.getLastRow();
    if (lastRow > 1) { // Header row exists
      const numCols = cardSheet.getLastColumn();
      const dataRange = cardSheet.getRange(2, 1, lastRow - 1, numCols);
      console.log(`Reading master inventory from ${SHEETS.CARDS} range ${dataRange.getA1Notation()}`);
      const data = dataRange.getValues();

      for (const type of CARD_TYPES) {
          const colIndex = typeMap[type] - 1; // 0-based index for array access
          if (colIndex >= numCols) {
              console.warn(`Column for card type "${type}" (expected ${colIndex + 1}) exceeds sheet width (${numCols}). Skipping type.`);
              inventoryMap.set(type, new Set()); // Initialize empty if column missing
              continue;
          }
          const cardSet = new Set();
          data.forEach(row => {
            // Check if the row array actually has an element at this index
            if (row.length > colIndex) {
                const card = _normalizeCardNumber(row[colIndex]);
                if (card) { // Only add if not blank after normalization
                  cardSet.add(card);
                }
            }
          });
          inventoryMap.set(type, cardSet);
          console.log(`Found ${cardSet.size} cards for type "${type}" in master list.`);
      }
    } else {
        console.log(`Sheet "${SHEETS.CARDS}" only has header row or is empty. Initializing empty sets.`);
        CARD_TYPES.forEach(type => inventoryMap.set(type, new Set()));
    }
  } else {
     console.error(`Sheet "${SHEETS.CARDS}" not found! Cannot build master inventory.`);
     _sendAdminAlert('Master Inventory Build Failed', `Sheet "${SHEETS.CARDS}" not found.`);
     CARD_TYPES.forEach(type => inventoryMap.set(type, new Set())); // Initialize empty
  }

  // Save the built map to cache
  try {
    const serializable = {};
    for (const [type, numberSet] of inventoryMap.entries()) {
      serializable[type] = [...numberSet]; // Convert Set to Array for JSON
    }
    const jsonString = JSON.stringify(serializable);
     console.log(`Saving master inventory to cache (${jsonString.length} chars). Duration: ${CACHE_DURATION}s`);
    cache.put(CACHE_KEYS.MASTER_INVENTORY, jsonString, CACHE_DURATION);
  } catch (e) {
     console.error("Error saving master inventory to cache:", e.message, e.stack);
     // Don't alert for cache save errors unless persistent
  }
  return inventoryMap;
}

/**
 * V4 PERFORMANCE FIX + Logging: Gets used cards, reading in batches.
 * Reads from BOTH active and archived audit logs.
 */
function _getUsedCardSet() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEYS.USED_CARDS);
  if (cached) {
    try {
      const parsedSet = new Set(JSON.parse(cached));
      // console.log(`Used card set loaded from cache. Size: ${parsedSet.size}`); // Optional: Debug cache hits
      return parsedSet;
    } catch (e) {
      console.error("Error parsing cached used card set. Rebuilding.", e.message, e.stack);
      cache.remove(CACHE_KEYS.USED_CARDS); // Clear bad cache entry
    }
  }

  console.log('Cache miss/rebuild: Building used card set from sheet(s)...');
  const startTime = Date.now();
  const BATCH_SIZE = 10000; // Process 10k rows at a time
  const TIME_LIMIT_MS = 240000; // 4 minutes Apps Script limit is 6 mins, leave buffer

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usedCardsSet = new Set();

  // Helper function to process a sheet in batches
  const processSheetInBatches = (sheet, counterRef) => { // counterRef should be like { count: 0 }
    if (!sheet) {
      console.log(`_getUsedCardSet processSheet: Sheet object is null/undefined, skipping.`);
      return;
    }

    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { // Need header + data
      console.log(`_getUsedCardSet processSheet: No data rows in ${sheetName} (lastRow=${lastRow}), skipping.`);
      return;
    }

    console.log(`_getUsedCardSet processSheet: Processing ${sheetName} (${lastRow - 1} data rows)...`);
    let rowsProcessedInSheet = 0;

    for (let startRow = 2; startRow <= lastRow; startRow += BATCH_SIZE) {
      const elapsedTime = Date.now() - startTime;
      if (elapsedTime > TIME_LIMIT_MS) {
        console.warn(`_getUsedCardSet: Time limit approaching (> ${TIME_LIMIT_MS / 1000}s) while processing ${sheetName} at row ${startRow}. Stopping build. Set will be incomplete.`);
        _sendAdminAlert('_getUsedCardSet Timeout', `Building the used card set took too long (>${TIME_LIMIT_MS / 1000}s) and was aborted while processing ${sheetName}. The set is incomplete. Ensure archiving runs successfully.`);
        return; // Stop processing this sheet and return prematurely
      }

      const numRowsToFetch = Math.min(BATCH_SIZE, lastRow - startRow + 1);
      if (numRowsToFetch <= 0) break; // Should not happen with loop condition, but safety check

      // Read Type (col 2) and Card Number (col 3)
      const range = sheet.getRange(startRow, 2, numRowsToFetch, 2);
      // console.log(`   Reading batch from ${sheetName}: ${range.getA1Notation()}`); // Verbose logging
      const auditData = range.getValues();
      let batchAdds = 0;

      auditData.forEach(row => {
        const type = (row[0] || '').toString().trim(); // Type is in the first column read (col B)
        const card = _normalizeCardNumber(row[1]);    // Card Number is second (col C)
        if (type && card) {
          const key = `${type}|${card}`;
          if (!usedCardsSet.has(key)) { // Check if it's already in the set
            usedCardsSet.add(key);
            counterRef.count++; // Increment the counter passed by reference
            batchAdds++;
          }
        }
      });
      rowsProcessedInSheet += auditData.length;
      console.log(`   Processed batch from ${sheetName} up to row ${startRow + numRowsToFetch - 1}. Added ${batchAdds} new unique cards. Current total set size: ${usedCardsSet.size}`);
    }
    console.log(`_getUsedCardSet processSheet: Finished ${sheetName}. Processed ${rowsProcessedInSheet} rows. Added ${counterRef.count} unique entries in total from this sheet.`);
  }; // End of processSheetInBatches helper

  // Process Active Log
  const auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
  let activeCounter = { count: 0 }; // Use object wrapper to pass by reference
  if (!auditSheet) {
      console.error(`Sheet "${SHEETS.AUDIT_LOG}" not found! Used card set build might be incomplete.`);
      _sendAdminAlert('Used Card Set Build Failed', `Sheet "${SHEETS.AUDIT_LOG}" not found.`);
  } else {
    processSheetInBatches(auditSheet, activeCounter);
  }

  // Process Archive Log
  const archivedAuditSheet = ss.getSheetByName(SHEETS.ARCHIVED_AUDIT_LOG);
  let archiveCounter = { count: 0 }; // Use object wrapper
  if (!archivedAuditSheet) {
    console.log(`Note: Archive sheet named "${SHEETS.ARCHIVED_AUDIT_LOG}" not found. Skipping.`);
  } else {
    // Only proceed if no timeout occurred during active log processing
    if (Date.now() - startTime <= TIME_LIMIT_MS) {
        processSheetInBatches(archivedAuditSheet, archiveCounter);
    } else {
         console.warn(`Skipping processing of ${SHEETS.ARCHIVED_AUDIT_LOG} due to prior timeout.`);
    }
  }

  const totalTime = Date.now() - startTime;
  console.log(`_getUsedCardSet: Build complete. Total unique used cards: ${usedCardsSet.size} (Active Log: ${activeCounter.count}, Archive Log: ${archiveCounter.count}). Total Time: ${totalTime} ms.`);

  // Save the complete set to cache if it didn't time out
  if (totalTime <= TIME_LIMIT_MS) {
      try {
          const jsonString = JSON.stringify([...usedCardsSet]);
          console.log(`Saving used card set to cache (${jsonString.length} chars). Duration: ${CACHE_DURATION}s`);
         cache.put(CACHE_KEYS.USED_CARDS, jsonString, CACHE_DURATION);
      } catch (e) {
         console.error("Error saving used card set to cache:", e.message, e.stack);
         // Don't alert unless persistent
      }
  } else {
       console.warn("Used card set build timed out. Cache will not be updated with potentially incomplete data.");
  }

  return usedCardsSet;
}


/**
 * Checks card availability using server-side cache.
 */
function isCardAvailable(type, cardNumber) {
  try {
    if (!type || !cardNumber) return false;
    const cleanType = type.toString().trim();
    const cleanCardNumber = _normalizeCardNumber(cardNumber);
    if (!cleanType || !cleanCardNumber) return false;

    const cardKey = `${cleanType}|${cleanCardNumber}`;
    // Always call _getUsedCardSet - it handles caching internally
    const usedCardsSet = _getUsedCardSet();
    return !usedCardsSet.has(cardKey);
  } catch (error) {
    console.error(`isCardAvailable failed for ${type} ${cardNumber}: ${error.message}`, error.stack);
    // In case of error, safer to assume card is NOT available
    return false;
  }
}

/**
 * ENHANCED: Gets card counts for the UI using cached data.
 * Helper for getAppData().
 */
function getCardCounts() {
  try {
    const counts = {};
    // Get data using functions that handle caching
    const masterInventoryMap = _getAllCardsFromInventory();
    const usedCardsSet = _getUsedCardSet();

    if (typeof CARD_TYPES === 'undefined' || !Array.isArray(CARD_TYPES)) {
       console.error("getCardCounts: CARD_TYPES is missing. Cannot calculate counts.");
       return {};
    }

    for (const type of CARD_TYPES) {
      const allCardsForType = masterInventoryMap.get(type); // This is expected to be a Set
      if (!allCardsForType || !(allCardsForType instanceof Set) || allCardsForType.size === 0) {
        counts[type] = 0;
      } else {
        let availableCount = 0;
        // Iterate over the Set from the master inventory
        for (const card of allCardsForType) {
          // Check against the used set
          if (!usedCardsSet.has(`${type}|${card}`)) {
            availableCount++;
          }
        }
        counts[type] = availableCount;
      }
    }
    // console.log("Calculated available card counts:", counts); // Optional: Debug counts
    return counts;
  } catch (error) {
    console.error('getCardCounts failed:', error.message, error.stack);
    _sendAdminAlert('Card Count Failed', `Error calculating available counts: ${error.message}`);
    // Return default counts on error
    return CARD_TYPES.reduce((acc, type) => { acc[type] = 0; return acc; }, {});
  }
}

/**
 * ENHANCED: Gets the structured list of all "unused" cards for the public UI.
 * Helper for getAppData().
 */
function getPublicInventoryList() {
  try {
    const availableCardsMap = {};
    // Get data using functions that handle caching
    const masterInventoryMap = _getAllCardsFromInventory();
    const usedCardsSet = _getUsedCardSet();

     if (typeof CARD_TYPES === 'undefined' || !Array.isArray(CARD_TYPES)) {
       console.error("getPublicInventoryList: CARD_TYPES is missing. Cannot get list.");
       return {};
    }


    for (const type of CARD_TYPES) {
      const allCardsForType = masterInventoryMap.get(type); // Expected to be a Set
      if (!allCardsForType || !(allCardsForType instanceof Set) || allCardsForType.size === 0) {
        availableCardsMap[type] = [];
      } else {
        const available = [];
        // Iterate Set from master inventory
        for (const card of allCardsForType) {
          // Check against used Set
          if (!usedCardsSet.has(`${type}|${card}`)) {
            available.push(card);
          }
        }
        available.sort(); // Sort the available list
        availableCardsMap[type] = available;
      }
    }
    // console.log("Generated public inventory list."); // Optional: Debug
    return availableCardsMap;

  } catch (e) {
    console.error(`getPublicInventoryList error: ${e.message}`, e.stack);
    _sendAdminAlert('getPublicInventoryList Failure', `Error: ${e.message}\nStack: ${e.stack}`);
    // Return empty structure on error
    return CARD_TYPES.reduce((acc, type) => { acc[type] = []; return acc; }, {});
  }
}
