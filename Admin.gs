// --- FILE: Admin.gs ---
// Contains functions for the Spreadsheet Admin Menu and user authentication.

/**
 * Checks if the current effective user is an admin based on ADMIN_EMAILS list.
 */
function isUserAdmin() {
  try {
    const email = Session.getEffectiveUser().getEmail();
     if (!email) return false;
    return ADMIN_EMAILS.map(adminEmail => adminEmail.toLowerCase()).includes(email.toLowerCase());
  } catch (error) {
     console.error("Error checking admin status:", error);
    return false;
  }
}

/**
 * Sets up sheet permissions for protected sheets.
 */
function setupSheetPermissions() {
  try {
    console.log("Setting up sheet permissions...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = [SHEETS.CARDS, SHEETS.APPROVAL_CODES];
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        console.warn(`Sheet "${sheetName}" not found. Skipping protection.`);
        return;
      }
      
      // Remove existing protections
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(p => p.remove());
      
      // Add new protection
      const protection = sheet.protect();
      protection.setDescription(`Protected ${sheetName} sheet`);
      protection.setWarningOnly(false);
      
      // Add admins as editors
      const adminsToAdd = ADMIN_EMAILS.filter(email => email && email.length > 0);
      if (adminsToAdd.length > 0) {
        protection.addEditors(adminsToAdd);
      }
      
      console.log(`Protected sheet: ${sheetName}`);
    });
    
    SpreadsheetApp.getUi().alert("Sheet permissions have been set up successfully!");
    console.log("Sheet permissions setup complete.");
  } catch (error) {
    console.error("Error setting up sheet permissions:", error.message);
    _sendAdminAlert('Sheet Permissions Setup Failed', error.message);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

/**
 * Makes a card available again by removing it from the audit log.
 * Note: This is an Admin Menu item, as per the original onOpen code.
 */
function makeCardAvailable(type, cardNumber) {
  if (!isUserAdmin()) {
    throw new Error('Permission denied. Admin access required.');
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    
    if (!auditSheet) {
      throw new Error(`Sheet "${SHEETS.AUDIT_LOG}" not found.`);
    }
    
    const cleanType = type.toString().trim();
    const cleanCardNumber = _normalizeCardNumber(cardNumber);
    
    const lastRow = auditSheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: 'Audit log is empty.' };
    }
    
    // Find and delete the matching entry
    const data = auditSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    let rowToDelete = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === cleanType && _normalizeCardNumber(data[i][2]) === cleanCardNumber) {
        rowToDelete = i + 2; // Convert to 1-based row index
        break;
      }
    }
    
    if (rowToDelete > 0) {
      auditSheet.deleteRow(rowToDelete);
      
      // Invalidate cache
      CacheService.getScriptCache().remove(CACHE_KEYS.USED_CARDS);
      
      console.log(`Admin ${getCurrentUserEmail()} made card available: ${type} ${cardNumber}`);
      return { success: true, message: `Card ${cardNumber} (${type}) is now available.` };
    } else {
      return { success: false, message: 'Card not found in audit log.' };
    }
    
  } catch (error) {
    console.error(`makeCardAvailable failed: ${error.message}`);
    _sendAdminAlert('Make Card Available Failed', error.message);
    throw error;
  }
}

/**
 * Clears the server cache.
 */
function clearServerCache() {
  try {
    console.log("Clearing server cache...");
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEYS.USED_CARDS);
    cache.remove(CACHE_KEYS.MASTER_INVENTORY);
    console.log("Server cache cleared successfully.");
    SpreadsheetApp.getUi().alert("Server cache has been cleared successfully!");
    return "Cache cleared";
  } catch (error) {
    console.error("Error clearing cache:", error.message);
    _sendAdminAlert('Cache Clear Failed', error.message);
    throw error;
  }
}

/**
 * Shows a list of available admin functions.
 */
function showAdminFunctions() {
  const helpText = `
ADMIN FUNCTIONS AVAILABLE:

1. Setup Permissions
    - Protects the Cards and ApprovalCodes sheets
    - Adds admin emails as editors

2. Setup Maintenance Triggers
    - Sets up automated daily diagnostics
    - Sets up weekly deep maintenance
    - Sets up monthly archiving

3. Run Archiving Manually
    - Archives old data from Data and Audit Log sheets
    - Moves data older than ${ARCHIVE_CONFIG.ARCHIVE_DAYS_AGO} days

4. Clear Server Cache
    - Clears cached card inventory and usage data
    - Forces fresh data load on next access

5. From Web App Admin Panel:
    - Add Cards: Add new cards to inventory
    - Remove Cards: Remove cards from inventory
    - List Inventory: View all cards in inventory
    - Make Card Available: Remove card from audit log

For more information, contact: ${ADMIN_EMAILS.join(', ')}
  `;
  
  SpreadsheetApp.getUi().alert("Admin Functions", helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Admin function to get all cards in the inventory sheet (master list).
 * Note: This is not currently called by the client app.
 */
function getInventoryList(type) {
  if (!isUserAdmin()) {
    throw new Error('Permission denied. Admin access required.');
  }
  if (!type || !CARD_TYPES.includes(type)) {
    throw new Error('Invalid card type specified.');
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cardSheet = ss.getSheetByName(SHEETS.CARDS);
    if (!cardSheet) {
      throw new Error(`Sheet "${SHEETS.CARDS}" not found.`);
    }
    
    const typeMap = CARD_TYPES.reduce((map, t, i) => { map[t] = i + 1; return map; }, {});
    const col = typeMap[type];
    
    const lastDataRow = cardSheet.getLastRow();
    if (lastDataRow < 2) {
      return [];
    }
    
    const allCards = cardSheet.getRange(2, col, lastDataRow - 1, 1)
      .getValues()
      .flat()
      .map(_normalizeCardNumber)
      .filter(v => v !== '');
      
    return allCards.sort();
    
  } catch (error) {
    console.error(`getInventoryList failed: ${error.message}`);
     _sendAdminAlert('Admin List Inventory Failure', `Error listing cards for ${type}.\n\nError: ${error.message}`);
    throw error;
  }
}

/**
 * Returns web app configuration data to the client, including admin status.
 * Note: This is not currently called by the client app.
 */
function getWebAppConfig() {
  console.warn("getWebAppConfig() called - consider removing if only admin status was needed.");
  try {
    return {
      isAdmin: isUserAdmin()
    };
  } catch (error) {
    console.error("Error in getWebAppConfig:", error);
    return {
      isAdmin: false,
      error: `Error fetching config: ${error.message}`
    };
  }
}
