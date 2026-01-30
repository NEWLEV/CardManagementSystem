/**
 * Serves the web app using templated HTML.
 * This is corrected to use createTemplateFromFile to support <?!= ... ?> tags.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('Card Management');
}

/**
 * Helper function to include other .html files (like CSS or JS) into the main index.html.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Creates Admin Tools menu and checks triggers on spreadsheet open.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('Admin Tools')
      .addItem('Setup Permissions', 'setupSheetPermissions')
      .addItem('Setup Maintenance Triggers', 'setupMaintenanceTriggers')
      .addItem('Run Archiving Manually', 'archiveOldData')
      .addItem('Clear Server Cache', 'clearServerCache')
      .addItem('Show Admin Function List', 'showAdminFunctions')
      .addToUi();
  
  // Check and setup triggers if missing
  const triggers = ScriptApp.getProjectTriggers();
  const requiredTriggers = ['scheduledQuickDiagnosis', 'scheduledDeepMaintenance', 'archiveOldData'];
  let foundTriggers = requiredTriggers.filter(funcName =>
      triggers.some(t => t.getHandlerFunction() === funcName)
  );
  if (foundTriggers.length < requiredTriggers.length) {
    console.log("Maintenance triggers missing or incomplete. Attempting setup.");
    try {
      console.warn("Triggers need setup. Please run 'Setup Maintenance Triggers' from the Admin Tools menu.");
    } catch (err) {
      console.error("Failed to check maintenance triggers:", err);
    }
  }
}
