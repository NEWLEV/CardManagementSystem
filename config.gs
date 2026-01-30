// --- FILE: Config.gs ---
// Contains all global constants and configuration settings for the app.

const SHEETS = {
  DATA: 'Data',
  CARDS: 'Cards',
  AUDIT_LOG: 'Card Audit Log',
  ARCHIVED_DATA: 'Archived_Data',
  ARCHIVED_AUDIT_LOG: 'Archived_Card_Audit_Log',
  APPROVAL_CODES: 'ApprovalCodes'
};

// Column constants - UPDATED
const COLUMNS = {
  DATA_DATE: 1,      // A
  DATA_NAME: 2,      // B
  DATA_PRN: 3,       // C - NEW
  DATA_TYPE: 4,      // D
  DATA_NUMBER: 5,    // E
  SIGNATURE: 6,      // F
  DATA_REASON: 7,    // G
  DATA_NOTE: 8,      // H
  DATA_TIMESTAMP: 9, // I
  BASE64: 10,        // J
  DATA_ENTRY_ID: 11  // K - Shifted
};

// Config for row heights
const ROW_HEIGHT_CONFIG = {
    STANDARD: 25,
};

const ADMIN_EMAILS = [
  'pierremontalvo@continentalwellnesscenter.com'
  // Add other admin emails here, lowercase
];

const EXPORT_SHEET_ID = '1zET1FMweB80B7U5kw-jvtcWjr4ORjrw_Gn306l8ORXg';
const CARD_TYPES = ["McDonald's", "Publix", "Walmart", "Bus Pass"];

const CACHE_KEYS = {
  USED_CARDS: 'USED_CARDS_SET',
  MASTER_INVENTORY: 'MASTER_CARD_INVENTORY_CACHE'
};
const CACHE_DURATION = 21600; // 6 hours

// --- ARCHIVING CONFIGURATION ---
const ARCHIVE_CONFIG = {
  ARCHIVE_DAYS_AGO: 90,
  MAX_ACTIVE_ROWS: 20000,
  ROWS_TO_TRIM_ON_HIGH_WATER: 5000
};
