/**
 * Constants - Configuration and field definitions for the timecard reporting system
 * Centralized configuration for API fields, filters, and system settings
 */

// === API CONFIGURATION ===

/**
 * Timecard API configuration
 */
const TIMECARD_API_NAME = "pse__Timecard__c";

const TIMECARD_FIELDS = [
  // Project and opportunity information
  "pse__Project__r.pse__Opportunity__r.OpportunityNumber__c",
  "pse__Project__r.OPA_Project_Number__c",
  "pse__Project__r.pse__Start_Date__c",
  "pse__Project__r.pse__End_Date__c",
  "pse__Project__r.Name",
  
  // Milestone information
  "pse__Milestone__r.OPA_Task_Number__c",
  "pse__Milestone__r.Name",
  
  // Resource information
  "pse__Resource__r.Name",
  
  // Timecard specific fields
  "pse__Start_Date__c",
  "pse__Total_Hours__c",
  "pse__Submitted__c",
  
  // Daily notes (excluding Sunday/Saturday for business days only)
  "pse__Sunday_Notes__c",
  "pse__Monday_Notes__c",
  "pse__Tuesday_Notes__c",
  "pse__Wednesday_Notes__c",
  "pse__Thursday_Notes__c",
  "pse__Friday_Notes__c",
  "pse__Saturday_Notes__c"
];

const TIMECARD_FILTER = [
  // Exclude rejected timecards
  "pse__Status__c+NOT+IN+('Rejected')",
  
  // Only include timecards with actual hours
  "pse__Total_Hours__c+%21%3D+0",
  
  // Exclude non-billable milestones (those starting with 'NB')
  "(NOT+pse__Milestone__r.Name+LIKE+'NB%25')"
];

/**
 * Project API configuration
 */
const PROJECT_API_NAME = "pse__Proj__c";

const PROJECT_FIELDS = [
  // Opportunity information
  "pse__Opportunity__r.OpportunityNumber__c",
  
  // Project basic information
  "Name",
  "pse__Account__c",
  "pse__Account__r.Name",
  "Region_Level_2__c",
  "pse__Stage__c",
  "pse__Start_Date__c",
  "pse__End_Date__c",
  "OPA_Project_Number__c"
];

const PROJECT_FILTER = [
  // Add project-specific filters here if needed
  // Currently using dynamic filtering based on opportunity number
];

// === SHEET CONFIGURATION ===

/**
 * Sheet names and references
 */
const SHEET_NAMES = {
  INSTRUCTIONS: "Instructions",
  PROJECT_DETAILS: "PROJECT_DETAILS"
};

/**
 * Cell references for input and output
 */
const CELL_REFERENCES = {
  // Input cells
  OPPORTUNITY_NUMBER: "B21",
  START_DATE: "B22",
  END_DATE: "B23",
  ROWS_TO_ADD: "D22",
  
  // Display cells
  PROJECT_NAME: "E10",
  START_DATE_DISPLAY: "A2",
  END_DATE_DISPLAY: "A3",
  TOTAL_ROWS: "A1",
  
  // Data ranges
  TIMECARD_DATA_START: "A26",
  PROJECT_DATA_START: "D21",
  REPORT_TEMPLATE: "B13:M13",
  
  // Special cells
  EXPORT_COLOR_CELL: "N1",
  RESET_CHECK_CELL: "B14"
};

/**
 * Data range configurations
 */
const DATA_RANGES = {
  TIMECARD_CLEAR: "A26:I",
  PROJECT_CLEAR: "D21:G21",
  TAB_LIST: "F2:F10"
};

/**
 * Report configuration
 */
const REPORT_CONFIG = {
  FIRST_DATA_ROW: 13,
  TEMPLATE_ROW: 13,
  MAX_ROWS_TO_ADD: 1000,
  EXPORT_ROWS_BASE: 20
};

// === SYSTEM CONFIGURATION ===

/**
 * API retry and timeout settings
 */
const API_SETTINGS = {
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 1000,
  BATCH_SIZE: 2000,
  TIMEOUT_MS: 30000
};

/**
 * PDF export settings
 */
const PDF_SETTINGS = {
  SIZE: "A4",
  PORTRAIT: false,
  FIT_WIDTH: true,
  TOP_MARGIN: 0.75,
  BOTTOM_MARGIN: 0.75,
  LEFT_MARGIN: 0.7,
  RIGHT_MARGIN: 0.7,
  SHOW_SHEET_NAMES: false,
  SHOW_PRINT_TITLE: false,
  SHOW_GRIDLINES: true,
  FREEZE_ROWS: false,
  MAX_RETRIES: 5,
  RETRY_DELAY_MS: 3000
};

/**
 * User interface settings
 */
const UI_SETTINGS = {
  DIALOG_WIDTH: 400,
  DIALOG_HEIGHT: 200,
  INFO_DIALOG_WIDTH: 500,
  INFO_DIALOG_HEIGHT: 400,
  PROGRESS_UPDATE_INTERVAL: 100
};

// === VALIDATION RULES ===

/**
 * Input validation rules
 */
const VALIDATION_RULES = {
  OPPORTUNITY_NUMBER: {
    required: true,
    minLength: 1,
    maxLength: 50
  },
  DATE_RANGE: {
    required: true,
    maxDays: 365
  },
  ROWS_TO_ADD: {
    min: 1,
    max: 1000
  }
};

// === ERROR MESSAGES ===

/**
 * Japanese error messages for user interface
 */
const ERROR_MESSAGES = {
  OPPORTUNITY_NUMBER_REQUIRED: "オポチュニティ番号が必要です",
  INVALID_DATE_RANGE: "有効な日付範囲を入力してください",
  DATE_RANGE_TOO_LARGE: "日付範囲が大きすぎます（最大365日）",
  START_DATE_AFTER_END_DATE: "開始日は終了日より前である必要があります",
  INVALID_ROWS_TO_ADD: "追加する行数は1から1000の間で入力してください",
  API_CONNECTION_ERROR: "APIへの接続に失敗しました",
  SHEET_UPDATE_ERROR: "シートの更新に失敗しました",
  PDF_EXPORT_ERROR: "PDFエクスポートに失敗しました",
  GENERIC_ERROR: "エラーが発生しました"
};

/**
 * Success messages
 */
const SUCCESS_MESSAGES = {
  TIMECARD_UPDATE_SUCCESS: "タイムカードデータが正常に更新されました",
  PROJECT_UPDATE_SUCCESS: "プロジェクトデータが正常に更新されました",
  REPORT_RESET_SUCCESS: "レポートが正常にリセットされました",
  REPORT_CREATE_SUCCESS: "レポートが正常に作成されました",
  PDF_EXPORT_SUCCESS: "PDFが正常にエクスポートされました",
  TAB_DELETE_SUCCESS: "不要なタブが正常に削除されました"
};

// === HELPER FUNCTIONS ===

/**
 * Get error message by key
 * @param {string} key - Error message key
 * @returns {string} Error message
 */
function getErrorMessage(key) {
  return ERROR_MESSAGES[key] || ERROR_MESSAGES.GENERIC_ERROR;
}

/**
 * Get success message by key
 * @param {string} key - Success message key
 * @returns {string} Success message
 */
function getSuccessMessage(key) {
  return SUCCESS_MESSAGES[key] || "操作が完了しました";
}

/**
 * Get cell reference by key
 * @param {string} key - Cell reference key
 * @returns {string} Cell reference
 */
function getCellReference(key) {
  return CELL_REFERENCES[key] || "";
}

/**
 * Get data range by key
 * @param {string} key - Data range key
 * @returns {string} Data range
 */
function getDataRange(key) {
  return DATA_RANGES[key] || "";
}

/**
 * Get configuration value
 * @param {string} section - Configuration section
 * @param {string} key - Configuration key
 * @returns {*} Configuration value
 */
function getConfig(section, key) {
  const sections = {
    REPORT: REPORT_CONFIG,
    API: API_SETTINGS,
    PDF: PDF_SETTINGS,
    UI: UI_SETTINGS
  };
  
  return sections[section] && sections[section][key];
}

// === BACKWARD COMPATIBILITY ===

/**
 * Legacy constants for backward compatibility
 * These are maintained for any existing code that might reference them
 */
const PROJECT_SHEET_NAME = SHEET_NAMES.PROJECT_DETAILS;

// Export key constants for external use
// (In Apps Script, these are available globally, but documenting for clarity)
/* globals: 
 * TIMECARD_FIELDS, TIMECARD_API_NAME, TIMECARD_FILTER,
 * PROJECT_FIELDS, PROJECT_API_NAME, PROJECT_FILTER,
 * SHEET_NAMES, CELL_REFERENCES, DATA_RANGES, REPORT_CONFIG,
 * API_SETTINGS, PDF_SETTINGS, UI_SETTINGS, VALIDATION_RULES,
 * ERROR_MESSAGES, SUCCESS_MESSAGES
 */