/**
 * SheetService - Handles all Google Sheets operations
 * Provides standardized methods for reading and writing data with proper error handling
 */

/**
 * Sheet configuration - centralized cell references and sheet names
 */
const SHEET_CONFIG = {
  // Input cells
  OPPORTUNITY_NUMBER_CELL: "B21",
  START_DATE_CELL: "B22",
  END_DATE_CELL: "B23",
  ROWS_TO_ADD_CELL: "D22",
  PROJECT_NAME_CELL: "E10",
  START_DATE_DISPLAY_CELL: "A2",
  END_DATE_DISPLAY_CELL: "A3",
  TOTAL_ROWS_CELL: "A1",
  
  // Data ranges
  TIMECARD_DATA_RANGE: "A26:I",
  PROJECT_DATA_RANGE: "D21:G21",
  REPORT_TEMPLATE_RANGE: "B13:M13",
  
  // Report generation
  FIRST_DATA_ROW: 13,
  REPORT_EXPORT_COLOR_CELL: "N1",
  
  // Instructions sheet
  INSTRUCTIONS_SHEET: "Instructions",
  TAB_LIST_RANGE: "F2:F10"
};

/**
 * GPS code mapping for milestone names
 * Maps milestone names to standardized GPS codes
 */
const GPS_CODE_MAPPING = {
  // Common GPS codes mapping
  'project manager': 'GPS-PJM',
  'principal consultant': 'GPS-PRCON',
  'senior consultant': 'GPS-SC',
  'consultant': 'GPS-C',
  'principal architect': 'GPS-PA',
  'senior architect': 'GPS-SA',
  'architect': 'GPS-A',
};

/**
 * Map milestone name to GPS code
 * @param {string} milestoneName - The milestone name to map
 * @returns {string} GPS code or original name if no mapping found
 */
function mapToGPSCode(milestoneName) {
  if (!milestoneName || typeof milestoneName !== 'string') {
    return milestoneName; // Return original for empty/invalid names
  }
  
  // First, check if the milestone already contains a GPS code pattern
  const gpsCodePattern = /GPS-[A-Z]+/i;
  const existingGPSMatch = milestoneName.match(gpsCodePattern);
  
  if (existingGPSMatch) {
    // Extract and return the existing GPS code (uppercase)
    return existingGPSMatch[0].toUpperCase();
  }
  
  const lowerName = milestoneName.toLowerCase().trim();
  
  // Check for exact matches first
  if (GPS_CODE_MAPPING[lowerName]) {
    return GPS_CODE_MAPPING[lowerName];
  }
  
  // Check for partial matches
  for (const [key, value] of Object.entries(GPS_CODE_MAPPING)) {
    if (lowerName.includes(key)) {
      return value;
    }
  }
  
  // If no mapping found, return the original milestone name
  return milestoneName;
}

/**
 * Get input parameters from the active sheet
 * @returns {Object} Object containing all input parameters
 * @throws {Error} If required parameters are missing or invalid
 */
function getInputParameters() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    const opportunityNumber = sheet.getRange(SHEET_CONFIG.OPPORTUNITY_NUMBER_CELL).getValue();
    const startDate = sheet.getRange(SHEET_CONFIG.START_DATE_CELL).getValue();
    const endDate = sheet.getRange(SHEET_CONFIG.END_DATE_CELL).getValue();
    const rowsToAdd = sheet.getRange(SHEET_CONFIG.ROWS_TO_ADD_CELL).getValue();
    
    // Validate required parameters
    if (!opportunityNumber) {
      throw new Error(`Opportunity number is required in cell ${SHEET_CONFIG.OPPORTUNITY_NUMBER_CELL}`);
    }
    
    if (!startDate || !(startDate instanceof Date)) {
      throw new Error(`Valid start date is required in cell ${SHEET_CONFIG.START_DATE_CELL}`);
    }
    
    if (!endDate || !(endDate instanceof Date)) {
      throw new Error(`Valid end date is required in cell ${SHEET_CONFIG.END_DATE_CELL}`);
    }
    
    if (startDate > endDate) {
      throw new Error("Start date must be before end date");
    }
    
    return {
      opportunityNumber: String(opportunityNumber).trim(),
      startDate: new Date(startDate),
      endDate: new Date(endDate),
      rowsToAdd: Number(rowsToAdd) || 0
    };
    
  } catch (error) {
    console.error(`Error getting input parameters: ${error.message}`);
    throw error;
  }
}

/**
 * Update timecard data in the sheet
 * @param {Array} timecardRecords - Array of timecard records
 * @throws {Error} If update fails
 */
function updateTimecardData(timecardRecords) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Clear existing data
    sheet.getRange(SHEET_CONFIG.TIMECARD_DATA_RANGE).clearContent();
    
    if (!timecardRecords || timecardRecords.length === 0) {
      console.log("No timecard records to update");
      return;
    }
    
    // Transform records to sheet format
    const sheetData = timecardRecords.map(record => [
      record.pse__Resource__r.Name,
      record.pse__Start_Date__c,
      record.pse__Total_Hours__c,
      record.pse__Monday_Notes__c,
      record.pse__Tuesday_Notes__c,
      record.pse__Wednesday_Notes__c,
      record.pse__Thursday_Notes__c,
      record.pse__Friday_Notes__c,
      mapToGPSCode(record.pse__Milestone__r.Name), 
    ]);
    
    // Write data to sheet
    if (sheetData.length > 0) {
      const range = sheet.getRange(26, 1, sheetData.length, sheetData[0].length);
      range.setValues(sheetData);
      console.log(`Updated ${sheetData.length} timecard records with GPS codes`);
    }
    
  } catch (error) {
    console.error(`Error updating timecard data: ${error.message}`);
    throw error;
  }
}

/**
 * Update project data in the sheet
 * @param {Array} projectRecords - Array of project records
 * @throws {Error} If update fails
 */
function updateProjectData(projectRecords) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Clear existing data
    sheet.getRange(SHEET_CONFIG.PROJECT_DATA_RANGE).clearContent();
    
    if (!projectRecords || projectRecords.length === 0) {
      console.log("No project records to update");
      return;
    }
    
    // Transform records to sheet format
    const sheetData = projectRecords.map(record => {
      const accountName = record.pse__Account__c ? record.pse__Account__r.Name : "";
      return [
        record.Name,
        accountName,
        record.OPA_Project_Number__c
      ];
    });
    
    // Write first project record to sheet
    if (sheetData.length > 0) {
      const range = sheet.getRange(21, 4, 1, sheetData[0].length);
      range.setValues([sheetData[0]]);
      console.log(`Updated project data: ${sheetData[0].join(", ")}`);
    }
    
  } catch (error) {
    console.error(`Error updating project data: ${error.message}`);
    throw error;
  }
}

/**
 * Reset the report sheet by removing data rows
 * @throws {Error} If reset fails
 */
function resetReportSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Check if already reset
    const checkCell = sheet.getRange('B14').getValue();
    if (checkCell === "稼働時間合計") {
      SpreadsheetApp.getUi().alert('シートはすでにリセット済みです');
      return;
    }
    
    // Calculate rows to delete
    const totalRows = sheet.getRange(SHEET_CONFIG.TOTAL_ROWS_CELL).getValue();
    const firstRow = SHEET_CONFIG.FIRST_DATA_ROW;
    const lastRow = firstRow + totalRows;
    
    // Delete data rows
    const rangeToDelete = sheet.getRange(`${firstRow + 1}:${lastRow}`);
    sheet.deleteRows(rangeToDelete.getRow(), rangeToDelete.getNumRows());
    
    // Reset H13 to formula =I26
    //sheet.getRange('H13').setFormula('=I26');
    
    console.log(`Reset complete: Deleted rows ${firstRow + 1} to ${lastRow}, Reset H13 to =I26`);
    
  } catch (error) {
    console.error(`Error resetting report sheet: ${error.message}`);
    throw error;
  }
}

/**
 * Create report by adding data rows and copying formulas
 * @throws {Error} If creation fails
 */
function createReportSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Get number of rows to add
    const rowsToAdd = sheet.getRange(SHEET_CONFIG.ROWS_TO_ADD_CELL).getValue();
    
    // Validate input
    if (!rowsToAdd || rowsToAdd <= 0 || rowsToAdd > 1000) {
      SpreadsheetApp.getUi().alert(
        'セルD22に追加したい行数が入るようにしてください。\n' +
        'Sampleタブのように13行と合計の行だけが残るようにすればD22を行数にすることができます。\n' +
        '不安な場合はSampleをDuplicateしてからマクロを動かすことをお勧めします。'
      );
      return;
    }
    
    // Insert rows after row 13
    console.log(`Inserting ${rowsToAdd} rows after row 13`);
    sheet.insertRowsAfter(13, rowsToAdd);
    
    // Copy formulas from template row to new rows
    const templateRange = sheet.getRange(SHEET_CONFIG.REPORT_TEMPLATE_RANGE);
    const newDataRange = sheet.getRange(14, 2, rowsToAdd, 12); // B14 to M(13+rowsToAdd)
    
    templateRange.copyTo(newDataRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
    // Set sum formula
    const sumCell = sheet.getRange(14 + rowsToAdd, 6); // F(14+rowsToAdd)
    sumCell.setFormulaR1C1(`=SUM(R[-${rowsToAdd + 1}]C[0]:R[-1]C[1])`);
    
    console.log(`Report creation complete: Added ${rowsToAdd} rows with formulas`);
    
  } catch (error) {
    console.error(`Error creating report sheet: ${error.message}`);
    throw error;
  }
}

/**
 * Get report export parameters
 * @returns {Object} Object containing export parameters
 */
function getExportParameters() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    const totalRows = sheet.getRange(SHEET_CONFIG.TOTAL_ROWS_CELL).getValue();
    const projectName = sheet.getRange(SHEET_CONFIG.PROJECT_NAME_CELL).getValue();
    const startDate = sheet.getRange(SHEET_CONFIG.START_DATE_DISPLAY_CELL).getValue();
    const endDate = sheet.getRange(SHEET_CONFIG.END_DATE_DISPLAY_CELL).getValue();
    
    const exportRows = 20 + totalRows;
    const exportRange = sheet.getRange(`A1:N${exportRows}`);
    
    return {
      totalRows,
      projectName,
      startDate,
      endDate,
      exportRows,
      exportRange
    };
    
  } catch (error) {
    console.error(`Error getting export parameters: ${error.message}`);
    throw error;
  }
}

/**
 * Set export color for PDF generation
 * @param {string} color - Color to set ('white' or 'black')
 */
function setExportColor(color) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(SHEET_CONFIG.REPORT_EXPORT_COLOR_CELL).setFontColor(color);
  } catch (error) {
    console.error(`Error setting export color: ${error.message}`);
    throw error;
  }
}

/**
 * Get list of allowed tabs from Instructions sheet
 * @returns {Array} Array of allowed tab names
 */
function getAllowedTabs() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const instructionsSheet = spreadsheet.getSheetByName(SHEET_CONFIG.INSTRUCTIONS_SHEET);
    
    if (!instructionsSheet) {
      console.warn("Instructions sheet not found");
      return [];
    }
    
    const listRange = instructionsSheet.getRange(SHEET_CONFIG.TAB_LIST_RANGE);
    const listValues = listRange.getValues();
    
    // Flatten and filter out empty values
    const allowedTabs = listValues
      .map(row => row[0])
      .filter(value => value && String(value).trim() !== "")
      .map(value => String(value).toLowerCase());
    
    return allowedTabs;
    
  } catch (error) {
    console.error(`Error getting allowed tabs: ${error.message}`);
    return [];
  }
}

/**
 * Delete tabs not in the allowed list
 * @param {Array} allowedTabs - Array of allowed tab names
 */
function deleteUnallowedTabs(allowedTabs) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    // Process sheets in reverse order to avoid index shifting
    for (let i = sheets.length - 1; i >= 0; i--) {
      const sheet = sheets[i];
      const sheetName = sheet.getName().toLowerCase();
      const isHidden = sheet.isSheetHidden();
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      
      // Delete if not in allowed list and either hidden or unprotected
      if (!allowedTabs.includes(sheetName) && (isHidden || protections.length === 0)) {
        console.log(`Deleting sheet: ${sheet.getName()}`);
        spreadsheet.deleteSheet(sheet);
      }
    }
    
  } catch (error) {
    console.error(`Error deleting unallowed tabs: ${error.message}`);
    throw error;
  }
} 