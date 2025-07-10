/**
 * ReportManager - Orchestrates the entire report generation process
 * Coordinates between ApiService and SheetService for complete report workflow
 */

/**
 * Update timecard data by fetching from API and updating the sheet
 * This is the main entry point for refreshing timecard data
 */
function updateTimecards() {
  try {
    console.log("Starting timecard update process...");
    
    // Get input parameters from sheet
    const params = getInputParameters();
    console.log(`Parameters: Opportunity ${params.opportunityNumber}, ${params.startDate.toDateString()} to ${params.endDate.toDateString()}`);
    
    // Fetch timecard data from API
    const timecardRecords = queryTimecards(params.opportunityNumber, params.startDate, params.endDate);
    
    // Update sheet with timecard data
    updateTimecardData(timecardRecords);
    
    // Also update project data
    updateProjectsData();
    
    console.log("Timecard update process completed successfully");
    
  } catch (error) {
    console.error(`Timecard update failed: ${error.message}`);
    SpreadsheetApp.getUi().alert(`タイムカードの更新中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Update project data by fetching from API and updating the sheet
 */
function updateProjectsData() {
  try {
    console.log("Starting project data update...");
    
    // Get input parameters from sheet
    const params = getInputParameters();
    
    // Fetch project data from API
    const projectRecords = queryProjects(params.opportunityNumber);
    
    // Update sheet with project data
    updateProjectData(projectRecords);
    
    console.log("Project data update completed successfully");
    
  } catch (error) {
    console.error(`Project data update failed: ${error.message}`);
    SpreadsheetApp.getUi().alert(`プロジェクトデータの更新中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Reset the report sheet by removing data rows
 * This function is called from the menu
 */
function resetReport() {
  try {
    console.log("Starting report reset...");
    
    resetReportSheet();
    
    console.log("Report reset completed successfully");
    
  } catch (error) {
    console.error(`Report reset failed: ${error.message}`);
    SpreadsheetApp.getUi().alert(`レポートのリセット中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Create report by adding data rows and setting up formulas
 * This function is called from the menu
 */
function createReport() {
  try {
    console.log("Starting report creation...");
    
    createReportSheet();
    
    console.log("Report creation completed successfully");
    
  } catch (error) {
    console.error(`Report creation failed: ${error.message}`);
    SpreadsheetApp.getUi().alert(`レポートの作成中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Export report as PDF
 * This function is called from the menu
 */
function exportReport() {
  try {
    console.log("Starting report export...");
    
    // Get export parameters
    const exportParams = getExportParameters();
    
    // Set export color to white (for PDF generation)
    setExportColor('white');
    
    // Activate the export range
    const sheet = SpreadsheetApp.getActiveSheet();
    exportParams.exportRange.activate();
    
    console.log(`Exporting ${exportParams.exportRows} rows`);
    
    // Generate PDF
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const blob = getSheetAsBlob(spreadsheet.getUrl(), sheet, exportParams.exportRange);
    
    // Create filename
    const filename = `作業報告書 ${exportParams.projectName} ${exportParams.startDate}-${exportParams.endDate}.pdf`;
    
    // Export the PDF
    exportBlob(blob, filename);
    
    // Reset export color to black
    setExportColor('black');
    
    console.log("Report export completed successfully");
    
  } catch (error) {
    console.error(`Report export failed: ${error.message}`);
    
    // Ensure color is reset even on error
    try {
      setExportColor('black');
    } catch (resetError) {
      console.error(`Failed to reset export color: ${resetError.message}`);
    }
    
    SpreadsheetApp.getUi().alert(`レポートのエクスポート中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Generate PDF blob from sheet range
 * @param {string} url - Spreadsheet URL
 * @param {Sheet} sheet - Sheet object
 * @param {Range} range - Range to export
 * @returns {Blob} PDF blob
 */
function getSheetAsBlob(url, sheet, range) {
  try {
    let rangeParam = '';
    let sheetParam = '';
    
    if (range) {
      rangeParam = 
        '&r1=' + (range.getRow() - 1) +
        '&r2=' + range.getLastRow() +
        '&c1=' + (range.getColumn() - 1) +
        '&c2=' + range.getLastColumn();
    }
    
    if (sheet) {
      sheetParam = '&gid=' + sheet.getSheetId();
    }
    
    // Build export URL with parameters
    // Credit: https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
    // Note: These parameters are reverse-engineered and may change over time
    const exportUrl = url.replace(/\/edit.*$/, '') +
      '/export?exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=false' +
      '&fitw=true' +
      '&top_margin=0.75' +
      '&bottom_margin=0.75' +
      '&left_margin=0.7' +
      '&right_margin=0.7' +
      '&sheetnames=false&printtitle=false' +
      '&pagenum=UNDEFINED' +
      '&gridlines=true' +
      '&fzr=FALSE' +
      sheetParam +
      rangeParam;
    
    console.log(`Export URL: ${exportUrl}`);
    
    // Retry logic for PDF generation
    let response;
    let retryCount = 0;
    const maxRetries = 5;
    
    while (retryCount < maxRetries) {
      response = UrlFetchApp.fetch(exportUrl, {
        muteHttpExceptions: true,
        headers: { 
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        },
      });
      
      if (response.getResponseCode() === 429) {
        // Rate limited, wait and retry
        console.log(`Rate limited, retrying in 3 seconds... (attempt ${retryCount + 1}/${maxRetries})`);
        Utilities.sleep(3000);
        retryCount++;
      } else {
        break;
      }
    }
    
    if (retryCount === maxRetries) {
      throw new Error('PDF generation failed: Too many requests. Please try again later.');
    }
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`PDF generation failed: HTTP ${response.getResponseCode()}`);
    }
    
    return response.getBlob();
    
  } catch (error) {
    console.error(`PDF generation failed: ${error.message}`);
    throw error;
  }
}

/**
 * Export blob as file and show success dialog
 * @param {Blob} blob - The PDF blob
 * @param {string} fileName - The filename for the exported file
 */
function exportBlob(blob, fileName) {
  try {
    // Set filename on blob
    blob = blob.setName(fileName);
    
    // Create file in Drive
    const folder = DriveApp;
    const pdfFile = folder.createFile(blob);
    
    // Display success dialog with link
    const htmlOutput = HtmlService
      .createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h3>エクスポート完了</h3>
          <p>PDFファイルが正常に作成されました。</p>
          <p>
            <a href="${pdfFile.getUrl()}" target="_blank" style="
              display: inline-block;
              background-color: #4285f4;
              color: white;
              padding: 10px 20px;
              text-decoration: none;
              border-radius: 4px;
              font-weight: bold;
            ">
              ${fileName} を開く
            </a>
          </p>
        </div>
      `)
      .setWidth(400)
      .setHeight(200);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'エクスポート完了');
    
    console.log(`File exported successfully: ${fileName}`);
    
  } catch (error) {
    console.error(`File export failed: ${error.message}`);
    throw error;
  }
}

/**
 * Delete tabs not in the allowed list
 * This function is called from the menu or can be run manually
 */
function deleteTabsNotInList() {
  try {
    console.log("Starting tab cleanup...");
    
    // Get allowed tabs from Instructions sheet
    const allowedTabs = getAllowedTabs();
    
    if (allowedTabs.length === 0) {
      console.warn("No allowed tabs found, skipping tab cleanup");
      return;
    }
    
    console.log(`Allowed tabs: ${allowedTabs.join(', ')}`);
    
    // Delete unallowed tabs
    deleteUnallowedTabs(allowedTabs);
    
    console.log("Tab cleanup completed successfully");
    
  } catch (error) {
    console.error(`Tab cleanup failed: ${error.message}`);
    SpreadsheetApp.getUi().alert(`タブの削除中にエラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * Complete report generation workflow
 * This function can be used to automate the entire process
 */
function generateCompleteReport() {
  try {
    console.log("Starting complete report generation workflow...");
    
    // Step 1: Update timecard data
    updateTimecards();
    
    // Step 2: Create report structure
    createReport();
    
    // Step 3: Export to PDF
    exportReport();
    
    console.log("Complete report generation workflow finished successfully");
    
  } catch (error) {
    console.error(`Complete report generation failed: ${error.message}`);
    SpreadsheetApp.getUi().alert(`レポート生成中にエラーが発生しました: ${error.message}`);
    throw error;
  }
} 