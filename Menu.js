/**
 * Menu - Creates custom menu for the timecard reporting system
 * Provides user-friendly interface for all report generation functions
 */

/**
 * Creates custom menu when spreadsheet is opened
 * This function is automatically called when the spreadsheet is opened
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Create main menu
    const menu = ui.createMenu('📋 レポート作成');

    // Report generation section
    menu.addSubMenu(ui.createMenu('📄 レポート操作')
      .addItem('完全自動生成', 'generateCompleteReport')
      .addSeparator()
      .addItem('リセット', 'resetReport')
      .addItem('作成', 'createReport')
      .addItem('PDF エクスポート', 'exportReport')
    );

    menu.addSeparator();

    // Data update section
    menu.addSubMenu(ui.createMenu('📊 データ更新')
      .addItem('全データ更新', 'updateAllData')
      .addSeparator()
      .addItem('タイムカード更新', 'updateTimecards')
      .addItem('プロジェクト更新', 'updateProjectsData')
    );
    
    menu.addSeparator();
    
    // Utility functions
    menu.addSubMenu(ui.createMenu('🔧 ユーティリティ')
      //.addItem('不要タブ削除', 'deleteTabsNotInList')
      .addItem('システム情報', 'showSystemInfo')
    );
    
    // Add menu to spreadsheet
    menu.addToUi();
    
    console.log("Custom menu created successfully");
    
  } catch (error) {
    console.error(`Error creating menu: ${error.message}`);
    // Don't show UI alert here as it might interfere with spreadsheet opening
  }
}

/**
 * Update all data (timecards and projects)
 * Convenience function for updating both timecard and project data
 */
function updateAllData() {
  try {
    console.log("Starting full data update...");
    
    // Show progress notification
    const ui = SpreadsheetApp.getUi();
    const progressDialog = ui.alert(
      'データ更新中',
      'タイムカードとプロジェクトデータを更新しています...\n\n処理が完了するまでお待ちください。',
      ui.ButtonSet.OK
    );
    
    // Update timecards (this will also update projects)
    updateTimecards();
    
    // Show completion message
    ui.alert(
      '更新完了',
      'すべてのデータが正常に更新されました。',
      ui.ButtonSet.OK
    );
    
    console.log("Full data update completed successfully");
    
  } catch (error) {
    console.error(`Full data update failed: ${error.message}`);
    SpreadsheetApp.getUi().alert(`データ更新中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * Show system information and current configuration
 * Useful for debugging and support
 */
function showSystemInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Get current parameters
    let params;
    try {
      params = getInputParameters();
    } catch (error) {
      params = { error: error.message };
    }
    
    // Build info HTML
    const infoHtml = `
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>システム情報</h3>
        
        <h4>スプレッドシート情報</h4>
        <ul>
          <li><strong>スプレッドシート名:</strong> ${spreadsheet.getName()}</li>
          <li><strong>アクティブシート:</strong> ${sheet.getName()}</li>
          <li><strong>スプレッドシート ID:</strong> ${spreadsheet.getId()}</li>
        </ul>
        
        <h4>現在のパラメータ</h4>
        ${params.error ? 
          `<p style="color: red;">エラー: ${params.error}</p>` :
          `<ul>
            <li><strong>オポチュニティ番号:</strong> ${params.opportunityNumber}</li>
            <li><strong>開始日:</strong> ${params.startDate.toDateString()}</li>
            <li><strong>終了日:</strong> ${params.endDate.toDateString()}</li>
            <li><strong>追加行数:</strong> ${params.rowsToAdd}</li>
          </ul>`
        }
        
        <h4>設定情報</h4>
        <ul>
          <li><strong>タイムカード API:</strong> ${TIMECARD_API_NAME}</li>
          <li><strong>プロジェクト API:</strong> ${PROJECT_API_NAME}</li>
          <li><strong>フィールド数:</strong> ${TIMECARD_FIELDS.length} (タイムカード), ${PROJECT_FIELDS.length} (プロジェクト)</li>
        </ul>
        
        <h4>サポート情報</h4>
        <p>問題が発生した場合は、この情報をサポートチームに提供してください。</p>
        <p><small>Generated at: ${new Date().toLocaleString()}</small></p>
      </div>
    `;
    
    const htmlOutput = HtmlService
      .createHtmlOutput(infoHtml)
      .setWidth(500)
      .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'システム情報');
    
  } catch (error) {
    console.error(`Error showing system info: ${error.message}`);
    SpreadsheetApp.getUi().alert(`システム情報の表示中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * Show confirmation dialog for potentially destructive operations
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 * @returns {boolean} True if user confirmed, false otherwise
 */
function showConfirmationDialog(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      title,
      message,
      ui.ButtonSet.YES_NO
    );
    
    return result === ui.Button.YES;
    
  } catch (error) {
    console.error(`Error showing confirmation dialog: ${error.message}`);
    return false;
  }
}

/**
 * Enhanced reset function with confirmation
 * Wrapper around resetReport() with user confirmation
 */
function resetReportWithConfirmation() {
  try {
    const confirmed = showConfirmationDialog(
      'レポートリセット確認',
      'レポートをリセットしますか？\n\n現在のデータ行がすべて削除されます。\nこの操作は取り消せません。'
    );
    
    if (confirmed) {
      resetReport();
    } else {
      console.log("Report reset cancelled by user");
    }
    
  } catch (error) {
    console.error(`Error in reset confirmation: ${error.message}`);
  }
}

/**
 * Enhanced tab deletion function with confirmation
 * Wrapper around deleteTabsNotInList() with user confirmation
 */
function deleteTabsWithConfirmation() {
  try {
    const confirmed = showConfirmationDialog(
      'タブ削除確認',
      '許可されていないタブを削除しますか？\n\n削除されたタブは復元できません。'
    );
    
    if (confirmed) {
      deleteTabsNotInList();
    } else {
      console.log("Tab deletion cancelled by user");
    }
    
  } catch (error) {
    console.error(`Error in tab deletion confirmation: ${error.message}`);
  }
}