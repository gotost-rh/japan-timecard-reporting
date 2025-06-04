function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('レポート作成')
      .addItem('Reset Report', 'resetReport')
      .addItem('Create Report', 'createReport')
      .addItem('Export Report (PDF)', 'exportReport')
      .addToUi();
}