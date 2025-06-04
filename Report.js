function resetReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();

  var a14 = sheet.getRange('B14').getValue();
  if(a14 == "稼働時間合計") {
    SpreadsheetApp.getUi().alert('シートはすでにリセット済みです');
    return;
  }

  var firstRow = 13;
  var rows = sheet.getRange('A1').getValue();
  var lastRow = 13 + rows;
  sheet.getRange("14:"+lastRow).activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());

}

function createReport() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getRange('D22').getValue();
  if(values == "" || values > 1000) {
    SpreadsheetApp.getUi().alert('セルD22に追加したい行数が入るようにしてください。\nSampleタブのように13行と合計の行だけが残るようにすればD22を行数にすることができます。\n不安な場合はSampleをDuplicateしてからマクロを動かすことをお勧めします。');
    return;
  }

  // create data rows based on how many records are retrieved from PSA
  console.log("Inserting "+values+" rows after row 13");
  sheet.insertRowsAfter(13,values);

  // copy-paste the formulas in B13:M13 into the newly created rows
  // B13:M13 should already contain formulas that point to the PSA data (eg. B13=A26, C13=B26, etc.)
  sheet.getRange('B14:B15').activate();
  var currentCell = sheet.getCurrentCell();
  sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  sheet.getCurrentCell().offset(0, 0, values, 1).activate();
  sheet.getRange('B13:M13').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  //set sum formula
  sheet.getCurrentCell().offset(values, 4).activate();
  var sum_values = values + 1;
  sheet.getCurrentCell().setFormulaR1C1("=SUM(R[-"+sum_values+"]C[0]:R[-1]C[1])");
}

function exportReport() {
  // set active range
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var rows = 20 + sheet.getRange('A1').getValue();
  sheet.getRange('N1').setFontColor('white');
  sheet.getRange('A1:N'+rows).activate();
  console.log("Rows to be exported "+rows);

  var projectName = sheet.getRange('E10').getValue();
  var startDate = sheet.getRange('A2').getValue();
  var endDate = sheet.getRange('A3').getValue();

  var blob = getSheetAsBlob(spreadsheet.getUrl(),sheet,sheet.getActiveRange());
  var filename = "作業報告書 "+projectName+' '+startDate+'-'+endDate+'.pdf';
  exportBlob(blob, filename);
  sheet.getRange('N1').setFontColor('black');
}

function getSheetAsBlob(url, sheet, range) {
  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=A4'
      + '&portrait=false'
      + '&fitw=true'       
      + '&top_margin=0.75'              
      + '&bottom_margin=0.75'          
      + '&left_margin=0.7'             
      + '&right_margin=0.7'           
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=true'
      + '&fzr=FALSE'      
      + sheetParam
      + rangeParam
      
  Logger.log('exportUrl=' + exportUrl)
  var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }
  
  return response.getBlob()
}

function exportBlob(blob, fileName) {
  blob = blob.setName(fileName)
  var folder = DriveApp
  var pdfFile = folder.createFile(blob)
  
  // Display a modal dialog box with custom HtmlService content.
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
    .setHeight(120)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
}

function deleteTabsNotInList() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the list data from the specified range
  var listSheet = spreadsheet.getSheetByName("Instructions");
  var listRange = listSheet.getRange("F2:F10");
  var listValues = listRange.getValues();

  // Flatten the list (assuming single column)
  var tabList = listValues.map(function(row) { return row[0]; }); // Adjust if list is in multiple columns
  
  // Convert the list to lowercase for case-insensitive comparison (optional)
  var listLower = tabList.map(function(name) {
    return name.toLowerCase();
  });

  // Get all sheets in the spreadsheet
  var sheets = spreadsheet.getSheets();
  
  // Loop through sheets in reverse
  for (var i = sheets.length - 1; i >= 0; i--) {
    var sheetNameLower = sheets[i].getName().toLowerCase();
    var isHidden = sheets[i].isSheetHidden();
    var protections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);

    if (listLower.indexOf(sheetNameLower) === -1 && (isHidden || protections.length === 0)) {
      spreadsheet.deleteSheet(sheets[i]);
    }
  }
}
