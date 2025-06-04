function updateProjects() {
  const queryString = buildProjectQuery_();
  console.log(queryString);
  let results = RHCCRestApi.query(queryString);
  const records = [];

  if(results && results.records) {
    records.push(...results.records);
  }
  
  // Handle multiple batches if more than 2000 results are returned
  if(results && results.done === false) {
    do {
      results = RHCCRestApi.queryMore(results.nextRecordsUrl);
      records.push(...results.records);
    } while(results.done === false)
  }
  console.log(records);

  updateProjectSheetWithResults_(records);
}

function buildProjectQuery_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getActiveSheet();

  const oppNumber = templateSheet.getRange("B21").getValue();
  const queryFields = PROJECT_FIELDS.join(",");
  //var queryFilter = PROJECT_FILTER.join("+AND+");
  var queryFilter = "pse__Opportunity__r.OpportunityNumber__c%3D" + "'"+oppNumber+"'"
  return `SELECT+${queryFields}+FROM+${PROJECT_API_NAME}+WHERE+${queryFilter}`; //+AND+pse__Project__r.pse__Opportunity__r.OpportunityNumber__c%3D'${oppNumber}'`;
}

function updateProjectSheetWithResults_(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectSheet = ss.getActiveSheet();

  projectSheet.getRange('D21:G21').clearContent();

  // Build an array from the returned records that can be used to update the sheet
  const sheetResults = [];
  records.forEach( (record) => {
    const oppNumber = record.pse__Opportunity__r ? record.pse__Opportunity__r.OpportunityNumber__c : "";  // set null when records do not have linked opportunities
    const accountName = record.pse__Account__c ? record.pse__Account__r.Name : "";  // set null when records do not have linked accounts
    sheetResults.push([
      //oppNumber,
      record.Name,
      accountName,
      //record.Region_Level_2__c,
      //record.pse__Stage__c,
      //record.pse__Start_Date__c,
      //record.pse__End_Date__c,
      record.OPA_Project_Number__c
    ]);
  });

  // Write timecard data to the sheet
  if(sheetResults.length > 0) {
    //timecardSheet.getRange(22, 4, sheetResults.length, sheetResults[0].length).setValues(sheetResults);
    projectSheet.getRange(21, 4, 1, sheetResults[0].length).setValues([
      sheetResults[0],
    ]);
  }  
}