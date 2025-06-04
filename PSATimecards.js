function updateTimecards() {
  const queryString = buildQuery_();
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

  updateTimecardSheetWithResults_(records);

  updateProjects();
}

function formatDate(in_date) {
  var out_date = new Date(in_date);
  var year = out_date.getFullYear();
  var month = ('00' + (out_date.getMonth()+1)).slice(-2);
  var day = ('00' + out_date.getDate()).slice(-2);
  return (year + '-' + month + '-' + day);
}

function buildQuery_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timecardSheet = ss.getActiveSheet();

  const oppNumber = timecardSheet.getRange("B21").getValue();
  var startDate = timecardSheet.getRange("B22").getValue(); startDate = formatDate(startDate);
  var endDate = timecardSheet.getRange("B23").getValue();  endDate = formatDate(endDate);
  const queryFields = TIMECARD_FIELDS.join(",");
  var queryFilter = TIMECARD_FILTER.join("+AND+");
  queryFilter = queryFilter + "+AND+pse__Project__r.pse__Opportunity__r.OpportunityNumber__c%3D" + "'"+oppNumber+"'"
//+ "+AND+pse__Start_Date__c+%3E%3D+" + "2024-12-01";
                            + "+AND+pse__Start_Date__c%3E%3D" +startDate
                            + "+AND+pse__Start_Date__c%3C%3D" +endDate;
  return `SELECT+${queryFields}+FROM+${TIMECARD_API_NAME}+WHERE+${queryFilter}+ORDER+BY+pse__Start_Date__c`
}

function updateTimecardSheetWithResults_(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timecardSheet = ss.getActiveSheet();
  //const timecardSheet = ss.getSheetByName("6615218");

  timecardSheet.getRange('A26:I').clearContent();

  // Build an array from the returned records that can be used to update the sheet
  const sheetResults = [];
  records.forEach( (record) => {

    sheetResults.push([
      record.pse__Resource__r.Name,
      
      record.pse__Start_Date__c,
      record.pse__Total_Hours__c,
      //record.pse__Sunday_Notes__c,
      record.pse__Monday_Notes__c,
      record.pse__Tuesday_Notes__c,
      record.pse__Wednesday_Notes__c,
      record.pse__Thursday_Notes__c,
      record.pse__Friday_Notes__c,
      //record.pse__Saturday_Notes__c,
      record.pse__Milestone__r.Name,
    ]);
  });

  // Write timecard data to the sheet
  if(sheetResults.length > 1) {
    timecardSheet.getRange(26, 1, sheetResults.length, sheetResults[0].length).setValues(sheetResults);
  }  
}