const TIMECARD_FIELDS = [
  "pse__Project__r.pse__Opportunity__r.OpportunityNumber__c",
  "pse__Project__r.OPA_Project_Number__c",
  //"pse__Project__r.pse__Planned_Hours__c",
  "pse__Project__r.pse__Start_Date__c",
  "pse__Project__r.pse__End_Date__c",
  "pse__Project__r.Name",
  //"pse__Project__r.pse__Total_Assigned_Hours__c",
  "pse__Milestone__r.OPA_Task_Number__c",
  "pse__Milestone__r.Name",
  //"pse__Milestone__r.Item_Name_c",
  "pse__Resource__r.Name",
  "pse__Start_Date__c",
  "pse__Total_Hours__c",
  //"pse__Timecard_Notes__c",
  "pse__Sunday_Notes__c",
  "pse__Monday_Notes__c",
  "pse__Tuesday_Notes__c",
  "pse__Wednesday_Notes__c",
  "pse__Thursday_Notes__c",
  "pse__Friday_Notes__c",
  "pse__Saturday_Notes__c",
  "pse__Submitted__c",
  //"pse__Assignment__r.pse__Planned_Hours__c",
];
const TIMECARD_API_NAME = "pse__Timecard__c";
const TIMECARD_FILTER = [
  "pse__Status__c+NOT+IN+('Rejected')",
  "pse__Total_Hours__c+%21%3D+0",
  //"pse__Start_Date__c+%3E%3D+LAST_YEAR",
  //"pse__Project__r.pse__Region__r.Name+%3D+'Japan_1'",
  //"pse__Project__r.pse__Stage__c+IN+('Completed','In%20Progress','On%20Hold','Planned')",
  //"pse__Milestone__r.Item_Name__c+NOT+IN+('Non-Billable%20Expense','Non-Billable%20Labor','Travel%20Time','Covered%20by%20Margin')"
  "(NOT+pse__Milestone__r.Name+LIKE+'NB%25')"
];

const PROJECT_FIELDS = [
  "pse__Opportunity__r.OpportunityNumber__c",
  "Name",
  "pse__Account__c",
  "pse__Account__r.Name",
  "Region_Level_2__c",
  "pse__Stage__c",
  "pse__Start_Date__c",
  "pse__End_Date__c",
  "OPA_Project_Number__c",
];
const PROJECT_API_NAME = "pse__Proj__c";
const PROJECT_FILTER = [
  //"Region_Level_2__c+%3D+'Japan'",
  //"pse__Stage__c+IN+('In%20Progress','Planned','Completed')",
  //"pse__End_Date__c+%3D+THIS_YEAR"
];
const PROJECT_SHEET_NAME = "PROJECT_DETAILS";