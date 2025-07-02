const instantlyApiKey = "xxx";
const mailwizzApiKey = "xxx";

const emailSheetName = "Email Management";
const emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(emailSheetName);
const serverSheetName = "Server Management";
const serverSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(serverSheetName);
const relationshipMailSheetName = "Relationship Mail Prefixes";
const relationshipMailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(relationshipMailSheetName);

var emailIndex = {};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // ui.createMenu('Management')
  //     .addSubMenu(ui.createMenu('Email Management')
  //         .addItem('Update Mailwizz & Instantly', 'postEmailManagementData'))
  //     // .addSeparator()
  //     .addToUi();

  // Email Management Menu - Push calculated quotas to external platforms
  ui.createMenu("Email Management")
    .addItem("Push Instantly Warmup Quotas", "updateInstantlyWarmupQuota")
    .addItem("Push CO Sending Quotas", "updateCoSendingQuota")
    .addItem("Push Both", "postEmailManagementData")
    .addToUi();
    
  // Server Management Menu - Calculate new quotas based on server capacity
  ui.createMenu("Server Management")
    .addItem("Build Instantly Warmup Quotas", "calculateInstantlyQuota")
    .addItem("Build CO Sending Quotas", "calculateCoQuota")
    .addItem("Build Both", "quotaCalculationsManual")
    .addToUi();
}
