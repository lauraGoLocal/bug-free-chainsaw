// This script contains the LASTMONTHBUDGET function, which plays a crucial role in determining 
// the budget amounts that get exported to AdPulse.
//
// The LASTMONTHBUDGET Function
// This function automatically determines what budget amount should have been added to AdPulse, regardless of any errors in record keeping. 
//
// How LASTMONTHBUDGET finds the right budget amount:
// It looks at last month's budget sheet and follows these rules in order:
//   1. If "No Ads This Month" was checked last month → Returns 0
//   2. If the "Recommended Budget" was approved → Returns that budget amount
//   3. If there was a pre-approved budget → Returns that amount
//   4. If there was a previously approved budget → Returns that amount
//   5. If none of the above → Returns "#N/A"
//
// Example: If you're in April's budget sheet, =LASTMONTHBUDGET("Campaign Name")
// would look in March's sheet for that campaign's approved budget amount.
//


function LASTMONTHBUDGET(sheetName, matchValue) {
  if (!sheetName) {
    sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    console.log("No sheet name provided. Using active sheet: " + sheetName);
  }

  console.log("Looking for campaign: '" + matchValue + "'"); // Added logging for campaign name

  console.log("Attempting to find sheet: " + sheetName);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getSheetByName(sheetName);

  if (!currentSheet) {
    console.error("Sheet not found: " + sheetName);
    return "#SHEET_NOT_FOUND";
  }

  var currentDate = new Date(sheetName.split(" ")[0] + " 1, " + sheetName.split(" ")[1]);
  console.log("Current date: " + currentDate);

  var prevDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
  var prevSheetName = Utilities.formatDate(prevDate, Session.getScriptTimeZone(), "MMMM yyyy") + " Budgets";

  var prevSheet = ss.getSheetByName(prevSheetName);
  if (!prevSheet) return "#N/A";

  console.log("Previous sheet name: " + prevSheetName);

  var prevValues = prevSheet.getRange("C:K").getValues();

  // Log the first few campaign names from previous sheet for debugging
  console.log("First few campaigns in previous sheet:");
  for (var i = 0; i < Math.min(5, prevValues.length); i++) {
    if (prevValues[i][0]) {
      console.log(`'${prevValues[i][0]}'`);
    }
  }

  var prevRow = prevValues.findIndex((row) => row[0] === matchValue);
  console.log("Found row index: " + prevRow); // Added logging for row index

  if (prevRow === -1) {
    console.log("No matching campaign found");
    return "#N/A";
  }

  var prevI = prevValues[prevRow][6]; // Column I is index 6
  var prevJ = prevValues[prevRow][7]; // Column J is index 7
  var prevH = prevValues[prevRow][5]; // Column H is index 5
  var prevG = prevValues[prevRow][4]; // Column G is index 4
  var prevK = prevValues[prevRow][8]; // Column K is index 8

  console.log("Previous month values found:");
  console.log("I (No Ads): " + prevI);
  console.log("H (Recommended Approved): " + prevH);
  console.log("G (Budget Total): " + prevG);
  console.log("J (Pre-Approved): " + prevJ);
  console.log("K (Last Month Approved): " + prevK);

  if (prevI === true) return 0; // No ads this month
  if (prevH === true) return Math.round(prevG); // Recommended budget approved
  if (typeof prevJ === "number" && prevJ > 0) return Math.round(prevJ); // Pre-approved budget
  if (typeof prevK === "number" && prevK > 0) return Math.round(prevK); // Last month's approved budget

  return "#N/A";
}

function CURRENT_SHEET_NAME() {
  return SpreadsheetApp.getActiveSheet().getName();
}
