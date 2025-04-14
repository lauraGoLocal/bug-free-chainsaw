// This script automatically adds timestamps whenever certain columns are edited in budget sheets
//
// What does it do?
// - Watches for changes in specific columns like "Recommended Budget Approved" and "No Ads This Month"
// - When these columns are edited, it automatically adds a timestamp in a corresponding column
// - For example, when someone marks "Recommended Budget Approved" as checked, it adds the current 
//   date and time in the "Approval Timestamp" column
//
// When does it run?
// - The script runs automatically whenever someone edits the spreadsheet
// - It only works on sheets that end with "24 Budgets" or "25 Budgets"
// - It only responds to changes in specific columns that we're monitoring
//
// What exactly does it track?
// 1. When "Recommended Budget Approved" is checked → Updates "Approval Timestamp"
// 2. When "No Ads This Month" is checked → Updates "Disapproval Timestamp"
// 3. When "Pre-Approved Client Budgets" is edited → Updates "Pre-approval Timestamp"
// 4. When "Budget Total" is edited → Updates "Budget Total Timestamp"
//    (This timestamp will be cleared if the Budget Total is cleared)

// MASTER onEdit function - this is the main entry point that handles all edit events
function onEdit(event) {
  if (!event) {
    throw new Error(
      "Please do not run the script in the script editor window. It runs automatically when you edit the spreadsheet."
    );
  }

  // Call the original timestamp tracking function
  try {
    processSheets(event, [
      {
        description: 'Sheets that end with "25 Budgets" or "26 Budgets"',
        sheetsToWatch: /^.*2[56] Budgets$/i,
        columnConfigs: [
          {
            description: "Update 'Approval Timestamp' when 'Recommended Budget Approved' is edited.",
            labelToWatch: "Recommended Budget Approved",
            labelToStamp: "Approval Timestamp",
            valueToWatch: true,
          },
          {
            description: "Update 'Disapproval Timestamp' when 'No Ads This Month' is edited.",
            labelToWatch: "No Ads This Month",
            labelToStamp: "Disapproval Timestamp",
            valueToWatch: true,
          },
          {
            description: "Update 'Pre-Approval Timestamp' when 'Pre-Approval Client Budgets' is edited.",
            labelToWatch: "Pre-Approved Client Budgets",
            labelToStamp: "Pre-Approval Timestamp",
          },
          {
            description: "Update 'Budget Total Timestamp' when 'Budget Total' is edited. Clear when cleared.",
            labelToWatch: "Budget Total",
            labelToStamp: "Budget Total Timestamp",
            eraseTimestamp: true,
          },
        ],
      },
    ]);
  } catch (e) {
    console.error("Error in processSheets:", e);
  }

  // Call the AdPulse Check function
  try {
    updateAdPulseCheck(event);
  } catch (e) {
    console.error("Error in updateAdPulseCheck:", e);
  }
  
  // Call the AdPulse Budget Verification edit handler
  try {
    handleAdPulseEdit(event);
  } catch (e) {
    console.error("Error in handleAdPulseEdit:", e);
  }
}
