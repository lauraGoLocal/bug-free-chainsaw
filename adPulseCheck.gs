// When someone makes a budget change, we record the timestamp of that change in the timestamp columns. The flag in AdPulse Check (column T) acts like a sticky note - it is added when a change has been made to the approved budget and stays there until someone removes the reminder.

// To manually run initializeStoredValues
// At top of the script editor window, you'll see a dropdown menu that currently shows "updateAdPulseCheck"  
// Click on this dropdown menu and select "initializeStoredValues" instead.
// Then click Run
// You should see a small notification in your spreadsheet saying "Stored values have been initialized"
// 	Run this function manually when:
// 		-you make major changes to many budget values at once
// 		-you notice the change detection isn't working quite right
// 		-you add many new rows


function updateAdPulseCheck(event) {
  if (!event) {
    throw new Error(
      "Please do not run the script in the script editor window. It runs automatically when you edit the spreadsheet."
    );
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Get the Properties service to store our last known values
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Get the edited row
  const editedRow = event.range.getRow();
  
  // Create a key for this row's stored value
  const storageKey = `budget_row_${editedRow}_${sheet.getName()}`;
  
  // Get the current budget value from column V
  const budgetCell = sheet.getRange(editedRow, 22); // Column V
  const currentBudgetValue = budgetCell.getValue();
  
  // Get the previously stored value for this row
  const storedValue = scriptProperties.getProperty(storageKey);
  
  // If we have a stored value, compare it with the current value
  if (storedValue !== null) {
    if (String(storedValue) !== String(currentBudgetValue)) {
      // The budget changed - set the flag in column T
      const flagCell = sheet.getRange(editedRow, 20);
      flagCell.setValue(false);
      
      // Log the change for debugging
      console.log(`Budget changed in row ${editedRow}`);
      console.log(`Old value: ${storedValue}`);
      console.log(`New value: ${currentBudgetValue}`);
    }
  }
  
  // Store the current value for future comparison
  scriptProperties.setProperty(storageKey, String(currentBudgetValue));
}

/**
 * Use this function to initialize or reset stored values
 * Run this manually when you want to start fresh
 */
function initializeStoredValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Clear all existing stored values
  scriptProperties.deleteAllProperties();
  
  // Store current values for each row
  for (let row = 2; row <= lastRow; row++) {
    const budgetValue = sheet.getRange(row, 22).getValue();
    const storageKey = `budget_row_${row}_${sheet.getName()}`;
    scriptProperties.setProperty(storageKey, String(budgetValue));
  }
  
  SpreadsheetApp.getActive().toast(
    'Stored values have been initialized',
    'Setup Complete',
    5
  );
}
