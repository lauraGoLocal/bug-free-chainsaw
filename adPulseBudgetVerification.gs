// AdPulse Budget Verification Script
//
// What does this script do?
// - Automatically verifies that budget amounts match the corresponding amounts in AdPulse
// - Applies conditional formatting to highlight mismatches
// - Adds a status column that shows matches and mismatches
//
// How it works:
// - Paste data from AdPulse into the "AdPulse" sheet (any columns, as long as it has "Budget Name" and "Budget Target")
// - The script runs automatically when the AdPulse sheet is updated
// - Mismatches will be highlighted in your budget sheets

// Create the AdPulse sheet if it doesn't exist when the spreadsheet is opened
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('AdPulse')) {
    createAdPulseSheet();
  }
}

/**
 * Creates the AdPulse sheet with simple instructions
 */
function createAdPulseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adPulseSheet = ss.insertSheet('AdPulse');
  
  // Add simple instructions
  adPulseSheet.getRange('A1').setValue('Paste AdPulse data below (with headers including "Budget Name" and "Budget Target")');
  adPulseSheet.getRange('A1').setFontWeight('bold');
  adPulseSheet.getRange('A1:F1').merge();
  
  // Format the sheet for readability
  adPulseSheet.setColumnWidth(1, 300);
  
  SpreadsheetApp.getActive().toast(
    'AdPulse sheet created. Paste your data including column headers "Budget Name" and "Budget Target".',
    'Setup Complete',
    5
  );
}

/**
 * Trigger function that runs when a spreadsheet is edited
 */
function onEdit(e) {
  // Verify we have a valid edit event
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // Only run if editing the AdPulse sheet
  if (sheetName === 'AdPulse') {
    // Run verification after a short delay to allow for paste operations to complete
    SpreadsheetApp.getActive().toast(
      'AdPulse data updated. Verification will run automatically.',
      'AdPulse Verification',
      3
    );
    
    // We can't use setTimeout in Apps Script directly, so we'll use a trigger
    // Delete any existing triggers for delayed verification
    const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'delayedVerification') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Create a new time-based trigger that runs after a short delay
    ScriptApp.newTrigger('delayedVerification')
      .timeBased()
      .after(5000) // 5 seconds
      .create();
  }
}

/**
 * Function that runs after a delay to perform verification
 */
function delayedVerification() {
  // Run the verification
  runBudgetVerification();
  
  // Clean up the trigger
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'delayedVerification') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Main function to run the budget verification
 */
function runBudgetVerification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if AdPulse sheet exists
  const adPulseSheet = ss.getSheetByName('AdPulse');
  if (!adPulseSheet) {
    createAdPulseSheet();
    return;
  }
  
  // Get all budget sheets (ending with "25 Budgets" or "26 Budgets")
  const sheets = ss.getSheets();
  const budgetSheets = sheets.filter(sheet => {
    const name = sheet.getName();
    return /^.*2[56] Budgets$/i.test(name);
  });
  
  if (budgetSheets.length === 0) {
    SpreadsheetApp.getActive().toast(
      'No budget sheets found (sheets ending with "25 Budgets" or "26 Budgets").',
      'AdPulse Verification',
      5
    );
    return;
  }
  
  // Show a loading message
  const toast = SpreadsheetApp.getActive().toast(
    'Running AdPulse budget verification...',
    'Please wait',
    -1
  );
  
  try {
    // First, get the AdPulse data
    const adPulseData = getAdPulseData(adPulseSheet);
    if (!adPulseData) {
      toast.dismiss();
      return; // Error already shown by getAdPulseData
    }
    
    // Now process each budget sheet
    let totalMatchCount = 0;
    let totalMismatchCount = 0;
    let totalNotFoundCount = 0;
    
    for (const budgetSheet of budgetSheets) {
      const result = verifyBudgetSheet(budgetSheet, adPulseData);
      totalMatchCount += result.matchCount;
      totalMismatchCount += result.mismatchCount;
      totalNotFoundCount += result.notFoundCount;
    }
    
    // Show completion message
    toast.dismiss();
    SpreadsheetApp.getActive().toast(
      `Verification complete: ${totalMatchCount} matches, ${totalMismatchCount} mismatches, ${totalNotFoundCount} not found.`,
      'AdPulse Verification',
      5
    );
    
  } catch (error) {
    toast.dismiss();
    SpreadsheetApp.getActive().toast(
      'Error during verification: ' + error.toString(),
      'Error',
      5
    );
    console.error('Verification error: ' + error.toString());
  }
}

/**
 * Gets budget data from the AdPulse sheet
 * @param {Sheet} adPulseSheet - The AdPulse sheet
 * @return {Object|null} - Map of campaign names to budget amounts, or null on error
 */
function getAdPulseData(adPulseSheet) {
  // Get all data including headers
  const data = adPulseSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getActive().toast(
      'No data found in AdPulse sheet. Please paste data including headers.',
      'AdPulse Verification',
      5
    );
    return null;
  }
  
  // Find the column indices for Budget Name and Budget Target
  const headers = data[0];
  let nameColIndex = -1;
  let budgetColIndex = -1;
  
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] && headers[i].toString().trim();
    if (header === 'Budget Name') {
      nameColIndex = i;
    } else if (header === 'Budget Target') {
      budgetColIndex = i;
    }
  }
  
  if (nameColIndex === -1 || budgetColIndex === -1) {
    SpreadsheetApp.getActive().toast(
      'Required columns "Budget Name" and "Budget Target" not found in AdPulse sheet.',
      'AdPulse Verification',
      5
    );
    return null;
  }
  
  // Create a map of campaign names to budget amounts
  // Store both original name and normalized (lowercase) name
  const adPulseBudgets = {
    originalNames: {}, // Maps lowercase names to original names
    values: {}         // Maps lowercase names to budget values
  };
  
  for (let i = 1; i < data.length; i++) {
    const campaignName = data[i][nameColIndex];
    const budget = data[i][budgetColIndex];
    
    if (campaignName && budget !== null && budget !== undefined && budget !== '') {
      // Try to convert budget to a number
      const budgetValue = typeof budget === 'number' ? budget : parseFloat(budget.toString().replace(/[^0-9.-]+/g, ''));
      
      if (!isNaN(budgetValue)) {
        // Store with lowercase key for case-insensitive matching
        const lowerCaseName = campaignName.toString().toLowerCase();
        adPulseBudgets.values[lowerCaseName] = budgetValue;
        adPulseBudgets.originalNames[lowerCaseName] = campaignName; // Keep original name for display
      }
    }
  }
  
  return adPulseBudgets;
}

/**
 * Verifies a single budget sheet against AdPulse data
 * @param {Sheet} budgetSheet - The budget sheet to verify
 * @param {Object} adPulseData - Map of campaign names to budget amounts
 * @return {Object} - Counts of matches, mismatches, and not found
 */
function verifyBudgetSheet(budgetSheet, adPulseData) {
  // Get all data including headers
  const data = budgetSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { matchCount: 0, mismatchCount: 0, notFoundCount: 0 };
  }
  
  // Get headers from the first row
  const headers = data[0];
  
  // Find the columns for Campaign Name and Budget Target
  // In budget sheets, we specifically use columns U and V
  const campaignColIndex = 20; // Column U is index 20 (0-based)
  const budgetColIndex = 21;   // Column V is index 21 (0-based)
  
  // Find or create the status column (column W, right after the budget column)
  const statusColIndex = 22; // Column W is index 22 (0-based)
  
  // Check if the header for column W is set to "AdPulse Match"
  if (!headers[statusColIndex] || headers[statusColIndex].toString().trim() !== 'AdPulse Match') {
    budgetSheet.getRange(1, statusColIndex + 1).setValue('AdPulse Match');
  }
  
  // Clear existing status values in column W
  if (data.length > 1) {
    budgetSheet.getRange(2, statusColIndex + 1, data.length - 1, 1).clearContent();
  }
  
  // Verify budgets and set status
  let matchCount = 0;
  let mismatchCount = 0;
  let notFoundCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const campaignName = data[i][campaignColIndex];
    const sheetBudget = data[i][budgetColIndex];
    
    // Skip empty campaign names or budget rows
    if (!campaignName || sheetBudget === null || sheetBudget === undefined || sheetBudget === '') {
      continue;
    }
    
    // Convert to lowercase for case-insensitive matching
    const lowerCaseName = campaignName.toString().toLowerCase();
    
    // Check if campaign exists in AdPulse data (case-insensitive)
    if (adPulseData.values.hasOwnProperty(lowerCaseName)) {
      const adPulseBudget = adPulseData.values[lowerCaseName];
      
      // Clean and normalize both budget values for comparison
      let sheetBudgetNum = cleanBudgetValue(sheetBudget);
      let adPulseBudgetNum = cleanBudgetValue(adPulseBudget);
      
      // Round both values to nearest dollar for comparison
      sheetBudgetNum = Math.round(sheetBudgetNum);
      adPulseBudgetNum = Math.round(adPulseBudgetNum);
      
      // Compare budgets - exact match after rounding to nearest dollar
      if (!isNaN(sheetBudgetNum) && !isNaN(adPulseBudgetNum) && sheetBudgetNum === adPulseBudgetNum) {
        budgetSheet.getRange(i + 1, statusColIndex + 1).setValue('✓ Match');
        matchCount++;
      } else {
        // Get original matched name from AdPulse for display purposes
        const originalAdPulseName = adPulseData.originalNames[lowerCaseName];
        
        // Show original name if it differs from sheet name (ignoring case)
        let mismatchMessage = '❌ Mismatch: $' + adPulseBudgetNum;
        if (originalAdPulseName !== campaignName) {
          mismatchMessage += ' (matched to "' + originalAdPulseName + '")';
        }
        
        budgetSheet.getRange(i + 1, statusColIndex + 1).setValue(mismatchMessage);
        mismatchCount++;
      }
    } else {
      budgetSheet.getRange(i + 1, statusColIndex + 1).setValue('⚠️ Not Found');
      notFoundCount++;
    }
  }
  
  // Apply conditional formatting
  applyConditionalFormatting(budgetSheet, statusColIndex + 1);
  
  return { matchCount, mismatchCount, notFoundCount };
}

/**
 * Helper function to clean and normalize budget values
 * @param {any} value - The budget value to clean
 * @return {number} The cleaned numeric value
 */
function cleanBudgetValue(value) {
  if (typeof value === 'number') {
    return value;
  }
  
  if (!value) return NaN;
  
  // Convert to string, remove currency symbols and commas, then parse to number
  let cleanedValue = value.toString()
    .replace(/[$,]/g, '') // Remove $ and commas
    .trim();
  
  return parseFloat(cleanedValue);
}

/**
 * Helper function to clean and normalize budget values
 * @param {any} value - The budget value to clean
 * @return {number} The cleaned numeric value
 */
function cleanBudgetValue(value) {
  if (typeof value === 'number') {
    return value;
  }
  
  if (!value) return NaN;
  
  // Convert to string, remove currency symbols and commas, then parse to number
  let cleanedValue = value.toString()
    .replace(/[$,]/g, '') // Remove $ and commas
    .trim();
  
  return parseFloat(cleanedValue);
}

/**
 * Applies conditional formatting to highlight matches, mismatches, and not found items
 * @param {Sheet} sheet - The sheet to apply formatting to
 * @param {number} statusColIndex - The column index of the status column (1-based)
 */
function applyConditionalFormatting(sheet, statusColIndex) {
  // Clear existing conditional formatting rules for the status column
  let rules = sheet.getConditionalFormatRules();
  const newRules = [];
  
  // Keep only rules that don't apply to our status column
  for (let i = 0; i < rules.length; i++) {
    const ranges = rules[i].getRanges();
    let keepRule = true;
    
    for (let j = 0; j < ranges.length; j++) {
      if (ranges[j].getColumn() === statusColIndex) {
        keepRule = false;
        break;
      }
    }
    
    if (keepRule) {
      newRules.push(rules[i]);
    }
  }
  
  // Add new rules for matches
  const matchRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('✓ Match')
    .setBackground('#e6f4ea') // Light green
    .setRanges([sheet.getRange(2, statusColIndex, sheet.getLastRow() - 1, 1)])
    .build();
  newRules.push(matchRule);
  
  // Add new rules for mismatches
  const mismatchRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('❌ Mismatch')
    .setBackground('#fce8e6') // Light red
    .setFontColor('#ea4335') // Red text
    .setRanges([sheet.getRange(2, statusColIndex, sheet.getLastRow() - 1, 1)])
    .build();
  newRules.push(mismatchRule);
  
  // Add new rules for not found
  const notFoundRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('⚠️ Not Found')
    .setBackground('#fff8e1') // Light yellow
    .setFontColor('#f09300') // Orange text
    .setRanges([sheet.getRange(2, statusColIndex, sheet.getLastRow() - 1, 1)])
    .build();
  newRules.push(notFoundRule);
  
  // Set the new rules
  sheet.setConditionalFormatRules(newRules);
}

/**
 * Convenient function for testing the verification on the current sheet
 */
function testVerifyCurrentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the active sheet (the one you're currently viewing)
  const currentSheet = SpreadsheetApp.getActiveSheet();
  
  // Get AdPulse data
  const adPulseSheet = ss.getSheetByName('AdPulse');
  if (!adPulseSheet) {
    SpreadsheetApp.getActive().toast(
      'AdPulse sheet not found. Create it first by running "createAdPulseSheet"',
      'Error',
      5
    );
    return;
  }
  
  const adPulseData = getAdPulseData(adPulseSheet);
  if (!adPulseData) {
    // Error message already shown by getAdPulseData
    return;
  }
  
  // Now verify the current sheet
  const result = verifyBudgetSheet(currentSheet, adPulseData);
  
  // Show results
  SpreadsheetApp.getActive().toast(
    `Verification complete: ${result.matchCount} matches, ${result.mismatchCount} mismatches, ${result.notFoundCount} not found.`,
    'Sheet Verification',
    5
  );
}
