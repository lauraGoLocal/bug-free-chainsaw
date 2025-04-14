# Guide to Modifying Google Ads Management Tool Using AI

*For non-technical team members who need to customize the Google Ads Management Tool*

## Introduction

This guide helps you customize your Google Ads Management Tool without needing programming experience. By using AI tools like Claude, ChatGPT, or Google Gemini, you can make changes to how budgets are calculated, update cell references, modify promotional text options, and work with the AdPulse Budget Verification feature.

## Table of Contents

1. [Before You Start](#before-you-start)
2. [Understanding Your Google Ads Management Tool](#understanding-your-google-ads-management-tool)
3. [Common Customization Scenarios](#common-customization-scenarios)
4. [AI Prompt Templates](#ai-prompt-templates)
   - [Modifying LastMonthBudget Logic](#modifying-lastmonthbudget-logic)
   - [Changing Cell References](#changing-cell-references)
   - [Customizing Promotional Text](#customizing-promotional-text)
   - [Modifying AdPulse Budget Verification](#modifying-adpulse-budget-verification)
   - [Modifying Spreadsheet Formulas](#modifying-spreadsheet-formulas)
   - [Troubleshooting Issues](#troubleshooting-issues)
5. [Documenting Your Changes](#documenting-your-changes)
6. [Testing Your Modifications](#testing-your-modifications)

## Before You Start

Always follow these steps before making any changes:

1. **Make a backup**: Create a copy of your Google Sheet with all its scripts
2. **Document the current state**: Note what's working and what needs changing
3. **Start small**: Make one change at a time and test before moving on

## Understanding Your Google Ads Management Tool

Your Google Ads Management Tool consists of several parts:

1. **Budget Timestamp System**: Automatically records when changes are made to budgets
2. **LastMonthBudget Calculator**: Figures out what budget should be used based on previous month's data
3. **Ad Customizer Tool**: Creates promotional text following specific rules
4. **AdPulse Check**: Monitors changes to approved budgets
5. **AdPulse Budget Verification**: Automatically verifies that budget amounts in your spreadsheet match the corresponding amounts in AdPulse

### Spreadsheet Structure

Based on the screenshots, your spreadsheet includes:

- **Campaign information** in columns A-C (Engine, Campaign Type, Campaign Name)
- **Budget calculations** in columns E-G (Budget, Management Fee, Budget Total)
- **Approval columns** in columns H-K (Recommended Budget Approved, No Ads This Month, Pre-Approved Client Budgets, Last Month Approved Budget)
- **Budget Export values** in column V that determine what gets sent to AdPulse
- **AdPulse Match** column W that shows verification status when using the AdPulse Budget Verification feature

### Key Formulas

Two important formulas drive the budget calculations:

1. **Last Month Budget Formula** (Column K):
   ```
   =LASTMONTHBUDGET("", C6)
   ```
   This formula looks up the previous month's approved budget for the campaign in cell C6.

2. **AdPulse Budget Export Formula** (Column V):
   ```
   =IFERROR(
       IFS(
           I6=TRUE, 0,
           H6=TRUE, E6,
           AND(ISNUMBER(J6), J6>0), ROUND(J6 / 1.2, 0),
           ISNUMBER(LASTMONTHBUDGET(CURRENT_SHEET_NAME(), C6)), 
               ROUND(LASTMONTHBUDGET(CURRENT_SHEET_NAME(), C6) / 1.2, 0),
           TRUE, "No valid budget found"
       ),
       "Error in calculation"
   )
   ```
   
   This formula determines what budget value should be exported to AdPulse based on the following conditions:
   - If "No Ads This Month" is checked (I6=TRUE) → Returns 0
   - If "Recommended Budget" is approved (H6=TRUE) → Returns the budget amount (E6)
   - If there's a pre-approved budget (J6) → Returns that amount divided by 1.2 (removing management fee)
   - If there's a value from LastMonthBudget → Returns that value divided by 1.2
   - Otherwise → Returns "No valid budget found"

## Common Customization Scenarios

You might need to customize:

1. **Business Logic**: Change how LastMonthBudget calculates values based on your specific needs
2. **Cell References**: Update which columns the system looks at
3. **Promotional Text**: Add or modify the text options available in the Ad Customizer
4. **Budget Verification**: Customize how the AdPulse Budget Verification works
5. **Spreadsheet Formulas**: Update formulas that determine budget calculations
6. **Troubleshoot Issues**: Fix problems that come up after changes are made

## AI Prompt Templates

### Modifying LastMonthBudget Logic

The LastMonthBudget function determines what budget should be used based on previous month's data. You might need to change its logic to match your business rules.

#### How it currently works

The function checks these conditions in order:
1. If "No Ads This Month" was checked → Returns 0
2. If "Recommended Budget" was approved → Returns that budget amount
3. If there was a pre-approved budget → Returns that amount
4. If there was a previously approved budget → Returns that amount
5. If none of the above → Returns "#N/A"

#### Prompt template for modifying LastMonthBudget logic

```
I need help modifying the business logic in a Google Apps Script function called LASTMONTHBUDGET. This function determines what budget should be used for a Google Ads campaign based on previous month's data.

Here's the current function:

[COPY AND PASTE THE LASTMONTHBUDGET FUNCTION HERE]

I need to change the logic to: [EXPLAIN YOUR NEW BUSINESS RULE]

For example, I want to: 
- Change the order of conditions
- Add a new condition
- Modify how certain values are calculated
- Change what value is returned in certain cases

Please modify the function to follow these new rules while keeping the rest of the code structure intact. Explain the changes you made in simple terms.
```

#### Example customization request

```
I need help modifying the business logic in a Google Apps Script function called LASTMONTHBUDGET. This function determines what budget should be used for a Google Ads campaign based on previous month's data.

Here's the current function:

function LASTMONTHBUDGET(sheetName, matchValue) {
  if (!sheetName) {
    sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    console.log("No sheet name provided. Using active sheet: " + sheetName);
  }

  console.log("Looking for campaign: '" + matchValue + "'");

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
  console.log("Found row index: " + prevRow);

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

I need to change the logic to prioritize Pre-approved budgets over Recommended budgets. Also, if there's no pre-approved budget but the recommended budget is over $1000, I want to cap it at $1000 for safety. Please modify the function to follow these new rules.
```

### Modifying Spreadsheet Formulas

Your spreadsheet uses formulas that rely on the same business logic as the custom functions. When you modify the functions, you may also need to update these formulas.

#### Prompt template for modifying spreadsheet formulas

```
I need help updating a Google Sheets formula that uses our custom LASTMONTHBUDGET function. 

Here's the current formula:

[COPY AND PASTE THE FORMULA HERE]

I need to change this formula because:
[EXPLAIN WHAT YOU WANT TO CHANGE]

The LASTMONTHBUDGET function has been modified to:
[EXPLAIN HOW THE FUNCTION NOW WORKS]

Please update this formula to match our new business rules while keeping all the error handling and structure intact.
```

#### Example formula update request

```
I need help updating a Google Sheets formula that uses our custom LASTMONTHBUDGET function.

Here's the current formula in column V that determines our AdPulse export value:

=IFERROR(
    IFS(
        I6=TRUE, 0,
        H6=TRUE, E6,
        AND(ISNUMBER(J6), J6>0), ROUND(J6 / 1.2, 0),
        ISNUMBER(LASTMONTHBUDGET(CURRENT_SHEET_NAME(), C6)), 
            ROUND(LASTMONTHBUDGET(CURRENT_SHEET_NAME(), C6) / 1.2, 0),
        TRUE, "No valid budget found"
    ),
    "Error in calculation"
)

I need to change this formula because we've updated our business rules to prioritize Pre-approved budgets over Recommended budgets, and to cap any recommended budget at $1000 for safety.

The LASTMONTHBUDGET function has been modified to prioritize Pre-approved budgets first, before checking if the Recommended Budget was approved.

Please update this formula to match our new business rules while keeping all the error handling and structure intact.
```

### Changing Cell References

If your spreadsheet structure changes, you might need to update which columns the system uses.

#### Prompt template for updating cell references

```
I need help updating column references in a Google Apps Script for our Google Ads Management Tool. When our spreadsheet structure changes, we need to update which columns the system looks at.

Here's the current code:

[COPY AND PASTE THE RELEVANT CODE SECTION]

I need to change these column references because:
[EXPLAIN YOUR CHANGES TO THE SPREADSHEET]

Specifically, I need to update:
- Column [OLD COLUMN] to column [NEW COLUMN]
- Column [OLD COLUMN] to column [NEW COLUMN]

Please modify the code to use the new column references while keeping everything else working the same.
```

#### Example cell reference update request

```
I need help updating column references in our Google Ads Management Tool. When our spreadsheet structure changes, we need to update which columns the system looks at.

Here's the current code from the onEdit.gs file:

function onEdit(event) {
  if (!event) {
    throw new Error(
      "Please do not run the script in the script editor window. It runs automatically when you edit the spreadsheet."
    );
  }

  processSheets(event, [
    {
      description: 'Sheets that end with "24 Budgets" or "25 Budgets"',
      sheetsToWatch: /^.*2[45] Budgets$/i,
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

  updateAdPulseCheck(event);
}

We've renamed some column headers in our spreadsheet:
- "Recommended Budget Approved" is now "Approved Budget"
- "Approval Timestamp" is now "Approval Date"
- "Pre-Approved Client Budgets" is now "Client Pre-Approved Amount"
- "Pre-Approval Timestamp" is now "Pre-Approval Date"

Please update the code to match our new column names while keeping all the functionality the same.
```

### Customizing Promotional Text

You might want to add new promotional options or change how the Ad Customizer works.

#### Prompt template for customizing promotional text

```
I need help modifying the Ad Customizer tool in our Google Ads Management System. This tool creates promotional text for our Google Ads.

Here's the current code section that defines the promotional options:

[COPY AND PASTE THE RELEVANT SECTION FROM adCustomizer.gs OR adCustomizerPage.html]

I want to make these changes:
[DESCRIBE WHAT NEW PROMOTIONS OR CHANGES YOU WANT]

For example:
- Add new standard promotion types
- Change the character limit
- Modify how discounts are formatted
- Change how multiple promotions are combined

Please modify the code to include these changes and explain what you updated.
```

#### Example promotional text customization

```
I need help modifying the Ad Customizer tool in our Google Ads Management System. This tool creates promotional text for our Google Ads.

Here's the current code section from adCustomizerPage.html that defines the standard promotion options:

<select id="standardPromo" class="form-control" onchange="updateCharCount()">
  <option value="first_month_dollar">$1 First Month's Rent</option>
  <option value="move_in">Move-In Specials</option>
  <option value="online_rates">See Our Online Rates Today</option>
  <option value="online_specials">View Online Specials</option>
  <option value="web_exclusive">Web-Exclusive Discounts</option>
  <option value="first_month_free">First Month Free</option>
  <option value="military">Military Special</option>
  <option value="first_responder">First Responder Discounts</option>
</select>

I want to add these new promotion options:
- "Student Discount" that shows as "Student Discount Available"
- "Senior Discount" that shows as "Senior Citizen Discount"
- "Holiday Special" that shows as "Limited Time Holiday Special"

I also want to change "Military Special" to say "Military Discount Available" instead.

Please show me the updated code with these changes.
```

### Modifying AdPulse Budget Verification

You might need to customize how the AdPulse Budget Verification feature works to match your needs.

#### Prompt template for AdPulse Budget Verification

```
I need help modifying the AdPulse Budget Verification feature in our Google Ads Management Tool. This feature verifies that budget amounts in our spreadsheet match the corresponding amounts in AdPulse.

Here's the current code:

[COPY AND PASTE THE RELEVANT SECTION FROM adPulseBudgetVerification.gs]

I want to make these changes:
[DESCRIBE WHAT CHANGES YOU WANT]

For example:
- Update the column references
- Change how matches/mismatches are displayed
- Modify the verification logic
- Update the sheet matching pattern

Please modify the code to include these changes and explain what you updated.
```

#### Example AdPulse Budget Verification customization

```
I need help modifying the AdPulse Budget Verification feature in our Google Ads Management Tool. This feature verifies that budget amounts in our spreadsheet match the corresponding amounts in AdPulse.

Here's the current function for verification from adPulseBudgetVerification.gs:

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

We've made some changes to our spreadsheet:
1. We've moved the Campaign Name column to column T (index 19) instead of U (index 20)
2. We want to change the mismatch symbol from '❌' to '🚨'
3. We want to add a tolerance of $1 for budget mismatches (so budgets that differ by $1 or less should be considered matches)

Please update the code to accommodate these changes while keeping the rest of the functionality intact.
```

### Troubleshooting Issues

When something goes wrong, you can use AI to help troubleshoot issues.

#### Prompt template for troubleshooting

```
I'm having a problem with our Google Ads Management Tool in Google Sheets after making some changes. Here's what's happening:

[DESCRIBE THE ERROR OR ISSUE IN DETAIL]

The error occurs when:
[EXPLAIN EXACTLY WHEN THE ERROR HAPPENS]

Here's the relevant code that might be causing the issue:

[COPY AND PASTE THE CODE SECTION YOU RECENTLY MODIFIED]

Before my changes, it was working fine. I changed:
[DESCRIBE WHAT YOU CHANGED]

Can you help identify what might be causing this problem and suggest how to fix it?
```

#### Example troubleshooting request

```
I'm having a problem with our Google Ads Management Tool in Google Sheets after making some changes. Here's what's happening:

When I try to approve a budget by checking the "Approved Budget" box, I get an error message saying "Cannot find column: Approval Date" and the timestamp isn't being added.

The error occurs when:
I check the box in the "Approved Budget" column for any campaign row.

Here's the relevant code that I recently modified:

function onEdit(event) {
  if (!event) {
    throw new Error(
      "Please do not run the script in the script editor window. It runs automatically when you edit the spreadsheet."
    );
  }

  processSheets(event, [
    {
      description: 'Sheets that end with "24 Budgets" or "25 Budgets"',
      sheetsToWatch: /^.*2[45] Budgets$/i,
      columnConfigs: [
        {
          description: "Update 'Approval Date' when 'Approved Budget' is edited.",
          labelToWatch: "Approved Budget",
          labelToStamp: "Approval Date",
          valueToWatch: true,
        },
        {
          description: "Update 'Disapproval Timestamp' when 'No Ads This Month' is edited.",
          labelToWatch: "No Ads This Month",
          labelToStamp: "Disapproval Timestamp",
          valueToWatch: true,
        },
        {
          description: "Update 'Pre-Approval Date' when 'Client Pre-Approved Amount' is edited.",
          labelToWatch: "Client Pre-Approved Amount",
          labelToStamp: "Pre-Approval Date",
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

  updateAdPulseCheck(event);
}

Before my changes, it was working fine with the old column names "Recommended Budget Approved" and "Approval Timestamp". I changed these to "Approved Budget" and "Approval Date" in both the code and my spreadsheet.

Can you help identify what might be causing this problem and suggest how to fix it?
```

## Documenting Your Changes

Keeping track of changes is crucial. Create a simple change log in your spreadsheet with:

| Date | What Changed | Why It Changed | Modified By |
|------|--------------|---------------|-------------|
| 4/14/2025 | Updated column references for budget approval | Column names changed in spreadsheet | Jane Smith |
| 4/14/2025 | Modified LastMonthBudget to prioritize pre-approved budgets | New business requirement | Jane Smith |

## Testing Your Modifications

After any change, test that the system still works correctly:

1. **For LastMonthBudget changes**: 
   - Check the results for multiple campaigns with different scenarios
   - Compare with manual calculations to verify correctness
   - Verify that column K shows the correct Last Month Approved Budget values
   - Check column V to ensure AdPulse export values reflect your new business rules

2. **For cell reference changes**:
   - Trigger the specific action (like checking a box)
   - Verify the timestamp appears in the right column
   - Check that any formulas using these columns still work correctly

3. **For formula changes**:
   - Check multiple rows with different scenarios (No Ads, Approved Budget, Pre-Approved)
   - Verify column V shows correct values based on your business rules
   - Check for any #ERROR values or unexpected results

4. **For promotion text changes**:
   - Open the Ad Customizer
   - Create promotions using your new options
   - Check that text displays correctly and character limits work

5. **For AdPulse Budget Verification**:
   - Paste data from AdPulse into the AdPulse sheet
   - Check that verification runs automatically
   - Verify the match/mismatch indicators appear correctly in column W
   - Test with known matches and mismatches to ensure the logic works as expected

### Testing Checklist

Use this simple checklist to verify changes:

| Test | Expected Result | Actual Result | Pass/Fail |
|------|-----------------|--------------|-----------|
| Check "No Ads This Month" | AdPulse value (col V) = 0 | | |
| Check "Recommended Budget Approved" | AdPulse value = Budget amount | | |
| Enter Pre-Approved Client Budget | AdPulse value = Pre-Approved ÷ 1.2 | | |
| Test LastMonthBudget function | Shows correct previous month's budget | | |
| Test timestamp for approval | Timestamp appears in correct column | | |
| Test Ad Customizer | Promotion displays correctly | | |
| Paste AdPulse data | Verification runs with correct matches/mismatches | | |

If your test fails, use the [Troubleshooting Prompt Template](#troubleshooting-issues) to get help from AI.

---

Remember, when working with AI:
- Be specific about what you want to change
- Provide the exact current code section
- Explain your business requirements clearly
- Test each change thoroughly before moving to the next

By following this guide, you'll be able to make necessary modifications to your Google Ads Management Tool without needing programming expertise.
