# Google Ads Management Tool

A Google Apps Script project that helps storage facility managers create, manage, and track advertising budgets and promotional text for Google Ads campaigns.

## Features

### Ad Customizer Tool
Create promotional text for Google Ads with a user-friendly interface:
- Create standardized promotion texts for Google Ads
- Set custom discount amounts (dollar or percentage)
- Configure time periods (monthly or lifetime)
- Add Move-In Special promotions
- Include optional Free Lock promotion
- Combine multiple promotion types
- Character counter with 30-character limit validation
- Search and select facilities from a dynamic list
- Automatic caching for performance

### Budget Management
- Automatic timestamp tracking when budget items are modified
- Track approvals and changes to budgets
- Pre-approval handling
- Audit trail for budget changes
- Intelligent budget determination with LASTMONTHBUDGET function

### AdPulse Budget Verification
- Automatically verifies that budget amounts match the corresponding amounts in AdPulse
- Highlights matches, mismatches, and missing items with color-coding
- Provides clear status indicators in the AdPulse Match column
- Shows the exact budget value from AdPulse when there's a mismatch
- Easy setup - just paste data from AdPulse into the AdPulse sheet
- Case-insensitive matching for campaign names
- Detailed verification summary with counts of matches, mismatches, and not found items

## File Structure

- **adCustomizer.gs** - Backend logic for the promotion editor dialog
- **adCustomizerPage.html** - Frontend UI for the promotion editor
- **onEdit.gs** - Event handler for sheet edits to add timestamps
- **AdPulseCheck.gs** - Monitors changes to approved budgets
- **processSheets.gs** - Core engine for timestamp functionality
- **sheetHelpers.gs** - Utility functions including LASTMONTHBUDGET
- **adPulseBudgetVerification.gs** - Handles verification between spreadsheet budgets and AdPulse data

## Setup Instructions

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Copy each file from this repository into your Apps Script project
4. Save all files and close the editor
5. Refresh your sheet
6. A new menu item "Ad Customizer" will appear in your spreadsheet

## Usage

### Ad Customizer
1. Click the "Ad Customizer" menu item
2. Select "Open Tool"
3. Search and select a facility
4. Configure your promotion:
   - Add a Move-In Special with a custom dollar amount
   - Select a standard promotion or create a custom one
   - For custom promotions, choose dollar or percentage discounts
   - Set time period (1-6 months or lifetime)
   - Add Free Lock promotion if desired
5. Review the preview and character count
6. Click "Save Promotion"

### Budget Timestamps
Timestamps are added automatically when:
1. "Recommended Budget Approved" is checked → Updates "Approval Timestamp"
2. "No Ads This Month" is checked → Updates "Disapproval Timestamp"
3. "Pre-Approved Client Budgets" is edited → Updates "Pre-approval Timestamp"
4. "Budget Total" is edited → Updates "Budget Total Timestamp"

### AdPulse Budget Verification
1. Go to the "AdPulse" sheet (it will be created automatically if it doesn't exist)
2. Paste data from AdPulse including the "Budget Name" and "Budget Target" columns
3. The verification will run automatically
4. Review the results in column W "AdPulse Match" of your budget sheets
   - "✓ Match" - Budget values match
   - "❌ Mismatch: $xxx" - Values don't match, shows the AdPulse value
   - "⚠️ Not Found" - Campaign not found in AdPulse data

## Notes
- This script only works on sheets that end with "25 Budgets" or "26 Budgets"
- Maximum promotion length is 30 characters (including the * symbol)
- When combining promotions (Move-In + discount), the text automatically shortens to "Move-In+" to save characters
- For AdPulse verification, make sure your pasted data includes column headers "Budget Name" and "Budget Target"

## Customization Guide

For non-technical team members who need to customize this tool, please refer to the "Guide to Modifying Google Ads Management Tool Using AI" document. This guide provides step-by-step instructions for using AI tools like Claude, ChatGPT, or Google Gemini to make common modifications.
