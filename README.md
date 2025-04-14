# Google Ads Management Tool

A Google Apps Script project that helps storage facility managers and specialists create, manage, and track advertising budgets and promotional text for Google Ads campaigns.

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

## File Structure

- **adCustomizer.gs** - Backend logic for the promotion editor dialog
- **adCustomizerPage.html** - Frontend UI for the promotion editor
- **onEdit.gs** - Event handler for sheet edits to add timestamps
- **AdPulseCheck.gs** - Monitors changes to approved budgets
- **processSheets.gs** - Core engine for timestamp functionality
- **sheetHelpers.gs** - Utility functions including LASTMONTHBUDGET

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

## Notes
- This script only works on sheets that end with "24 Budgets" or "25 Budgets"
- Maximum promotion length is 30 characters (including the * symbol)
- When combining promotions (Move-In + discount), the text automatically shortens to "Move-In+" to save characters
