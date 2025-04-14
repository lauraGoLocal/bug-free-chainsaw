// This is the engine behind our automatic timestamp system. It handles all the detailed work
// of finding the right columns and adding timestamps when changes are made.
//
// What does this script do?
// When someone edits a cell that we're monitoring (like marking a budget as approved),
// this script:
// 1. Checks if we should be watching this particular sheet
// 2. Finds the correct column headers (like "Approval Timestamp")
// 3. Adds the current date and time in the right place
// 4. Shows a small notification message to confirm the timestamp was added
//
// If anything goes wrong (like if column names have been changed), it will show
// an error message explaining what happened instead of silently failing.
//
// The timestamp format looks like this: "Tuesday 4/5/16 at 4:08 P.M."
//
// Important: This script works behind the scenes and is called automatically
// by our other scripts. You don't need to run it directly.
//
// Examples of what it does:
// - When you mark "Recommended Budget Approved" as checked, it finds the
//   "Approval Timestamp" column and adds the current time
// - If you clear a Budget Total, it will also clear the corresponding
//   timestamp (but only for columns where we've specifically set up this behavior)
// - It can handle multiple rows being changed at once (like if you paste
//   values into several rows)

/**
 * Shows a message in a pop-up and throws an error with the message.
 *
 * @param {String} message The message to show and throw.
 */
function showAndThrowError(message) {
  showMessage(message, 30);
  throw new Error(message);
}

/**
 * Shows a message in a pop-up.
 *
 * @param {String} message The message to show.
 */
function showMessage(message, timeoutSeconds) {
  SpreadsheetApp.getActive().toast(message, "Timestamp multiple columns", timeoutSeconds || 5);
}

function findSheetSettings(event, sheetSettings) {
  if (!event.source.getActiveRange() || JSON.stringify(event.range) === "{}") {
    showAndThrowError("Invalid sheet");
  }

  const activeSheet = event.source.getActiveSheet();
  const activeSheetName = activeSheet.getName();
  const settings = sheetSettings.find(({ sheetsToWatch }) => sheetsToWatch.test(activeSheetName));

  //if (!settings) {
//    showAndThrowError(`No settings found for ${activeSheetName}`);
 // }

  return { activeSheet, activeSheetName, columnConfigs: settings.columnConfigs };
}

function getColumnLabels(event) {
  const firstRow = event.range.getRow();
  const sheet = event.source.getActiveSheet();
  const labelRow = sheet.getFrozenRows() || 1;
  const numRows = 1;
  const firstColumn = 1;
  const lastColumn = sheet.getLastColumn();

  if (firstRow <= labelRow) {
    showAndThrowError("Could not find label row");
  }

  return sheet.getRange(labelRow, firstColumn, numRows, lastColumn).getValues()[0];
}

function getColumnIndex(sheetName, columnLabels, targetColumnLabel) {
  const watchColumnIndex = columnLabels.indexOf(targetColumnLabel) + 1 || null;
  if (!watchColumnIndex) {
    showAndThrowError(`Could not find ${targetColumnLabel} in ${sheetName}.`);
  }

  return watchColumnIndex;
}

function processSheets(event, sheetSettings) {
  try {
    const { activeSheet, activeSheetName, columnConfigs } = findSheetSettings(event, sheetSettings);
    const columnLabels = getColumnLabels(event);
    const editedColumn = event.range.getColumn();
    const columnConfig = columnConfigs.find(({ labelToWatch }) => {
      const watchColumnIndex = getColumnIndex(activeSheetName, columnLabels, labelToWatch);
      return watchColumnIndex === editedColumn;
    });

    if (!columnConfig) {
      // No configuration was found for the modified column
      return;
    }

    const editedRow = event.range.getRow();
    const rangeValues = event.range.getValues();
    const stampColumnIndex = getColumnIndex(activeSheetName, columnLabels, columnConfig.labelToStamp);

    for (let row = 0; row < event.range.getNumRows(); row++) {
      // Iterate through all rows in the modified range
      const cellValue = rangeValues[row][0];

      if (columnConfig.valueToWatch === undefined || columnConfig.valueToWatch === cellValue) {
        // If watching for any change OR if the cell value matches the value we're watching for...
        const timestampCell = activeSheet.getRange(row + editedRow, stampColumnIndex);

        if (String(cellValue).length === 0 && columnConfig.eraseTimestamp) {
          // If the cell is cleared and we can erase the timestamp, then erase the timestamp
          showMessage(`Clearing ${columnConfig.labelToStamp} on row ${row + editedRow} in ${activeSheetName}`);
          timestampCell.setValue(null);
        } else {
          // If the cell has a value or we cannot erase the timestamp, then update the timestamp
          showMessage(`Updating ${columnConfig.labelToStamp} on row ${row + editedRow} in ${activeSheetName}`);

          /**
           * gives a timestamp format like Tuesday, 4/5/16 at 4:08 PM
           * @see https://developers.google.com/sheets/api/guides/formats
           */
          timestampCell.setValue(new Date()).setNumberFormat('dddd m/d/yy at hh:mm A/P".M."');
        }
      }
    }
  } catch (error) {
    showAndThrowError_(error.message + ", stack: " + error.stack);
  }
}
