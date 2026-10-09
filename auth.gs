function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Accounting') // Name of the menu
    .addItem('Copy Data to Sheets', 'copyDataToNewSheet') // Menu item and function name
    .addItem('Sort Data', 'sortSheetByNumericPart')
    .addItem('Sort All Data', 'sortAllSheetsByNumericPart')
    .addToUi()
  ui.createMenu('Utilities')
    .addItem('Check Job Numbers', 'compareMasterToLocationsWithDifferences')
    .addItem('Daily Job Tally', 'updateDailyJobCounts2025')
    .addToUi()
}

function sortAllSheetsByNumericPart() {
  try {
    // Get all sheets in the active spreadsheet
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    // Loop through each sheet and apply the sorting function
    sheets.forEach(sheet => {
      sortSheetByNumericPart(sheet);
    });

    Logger.log("Sorting applied to all sheets successfully.");
  } catch (error) {
    Logger.log(`Error sorting sheets: ${error.message}`);
  }
}
function sortSheetByNumericPart(sheet) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var data = range.getValues();

  // Separate headers (top 3 rows) from data
  var headers = data.slice(0, 1); // Top 3 rows
  var rowsToSort = data.slice(1); // All rows after the headers

  // Sort the rows based on the numeric part of the identifier in column A
  rowsToSort.sort(function(rowA, rowB) {
    var numA = parseInt(rowA[0].split('-')[1], 10); // Extract the numeric part of the identifier
    var numB = parseInt(rowB[0].split('-')[1], 10);
    Logger.log(numA +" " + numB);

    // Handle invalid or non-numeric data gracefully
    if (isNaN(numA) || isNaN(numB)) {
      Logger.log(`Invalid numeric values for sorting: "${rowA[0]}", "${rowB[0]}"`);
      return 0; // Keep original order if invalid
    }

    return numA - numB;
  });

  // Combine headers and sorted data
  var sortedData = headers.concat(rowsToSort);

  // Write the combined data back to the sheet
  range.setValues(sortedData);
  Logger.log(`Sheet "${sheet.getName()}" sorted successfully.`);
}


