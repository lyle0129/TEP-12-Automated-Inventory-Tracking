function appendEWLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Collect the data
  const sourceRange = ss.getRangeByName("InputData");
  const sourceVals = sourceRange.getValues().flat();

  // Validate all Cells are filled
  const anyEmptyCell = sourceVals.findIndex(cell => cell == "");

  if (anyEmptyCell !== -1){
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Input Incomplete !!",
      "Please enter a value in ALL input cells before submitting",
      ui.ButtonSet.OK
    );
    return;
  };

  // Gather current data
  const data = [...sourceVals];

  // Append the data to data logger
  const destinationSheet = ss.getSheetByName("DataLogger");
  
  // Get the last row with data
  const lastRow = destinationSheet.getLastRow();
  
  // Check if last row is above 10,000
  let targetRow = lastRow + 1;
  if (targetRow >= 10000) {
    // Find the first empty row below 10,000
    targetRow = findFirstEmptyRowBelow(destinationSheet, 10000);
  }

  // Append data to the target row
  destinationSheet.getRange(targetRow, 1, 1, data.length).setValues([data]);

  // Clear the source sheet rows
  sourceRange.clearContent();

  // Set a specific cell's value after clearing the content
  const inputSheet = ss;
  inputSheet.getRange("C11").setValue("None");
  // inputSheet.getRange("C4").setValue("Drop Off");
  // inputSheet.getRange("C5").setValue("Not Applicable");

  ss.toast("Success. Para sa Bayan, Para sa Mundo !!");
}

// Function to find the first empty row below a certain row number
function findFirstEmptyRowBelow(sheet, belowRow) {
  const range = sheet.getRange(1, 1, belowRow - 1, 1);
  const values = range.getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === "") {
      return i + 1; // Rows are 1-indexed
    }
  }
  return belowRow; // If no empty row found, return belowRow
}



