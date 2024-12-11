function populateDropdowns() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = workbook.getSheetByName('Dashboard');
  const allSheets = workbook.getSheets();
  const dropdownCellRange = dashboard.getRange("B2:B4");
  
  const carVariants = [];
  
  // Collect unique car variant names from all sheets except 'Dashboard' and 'Notes'
  allSheets.forEach(sheet => {
    if (sheet.getName() !== 'Dashboard' && sheet.getName() !== 'Notes') {
      const lastColumn = sheet.getLastColumn();
      const row1 = sheet.getRange(1, 3, 1, lastColumn - 2).getValues()[0];
      carVariants.push(...row1);
    }
  });
  
  // Remove duplicates and sort variants
  const uniqueCarVariants = [...new Set(carVariants)].sort();
  
  // Add a placeholder option for the dropdown
  uniqueCarVariants.unshift("Select the car from the menu"); // Add placeholder as the first option
  
  // Populate dropdown menu
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueCarVariants, true)
    .setAllowInvalid(false)
    .build();
  dropdownCellRange.setDataValidation(rule);
  
  // Set initial placeholder text in the dropdown cells
  dropdownCellRange.setValue("Select the car from the menu");
}


function adjustSheetsVisibility() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = workbook.getSheetByName('Dashboard');
  const selectedCars = dashboard.getRange("B2:B4").getValues().flat().filter(car => car && car !== "Select the car from the menu");
  
  Logger.log("Selected Cars: " + selectedCars.join(", "));

  const allSheets = workbook.getSheets();
  
  allSheets.forEach(sheet => {
    if (sheet.getName() !== 'Dashboard' && sheet.getName() !== 'Notes') {
      const lastColumn = sheet.getLastColumn();
      const carVariants = sheet.getRange(1, 3, 1, lastColumn - 2).getValues()[0];
      const columnsToShow = [];

      Logger.log("Processing Sheet: " + sheet.getName());
      Logger.log("Car Variants in Sheet: " + carVariants.join(", "));

      // Create a map of the selected cars to their column indices
      selectedCars.forEach(selectedCar => {
        carVariants.forEach((variant, index) => {
          // Compare the variants exactly (with trim and case insensitivity)
          if (variant.trim().toLowerCase() === selectedCar.trim().toLowerCase()) {
            columnsToShow.push(index + 3); // Adjust for zero-based index, starting from column 3
          }
        });
      });

      // Sort columnsToShow to ensure the columns are displayed in the order of selection
      columnsToShow.sort((a, b) => a - b);  // Sort in ascending order of column index

      Logger.log("Columns to Show (Ordered): " + columnsToShow.join(", "));
      
      // Show and hide columns based on the selected cars' order
      for (let col = 3; col <= lastColumn; col++) {
        if (columnsToShow.includes(col)) {
          sheet.showColumns(col);
        } else {
          sheet.hideColumns(col);
        }
      }
      
      // Ensure first two columns are always visible
      sheet.showColumns(1, 2);
    }
  });
}



function resetSheetsVisibility() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = workbook.getSheetByName('Dashboard');
  const allSheets = workbook.getSheets();
  
  // Show all columns in all sheets except 'Dashboard' and 'Notes'
  allSheets.forEach(sheet => {
    if (sheet.getName() !== 'Dashboard' && sheet.getName() !== 'Notes') {
      const lastColumn = sheet.getLastColumn();
      sheet.showColumns(1, lastColumn);
    }
  });
  
  // Reset the dropdown cells to the placeholder text
  const dropdownCellRange = dashboard.getRange("B2:B4");
  dropdownCellRange.setValue("Select the car from the menu");
}

