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
      const row1 = sheet.getRange(1, 3, 1, lastColumn - 2).getValues()[0]; // Get car variants from first row
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

  // Perform status check after populating the dropdown
  performStatusCheck();
}



function performStatusCheck() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = workbook.getSheetByName('Dashboard');
  const selectedCars = dashboard.getRange("B2:B4").getValues().flat().filter(car => car && car !== "Select the car from the menu");
  
  // Get the car variants for each sheet
  const allSheets = workbook.getSheets();
  const statusCheckSheet = dashboard.getRange("J2:L2");  // Assume this is where status check will be placed
  
  let status = "OK";
  let differences = [];
  
  // Collect car variants from the first row (C onwards) of each sheet
  const variantsList = [];
  allSheets.forEach(sheet => {
    if (sheet.getName() !== 'Dashboard' && sheet.getName() !== 'Notes') {
      const lastColumn = sheet.getLastColumn();
      const carVariants = sheet.getRange(1, 3, 1, lastColumn - 2).getValues()[0]; // First row car variants
      variantsList.push({ sheetName: sheet.getName(), carVariants });
    }
  });
  
  // Compare the car variants across all sheets
  const firstSheetVariants = variantsList[0].carVariants;
  
  variantsList.forEach(item => {
    const sheetName = item.sheetName;
    const sheetVariants = item.carVariants;
    
    if (sheetVariants.length !== firstSheetVariants.length || 
        !sheetVariants.every((variant, index) => variant.trim().toLowerCase() === firstSheetVariants[index].trim().toLowerCase())) {
      status = "NOT OK";
      differences.push(`Mismatch in sheet: ${sheetName}. Variants: ${sheetVariants.join(", ")}`);
    }
  });

  // Set status message in the Dashboard (J2, K2, L2)
  dashboard.getRange("J2").setValue("Check-1");
  dashboard.getRange("K2").setValue("Models in all sheets are same");
  dashboard.getRange("L2").setValue(status);
  
  if (status === "NOT OK") {
    dashboard.getRange("L3").setValue(differences.join(", "));
  } else {
    dashboard.getRange("L3").setValue("");
  }
}


function adjustSheetsVisibility() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = workbook.getSheetByName('Dashboard');
  const selectedCars = dashboard.getRange("B2:B4").getValues().flat().filter(car => car && car !== "Select the car from the menu");

  const allSheets = workbook.getSheets();

  allSheets.forEach(sheet => {
    if (sheet.getName() !== 'Dashboard' && sheet.getName() !== 'Notes') {
      const lastColumn = sheet.getLastColumn();
      const carVariants = sheet.getRange(1, 3, 1, lastColumn - 2).getValues()[0];
      const columnsToShow = [];

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
