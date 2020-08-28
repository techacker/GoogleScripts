// ********* Start: Main Function to add new row in Summary Sheet that calls 'Copy Master Template function'

function addNewRowInSummarySheet() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("SUMMARY");
  var newSheetName = copyMasterTemplate();
  var range = summarySheet.getDataRange();
  var lc = range.getLastColumn();
  var firstSummaryRow = summarySheet.getRange(3, 1, 1, lc);
  var formulas = summarySheet.getRange(3, 2, 1, lc).getFormulas();
  
  /*
  if (newSheetName.length !== 0) {
    // Shift one row down
    firstSummaryRow.insertCells(SpreadsheetApp.Dimension.ROWS);
    
    // Set first cell value to be the sheet name
    summarySheet.getRange(3, 1).setValue(newSheetName);
    
    // Set formulas for other cells in this new row.
    for (var i=0; i<formulas[0].length-1; i++) {
      var splits = formulas[0][i].split('!');  // Since formula is referenced with '!'
      summarySheet.getRange(3, i+2).setValue(`=${newSheetName}!${splits[1]}`);
    };
  };
  */
}

// ********* End: Main Function to add new row in Summary Sheet that calls 'Copy Master Template function'



// ********* Start: Copy Master Template

function copyMasterTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheetNames = getSheetNames();
  var masterSheet = ss.getSheetByName("MASTER");
  var response = ui.prompt("Info Needed", "What would you like to name this sheet?", ui.ButtonSet.OK_CANCEL);  
  
  //Confirm user input
  if (response.getSelectedButton() == ui.Button.CANCEL) {
    var newSheetName = "";
    ui.alert("No sheet was created");
  }
  else {
    var newSheetName = response.getResponseText();
    // Check if the sheet exists, else copy the master with the given name
    if (sheetNames.includes(newSheetName)) {
      ui.alert("The sheet with " + newSheetName + " name already exists. Please choose a unique name.", ui.ButtonSet.OK)
    }
    // Check if the given name has multiple words
    else if (newSheetName.split(" ").length>1) {
      newSheetName = newSheetName.split(" ").join("");
      if (sheetNames.includes(newSheetName)){
        ui.alert("The sheet with " + newSheetName + " name already exists. Please choose a unique name.", ui.ButtonSet.OK)
        var newSheetName = "";
      }
      else {
        masterSheet.copyTo(ss).setName(newSheetName).showSheet();
        masterSheet.hideSheet();
      }
    }
    else {
      masterSheet.copyTo(ss).setName(newSheetName).showSheet();
      masterSheet.hideSheet();
    }
  };
  
  return newSheetName;
  
}

// ********* End: Copy Master Template
