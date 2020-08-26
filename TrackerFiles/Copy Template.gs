function copyMasterTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheetNames = getSheetNames();
  var masterSheet = ss.getSheetByName("MASTER");
  var response = ui.prompt("New Sheet Name Required", "What would you like to name this sheet?", ui.ButtonSet.OK_CANCEL);  
  
  //Confirm user input
  if (response.getSelectedButton() == ui.Button.CANCEL) {
    ui.alert("No sheet was created");
  }
  else {
    var newSheetName = response.getResponseText();
    // Check if the sheet exists, else copy the master with the given name
    if (sheetNames.includes(newSheetName)) {
      ui.alert("The sheet with " + newSheetName + " already exists. Please choose a unique name.", ui.ButtonSet.OK)
    }
    else {
      masterSheet.copyTo(ss).setName(newSheetName).showSheet();
      masterSheet.hideSheet();
    };
  };
  
}
