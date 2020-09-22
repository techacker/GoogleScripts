// ********* Start: Main Function to add new row in Summary Sheet that calls 'Copy Master Template function'

function addNewRowInSummarySheet() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("SUMMARY");
  var lc = summarySheet.getLastColumn();
  var lr = summarySheet.getLastRow();
  var newSheetName = copyMasterTemplate();
  var range = summarySheet.getRange(1, 1, lr, lc).getDisplayValues();
  var TabRowNum = 0;
  
  
  // Get "Tab" Row number
  for (var i=0; i<range.length; i++) {
    if (range[i][0].toUpperCase() === "TAB") {
      TabRowNum = i+1;
    }
  }
  
  var MasterRow = summarySheet.getRange(TabRowNum+1, 1, 1, lc);
  var formulas = summarySheet.getRange(TabRowNum+1, 2, 1, lc).getFormulas();

  // Logger.log(formulas);
  // [[=MASTER!H1, =MASTER!H2, =MASTER!P1, 
  // =IFERROR(J9/D9,0), =IFERROR(K9/D9,0), =IFERROR(L9/D9,0), =IFERROR(M9/D9,0), 
  // =MASTER!AD1, =MASTER!Q1, =MASTER!R1, =MASTER!S1, =MASTER!U1, ]]
  
  
  if (newSheetName.length !== 0) {
    // Shift one row down
    MasterRow.insertCells(SpreadsheetApp.Dimension.ROWS);
    var Tab = summarySheet.getRange(TabRowNum+1, 1).setValue(newSheetName);
    var PPPMEngr = summarySheet.getRange(TabRowNum+1, 2).setValue(`=${newSheetName}!H1`);
    var MRD = summarySheet.getRange(TabRowNum+1, 3).setValue(`=${newSheetName}!H2`);
    var NoOfParts = summarySheet.getRange(TabRowNum+1, 4).setValue(`=${newSheetName}!P1`);
    var Cost = summarySheet.getRange(TabRowNum+1, 9).setValue(`=${newSheetName}!AD1`);
    var REQSubmitted = summarySheet.getRange(TabRowNum+1, 10).setValue(`=${newSheetName}!Q1`);
    var POIssued = summarySheet.getRange(TabRowNum+1, 11).setValue(`=${newSheetName}!R1`);
    var PartsReceived = summarySheet.getRange(TabRowNum+1, 12).setValue(`=${newSheetName}!S1`);
    var RFQPending = summarySheet.getRange(TabRowNum+1, 13).setValue(`=${newSheetName}!U1`);
    var PercREQ = summarySheet.getRange(TabRowNum+1, 5).setValue(`=IFERROR(J9/D9,0)`);
    var PercPO = summarySheet.getRange(TabRowNum+1, 6).setValue(`=IFERROR(K9/D9,0)`);
    var PercRecd = summarySheet.getRange(TabRowNum+1, 7).setValue(`=IFERROR(L9/D9,0)`);
    var PercRFQPending = summarySheet.getRange(TabRowNum+1, 8).setValue(`=IFERROR(M9/D9,0)`);
    
  };
  
  // Hide the Master Row to avoid confusion.
  for (var i=0; i<range.length; i++) {
    if (range[i][0].toUpperCase() === "MASTER") {
      var MasterRowNum = i+2;
      summarySheet.hideRows(MasterRowNum);
      ss.getSheetByName("MASTER").hideSheet();
    }
  }
  
  
}

// ********* End: Main Function to add new row in Summary Sheet that calls 'Copy Master Template function'



// ********* Start: Copy Master Template

function copyMasterTemplate() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheetNames = getSheetNames();
  var response = ui.prompt("Info Needed", "What would you like to name this sheet?", ui.ButtonSet.OK_CANCEL); 
  
  // Master Template File: 1MA-DXm35QNXrbAVrxtdwTfNVdTk3p8n5Brh8Ya1oCIk/edit#gid=647179545
  const MasterTempFile = SpreadsheetApp.openById("1MA-DXm35QNXrbAVrxtdwTfNVdTk3p8n5Brh8Ya1oCIk");
  
  //Confirm user input
  if (response.getSelectedButton() == ui.Button.CANCEL || response.getSelectedButton() == ui.Button.CLOSE) {
    var newSheetName = "";
    ui.alert("No sheet was created");
  }
  else {
    var newSheetName = response.getResponseText();
    // Check if the sheet exists, else copy the master with the given name
    if (sheetNames.includes(newSheetName)) {
      ui.alert("The sheet with " + newSheetName + " name already exists. Please choose a different name.", ui.ButtonSet.OK)
    }
    // Check if the given name has multiple words
    else if (newSheetName.split(" ").length>1) {
      newSheetName = newSheetName.split(" ").join("");
      if (sheetNames.includes(newSheetName)){
        ui.alert("The sheet with " + newSheetName + " name already exists. Please choose a different name.", ui.ButtonSet.OK)
        var newSheetName = "";
      }
      else {
        MasterTempFile.getSheetByName("MASTER").copyTo(ss).setName(newSheetName).showSheet();
      }
    }
    else {
      MasterTempFile.getSheetByName("MASTER").copyTo(ss).setName(newSheetName).showSheet();
    }
  };
  
  return newSheetName;
  
}

// ********* End: Copy Master Template
