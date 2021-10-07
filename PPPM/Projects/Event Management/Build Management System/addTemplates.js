// ********* Start: Main Function to add new row in Summary Sheet that calls 'Copy Master Template function'

function getTemplate(tempID,biceepName) {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Build Summary");
  const lc = summarySheet.getLastColumn();
  let lr = summarySheet.getLastRow();
  let range = summarySheet.getRange(1, 1, lr, lc).getDisplayValues();
  let tabRowNum = 0;
  const engName = getUserName()

  // Get Event Name so it can be copied to new sheet
  const eventName = summarySheet.getRange(2, 2).getDisplayValues()[0][0];
  
  // Get "Header" Row number
  range.filter((item, rowInd) => {
    item.filter(colname => {
      if (colname === "Designs-G-#") {
        tabRowNum = rowInd + 1
      }
    })
  })

  const defaultRow = summarySheet.getRange(tabRowNum+2, 1, 1, lc);
  let formulas = summarySheet.getRange(tabRowNum+2, 1, 1, lc).getFormulas();
  const returnedValues = copySpecifiedTemplate(tempID,biceepName)

  const newSheetName = returnedValues[0];
  const system = returnedValues[1]
  
  if (newSheetName.length !== 0) {
    // Unhide all hidden rows
    summarySheet.unhideRow(summarySheet.getRange(1, 1, lr, lc));
    
    // Shift one row down
    defaultRow.insertCells(SpreadsheetApp.Dimension.ROWS);
    
    // Set first two cell value to be the sheet name & biceeps
    summarySheet.getRange(tabRowNum+2, 1).setValue(biceepName);
    summarySheet.getRange(tabRowNum+2, 2).setValue(system);
    lr++;
    
    // Set formulas for other cells in this new row.
    for (let i=0; i<28; i++) {
      let splits = formulas[0][i].split('!');  // Since formula is referenced with '!'
      summarySheet.getRange(tabRowNum+2, i+1).setValue(`='${newSheetName}'!${splits[1]}`);
    }
    for (let i=28; i<formulas[0].length; i++) {
      summarySheet.getRange(tabRowNum+2, i+1).setFormula(formulas[0][i]);
    }
  };
  
  summarySheet.hideRows(tabRowNum);
  
  // Get "DEFAULT" Row number
  range.filter((item, rowInd) => {
    item.filter(colname => {
      if (colname.toUpperCase() === "DEFAULT") {
        tabRowNum = rowInd + 2
        summarySheet.hideRows(tabRowNum);
        ss.getSheetByName("DEFAULT").hideSheet();
      }
    })
  })

  const sheetNames = getSheetNames_();
  
  if (sheetNames.includes("DEFAULT")) {
    ss.getSheetByName("DEFAULT").hideSheet();
  }
  
  if (sheetNames.includes(newSheetName)) {
    let newSheet = ss.getSheetByName(newSheetName)
    for (let i=0; i<13; i++) {
      newSheet.getRange(i+1, 3).setFormula(`=VLOOKUP(A${i+1},'Build Summary'!$A:$B,2,0)`)
    }
    newSheet.getRange(14, 3).setValue(biceepName)
    newSheet.getRange(15, 3).setValue(system)
    newSheet.getRange(16, 3).setValue(engName)
  }
  
}

// ********* End: Main Function to add new row in Summary Sheet that calls 'Copy Master Template function'



// ********* Start: Copy Master Template

function copySpecifiedTemplate(tempID,biceepName) {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetNames = getSheetNames_();
  let sysName = ""
  let newSheetName = ""
  let sysRes = ui.prompt("Info Required", "Name of the sub-system this sheet is for?", ui.ButtonSet.OK_CANCEL);

  const masterTempFile = SpreadsheetApp.openById(tempID);

  //Confirm user input
  if (sysRes.getSelectedButton() == ui.Button.CANCEL || sysRes.getSelectedButton() == ui.Button.CLOSE) {
    ui.alert("No sheet was created");
  }
  else {
    sysName = sysRes.getResponseText();
    newSheetName = biceepName + " - " + sysName
    // Check if the sheet exists, else copy the master with the given name
    if (sheetNames.includes(newSheetName)) {
      ui.alert("The sheet with " + newSheetName + " name already exists. Please choose a different name.", ui.ButtonSet.OK)
      newSheetName = "";
    }
    else {
      masterTempFile.getSheetByName("DEFAULT").copyTo(ss).setName(newSheetName).showSheet();
    }
  };
  
  return [newSheetName, sysName];
  
}

// ********* End: Copy Master Template
