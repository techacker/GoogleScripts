//***************************************
//**** GOODS RECEIPT PROCESS SCRIPT *****
//****     AUTHOR: ANURAG BANSAL    *****
//****     Revision: 1.0.4          *****
//****     Date: 01/15/2021         *****
//****     Revision History         *****
//***************************************
//1.0.3 - Bug fix for GR Log - 01/21/2021
//1.0.4 - Bug fix for GR Log - 02/17/2021


//Run when spreadsheet loads
function onOpen(){
  createMenu();  
}

//Create menu dropdown
function createMenu() {  
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("PPPM Tools");
  menu.addItem("Compile GR Report","importCurrent");
  //menu.addItem("Update SharePoint PO Data","updateSharePointInfo");
  //menu.addItem("Replace Data in Current","replaceData");
  //menu.addItem("Update PO Log","compileLog");
  menu.addToUi();
}

function importCurrent() {
  
  const pastDueGRfileURL = "https://docs.google.com/spreadsheets/d/1cGCxZVipep-2C7MMucJTMHxJz9yJb65MXUJTDTcJmcg/";
  //const pastDueGRfileURL = "https://docs.google.com/spreadsheets/d/1bU7DqrLSFKRK4UBk53hAWDaCYkvzhL8XzbhEQLMhBPg/"; //Master Code File
  
  const ssGRFile = SpreadsheetApp.openByUrl(pastDueGRfileURL);
  
  // Check if "New" and "Current" sheets are there, if not, exit the program.
  if (!checkSheets_(ssGRFile)) {
    return false
  }
  
  const ssNew = ssGRFile.getSheetByName("New");
  const lrNew = ssNew.getLastRow();
  let lcNew = ssNew.getLastColumn();
  const lcNewHeader = ssNew.getRange(1, lcNew).getValue();
  
  // Add new columns with their header titles in "New" tab
  if (lcNewHeader === "Age Category") {
    ssNew.insertColumnsAfter(lcNew, 5);
    ssNew.getRange(1, lcNew+1, 1, 5).setValues([["Type","Sharepoint ID","Requestor","Status","Notes"]]);
    //let destination = ssNew.getRange(1, lcNew+1, 1,5);
    //ssNew.getRange(1,1,1,5).copyFormatToRange(destination, lcNew+1, lcNew+5, 1, 1);
    lcNew += 5;
  }
  
  const rangeNew = ssNew.getRange(2, 1, lrNew, lcNew).getDisplayValues();  
  
  // Get details from "Current" Tab
  const ssCurrent = ssGRFile.getSheetByName("Current");
  const lrCurrent = ssCurrent.getLastRow();
  const lcCurrent = ssCurrent.getLastColumn();
  
  const rangeCurrent = ssCurrent.getRange(2, 1, lrCurrent, lcCurrent).getDisplayValues(); 
  
  // Compare the two lists and update the data in "New" tab
  
  for (let i=0; i<rangeNew.length; i++) {
    for (let j=0; j<rangeCurrent.length; j++) {
      if (rangeNew[i][9] === rangeCurrent[j][9]) {
        rangeNew[i][23] = rangeCurrent[j][23];
        rangeNew[i][24] = rangeCurrent[j][24];
        rangeNew[i][25] = rangeCurrent[j][25];
        rangeNew[i][26] = rangeCurrent[j][26];
        rangeNew[i][27] = rangeCurrent[j][27];
      } 
    }
    // for unmatching results, mark them as "New"
    if (rangeNew[i][23] === "") {
      rangeNew[i][23] = "New";
    }
  }
  
  let updatedNewRange = [];
  
  // Only update the revised section
  for (let i=0; i<rangeNew.length-1; i++) {
    let rangetoCopy = rangeNew[i].slice(23, 28);
    updatedNewRange.push(rangetoCopy);
  }
  
  ssNew.getRange(2, 24, lrNew-1, 5).setValues(updatedNewRange);
  
  // Format the rows
  ssNew.getRange(2, 24, lrNew-1, 5).setBorder(null, true, true, true, true, true);
  ssNew.autoResizeColumns(24,4);
  
  updateSharePointInfo_(ssGRFile, rangeNew);
}

function updateSharePointInfo_(ssGRFile, rangeNew) {
  
  const ssNew = ssGRFile.getSheetByName("New");
  const lrNew = ssNew.getLastRow();
  
  const SharePointSheet = ssGRFile.getSheetByName("Sharepoint Data");
  const spRange = SharePointSheet.getRange(2, 1, SharePointSheet.getLastRow(), SharePointSheet.getLastColumn()).getDisplayValues();
  
  let newPOs = [];
  rangeNew.forEach(function(item) {
    if (item[23] === "New" && !newPOs.includes(item[0])) {
      newPOs.push(item[0])
    }
  })
  
  let updatedPOData = [];
  newPOs.forEach(function(newPO) {
    spRange.forEach(function(spRecord) {
      if (newPO === spRecord[0]) {
        updatedPOData.push([newPO,spRecord[1],spRecord[2]])
      }
    })
  })
  
  //Logger.log(updatedPOData); 
  //[[48830077, PPPM-PVP-24190, Whitehouse Trey (FCA)], 
  //[48830610, PPTR-INT- 25007, Torres Coronel Bernardo (FCA)], 
  //[48838290, PPPM-ALL-25886, Schubring James (FCA)]]
  
  let updatedNewPORange = [];
  
  // If the PO matches with data in SharePoint sheet, mark them as "SharePoint"
  for (let i=0; i<updatedPOData.length; i++) {
    for (let j=0; j<rangeNew.length; j++) {
      if (rangeNew[j][23] === "New" && updatedPOData[i][0] === rangeNew[j][0]) {
        rangeNew[j][23] = "Sharepoint";
        rangeNew[j][24] = updatedPOData[i][1];
        rangeNew[j][25] = updatedPOData[i][2];
      }
    }
  }
  
  //If there are still records that are new, mark them as "Event"
  for (let j=0; j<rangeNew.length; j++) {
    if (rangeNew[j][23] === "New") {  
      rangeNew[j][23] = "Event";
    }
  }
  
  // Only update the revised section
  for (let i=0; i<rangeNew.length-1; i++) {
    let rangetoCopy = rangeNew[i].slice(23, 26);
    updatedNewPORange.push(rangetoCopy);
  }
  
  // Update the records with SharePoint information
  ssNew.getRange(2, 24, lrNew-1, 3).setValues(updatedNewPORange);
  
  replaceData_(ssGRFile);
  
}

function replaceData_(ssGRFile) {
  
  const ssNew = ssGRFile.getSheetByName("New"); 
  let lrNew = ssNew.getLastRow();
  let lcNew = ssNew.getLastColumn();
  
  const ssCurrent = ssGRFile.getSheetByName("Current");
  let lrCurrent = ssCurrent.getLastRow();

  Logger.log(lrCurrent, lrNew);
  
  // Delete data from row 2 to end to keep formatting
  ssCurrent.deleteRows(2,lrCurrent-2);
  
  // Copy data from "New" tab to "Current" tab
  const destination = ssCurrent.getRange(2, 1);
  ssNew.getRange(2, 1, lrNew, lcNew).copyTo(destination);
  
  // Update Log
  compileLog_(ssGRFile);

  Browser.msgBox("The 'Current' sheet and 'GR Log' was updated.");
  
  // Delete "New" sheet
  ssGRFile.deleteSheet(ssNew);
}

function compileLog_(GRFile) {
    
  const ssLog = GRFile.getSheetByName("GR Log"); 
  const lrLog = ssLog.getLastRow();
  const lcLog = ssLog.getLastColumn();
  
  const ssCurrent = GRFile.getSheetByName("Current");
  const lrCurrent = ssCurrent.getLastRow();
  const lcCurrent = ssCurrent.getLastColumn();
  const rangeCurrent = ssCurrent.getRange(2,1,lrCurrent-1,lcCurrent).getValues();

  // Collect Engineer Names from Current Sheet

  let engNames = [];
  rangeCurrent.forEach(row => {
    if (!engNames.includes(row[12])) {
      engNames.push(row[12])
    }
  })
  
  // Gather PO count per name and collect corresponding BICEE

  let POLogData = []
  let today = new Date().toLocaleDateString();

  engNames.forEach(name => {
    let PONum = []
    let grpName = "";
    for (let i=0; i<rangeCurrent.length; i++) {
      if (name === rangeCurrent[i][12] && !PONum.includes(rangeCurrent[i][0])) {
        PONum.push(rangeCurrent[i][0]);
        grpName = rangeCurrent[i][11];
      }
    }
    // Insert "Name, BICEE, PO Count, Date" to array
    POLogData.push([name, grpName, PONum.length.toString(), today]);

    // Clear the array for next name
    PONum = [];
    groupName = "";
  })

  // Write data to 'GR Log' tab
  const lastRunDate = ssLog.getRange(2,4,1,1).getDisplayValue();
  
  // If last run date is same as today's date, don't update the log
  if (lastRunDate !== today) {
    ssLog.insertRows(2,POLogData.length);
    ssLog.getRange(2,1,POLogData.length,4).setValues(POLogData);
    ssLog.getRange(2,1,ssLog.getLastRow(),4).sort(1);
  }
  
}

function checkSheets_(GRFile) {
  
  let sheetNames = [];
  const numSheets = GRFile.getNumSheets();
  GRFile.getSheets().forEach(function(sheet) {
    let name = sheet.getSheetName();
    sheetNames.push(name);
  })
  
  if (!sheetNames.includes("New")) {
    Browser.msgBox("Sheet named 'New' is not found!");
    return false
  } else if (!sheetNames.includes("Current")) {
    Browser.msgBox("Sheet named 'Current' is not found!");
    return false
  }

}

