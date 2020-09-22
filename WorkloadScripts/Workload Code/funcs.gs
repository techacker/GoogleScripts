function addNewRow(rowData) {

//Add new row based on inputs
  const currentDate = new Date();  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Events");
  ws.appendRow([
    currentDate,
    rowData.modelYear,
    rowData.vehicleFamily,
    rowData.eventName,
    rowData.eventType,
    rowData.eventTitle,
    rowData.origMRD,
    rowData.eventStatus,
    rowData.programManager,
    rowData.comments,
    rowData.initPNs    
    ]);       

//Format new row
  var as = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = as.getLastRow();
  var lc = as.getLastColumn();
  as.getRange(lr, 1, 1, lc).activate();
  as.getRange(lr-1, 1, 1, lc).copyTo(as.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

//Function Complete (for success handler)
  return true;  
}


function getProgMgrDDA(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Summary");
  const PMvals = ws.getRange(3, 1, 3, 1).getValues();
  return PMvals;
}

function getEventTypeDDA(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Summary");
  const ETvals = ws.getRange(39, 1, 7, 1).getValues();
  return ETvals;
}

function testFunction(){
  var testVal = "test";
          var testData = {
          testVal};
  var output = testData.testVal;
  Logger.log(output);


}

