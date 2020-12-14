function addNewRow(rowData) {
  
  //Add new row based on inputs
  const currentDate = new Date();  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Events");
  ws.appendRow([
    rowData.vehicleFamily,
    rowData.modelYear,
    rowData.eventName,
    rowData.eventType,
    rowData.eventTitle,
    rowData.origMRD,
    rowData.eventStatus,
    rowData.programManager,
    rowData.requestor,
    rowData.wbs,
    rowData.location,
    rowData.shipcode,
    rowData.shipadd,
    rowData.attn,
    rowData.comments,
    "", "", "", "", "", "", "", "",
    "", "", "", "", "", "", "", "",
    "", "", "", "", "", "",
    currentDate,
    rowData.initPNs
    
  ]);       
  
  formatRowCreateTracker();
  //Function Complete (for success handler)
  return true;  
}

function formatRowCreateTracker() {
 
  //Format new row
  var as = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = as.getLastRow();
  var lc = as.getLastColumn();
  as.getRange(lr, 1, 1, lc).activate();
  as.getRange(lr-1, 16, 1, 22).copyTo(as.getRange(lr, 16, 1, 22), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  as.getRange(lr-1, 1, 1, lc).copyTo(as.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  
  getNewTrackerURL();
  
  Browser.msgBox("Success!", "A new Tracker was created. Look for its URL in Tracker URL field.", Browser.Buttons.OK);
  
}


function getProgMgrDDA(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Summary");
  const PMvals = ws.getRange(3, 1, 6, 1).getValues();
  return PMvals;
}

function getEventTypeDDA(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Summary");
  const ETvals = ws.getRange(32, 1, 10, 1).getValues();
  return ETvals;
}
    
