/*
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
  //return true;  
}

function formatRowCreateTracker() {
    
  //Format new row
  var as = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Events");
  var lr = as.getLastRow();
  var lc = as.getLastColumn();
  as.getRange(lr, 1, 1, lc).activate();
  as.getRange(lr-1, 16, 1, 22).copyTo(as.getRange(lr, 16, 1, 22), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  as.getRange(lr-1, 1, 1, lc).copyTo(as.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  
  getNewTrackerURL(workloadfileURL);
  
  Browser.msgBox("Success", "Tracker was created and URL was updated.", Browser.Buttons.OK);

  
}

/*
function getProgMgrDDA(){
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/";  // Test PPPM Workload File
  //var ws = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Summary");
  //const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const ws = ss.getSheetByName("Summary");
  //const PMvals = ws.getRange(3, 1, 6, 1).getValues();
  //Logger.log(PMvals);
  const PMvals = [['Kelli Rodenbo'],['Mark Ballo'],['UNASSIGNED']];
  return PMvals;
}

function getEventTypeDDA(){
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/";  // Test PPPM Workload File
  //var ws = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Summary");
  //const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const ws = ss.getSheetByName("Summary");
  //const ETvals = ws.getRange(32, 1, 10, 1).getValues();
  const ETvals = [['PWB'],['QAF'],['DVPR'],['MEDIA'],['TTO'],['Safety'],['Dyno'],['Other']];
  //Logger.log(ETvals);
  return ETvals;
}

function getEventTitle(){
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lr = ss.getLastRow();
  //const ws = ss.getSheetByName("Summary");
  const EventTitles = ss.getRange(2, 5, lr-1, 1).getValues();
  return EventTitles;
}


function updateRow(rowData) {
 
  //Update row based on inputs  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Events");
  //ws.getRange(row, column, numRows, numColumns).setValues(values);
  /*
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
  
  //Function Complete (for success handler)
  return true;  
}
*/
