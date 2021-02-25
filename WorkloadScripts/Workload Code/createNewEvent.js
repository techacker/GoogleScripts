function createTracker() {

  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/";  // Test PPPM Workload File
  var neSheet = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Add New Event");
  var lr = neSheet.getLastRow();
  var lc = neSheet.getLastColumn();
  var neValues = neSheet.getRange(2,2,lr-1,1).getDisplayValues();
  
  var eventData = {
    modelYear: neValues[0],
    vehicleFamily: neValues[1],
    eventName: neValues[2],
    eventType:neValues[4],
    //eventTitle: 
    origMRD: neValues[3],
    //eventStatus,
    programManager: neValues[5],
    comments: neValues[6],
    initPNs: neValues[8], 
    requestor: neValues[9],
    wbs: neValues[7],
    location: neValues[10],
    attn: neValues[11],
    shipcode: neValues[12],
    shipadd: neValues[13]
    };

  // Check if mandatory fields are not blank

  if (eventData.modelYear[0] !=="" && eventData.vehicleFamily[0] !== "" && eventData.eventName[0] !== "" && eventData.origMRD[0] !== "" && eventData.eventType[0] !== "" && eventData.programManager[0] !== "") {
    addEvent(eventData);
    var blankData = [[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""]];
    neSheet.getRange(2,2,lr-1,1).setValues(blankData);
    formatRowCreateTracker(workloadfileURL);
  } else {
    Browser.msgBox("Alert", "Mandatory information to create tracker is missing.", Browser.Buttons.OK);  
  }
  
}

function addEvent(eventData) {

  //Add new row based on inputs
  const currentDate = new Date();  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Events");

  ws.appendRow([
    eventData.vehicleFamily[0],
    eventData.modelYear[0],
    eventData.eventName[0],
    eventData.eventType[0],
    eventData.modelYear[0]+" "+eventData.vehicleFamily[0].toUpperCase()+" "+eventData.eventName[0],
    eventData.origMRD[0],
    "In-Process",
    eventData.programManager[0],
    eventData.requestor[0],
    eventData.wbs[0],
    eventData.location[0],
    eventData.shipcode[0],
    eventData.shipadd[0],
    eventData.attn[0],
    eventData.comments[0],
    "", "", "", "", "", "", "", "",
    "", "", "", "", "", "", "", "",
    "", "", "", "", "", "",
    currentDate,
    eventData.initPNs[0]
  ]);  

}

function formatRowCreateTracker(workloadfileURL) {
    
  //Format new row
  var as = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Events");
  var lr = as.getLastRow();
  var lc = as.getLastColumn();
  as.getRange(lr, 1, 1, lc).activate();
  as.getRange(lr-1, 16, 1, 22).copyTo(as.getRange(lr, 16, 1, 22), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  as.getRange(lr-1, 1, 1, lc).copyTo(as.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  
  getNewTrackerURL(workloadfileURL);
  Browser.msgBox("Success", "A new tracker has been created.", Browser.Buttons.OK);
  
}
