function updateBOMFile_(BuildObj, parentFileID) {

  /*
  let BuildObj = {'Donor Order Pulse - 14wk (Plan)':null, 'Vehicle Quantity':'25', 'Noun Name Delivered - 30wk (Plan)':null, 'Weeks to Build Start':null, 'Build Title':'2026 XX Mule 9', 'Weeks to MRD':null, 'X-Flag Due Date':null, 'Vehicle QTY Published - 26wk (Pulse)':null, 'Date Created':'8/9/2021', 'Overall Build Pulse':null, 'Build Phase Funding Approved - 19wk (Actual)':null, 'Build Lead':'Anurag Bansal', 'Build Plan Complete - 16wk (Plan)':null, 'Build Phase':'9', 'Part Delivery Pulse':null, 'Build % REQ':null, 'Vehicle Family':'XX', 'Carrier Specs Published - 18wk (Pulse)':null, 'Bridge Design Complete - 12wk (Plan)':null, 'Carrier Specs Published - 18wk (Actual)':null, 'Program Scope':"", 'Build % Received':null, 'Tree FC#':"", 'Last Modified':'8/9/2021', 'Build Plan Complete - 16wk (Pulse)':null, 'Model Year':'2026', 'Noun Name Delivered - 30wk (Actual)':null, 'Bridge Design Complete - 12wk (Pulse)':null, 'Noun Name Delivered - 30wk (Pulse)':null, 'Build Shop Readiness':null, 'Build Status':'Active', 'Build Cost':null, 'Carrier Specs Published - 18wk (Plan)':null, 'Work Scope - 32wk (Actual)':null, 'Build Phase Funding Approved - 19wk (Plan)':null, 'Donor Order Pulse - 14wk (Pulse)':null, 'Build % PO':null, 'Bridge Design Complete - 12wk (Actual)':null, 'Final BOM Due Date':null, 'Build Location':'PVO', 'Build Phase Funding Approved - 19wk (Pulse)':null, 'Donor Order Pulse - 14wk (Actual)':null, 'Carrier':"", 'Build Status Comments':null, 'Process Sheet Due Date':null, 'Vehicle QTY Published - 26wk (Actual)':null, 'Work Scope - 32wk (Plan)':null, 'BOM Review Date':null, 'Build Strategy':"", 'Build Plan Complete - 16wk (Actual)':null, 'WBS Code':"", 'TCF Build Start':'9/20/2021', 'Build Total # of Parts':null, 'TCF MRD':'9/6/2021', 'Build Type':'Mule', 'Work Scope - 32wk (Pulse)':null, 'Build Tracker URL':'https://docs.google.com/spreadsheets/d/18iQwymYo5FXN0K3vqDlu_ns_p76J_vhQJGkGVPNWfhA/', 'Vehicle QTY Published - 26wk (Plan)':null}
  */
  
  let parentFile = SpreadsheetApp.openById(parentFileID)
  let sheetNames = parentFile.getSheets().map(sheets => sheets.getSheetName())

  if (sheetNames.includes('BOM Summary')) {
    addNewDataRow_(parentFileID,'BOM Summary','Build Tracker URL',BuildObj)
    if (sheetNames.includes('Settings')) {
      let settingsSheet = parentFile.getSheetByName('Settings')
      let reqdNames = settingsSheet.getRange(2,1,settingsSheet.getLastRow()-1,1).getValues();
      reqdNames.forEach(names => {
        if (sheetNames.includes(names[0])) {
          addNewDataRow_(parentFileID,names[0],'Build Title',BuildObj)
        }
      })
    }
  } 
  else {
    return false;
  }
}
