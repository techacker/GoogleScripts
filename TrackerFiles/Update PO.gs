// ********* Start: Update PO in Tracker Sheet

function UpdatePO(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkTemplate() function to get the required headerRow and column indexes.
  var headerRow = checkTemplate()[0];  
  var indices = checkTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
    
  // Get the index values of all necessary columns for Reqs.
  var REQInd = indices[10]+1;
  var PONumInd = indices[11]+1;
  var POIssueDate = indices[12]+1;
  
  var REQNum = ss.getRange(headerRow+1, REQInd, lr, 1).getDisplayValues();
  
  // Get PO details from Master File using PO Update Function
  var POs = getPOfromMasterPPPMREQFile();
  
  // Update POs in the tracker sheet with REQ Nos.
  for (var i=0; i<lr-headerRow; i++) {
    for (var j=0; j<POs.length; j++) {
      if (REQNum[i][0] === POs[j][0]) {
        // If no PO has been issued - Status = "Pending"
        if (POs[j][1] === "") {
          ss.getRange(i + (headerRow + 1), PONumInd, 1, 1).setValue("Pending");
          ss.getRange(i + (headerRow + 1), POIssueDate, 1, 1).setValue("Waiting");
        } 
        // If PO is issued, update PO details
        else {
          ss.getRange(i + (headerRow + 1), PONumInd, 1, 1).setValue(POs[j][1]);
          ss.getRange(i + (headerRow + 1), POIssueDate, 1, 1).setValue(POs[j][2]);
        };
      };
    };
  };
  
  Browser.msgBox("All POs released to the suppliers are updated in the current sheet!");
  
}; 

// ********* Start: Update PO in Tracker Sheet



// ********* Start: Get Requisition Info

function GetReqNo(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkTemplate() function to get the required headerRow and column indexes.
  var headerRow = checkTemplate()[0];  
  var indices = checkTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
  
  // Logger.log(indices);
  // indices = [1.0, 2.0, 4.0, 5.0, 9.0, 11.0, 13.0, 14.0, 17.0, 0.0, 15.0, 16.0, 19.0]
  
  // Get the index values of all necessary columns for Reqs.
 
  var REQInd = indices[10]+1;
  var PONumInd = indices[11]+1;
  var REQs = [];
  
  // Get the array of values for these indexes
  
  var REQNum = ss.getRange(headerRow+1, REQInd, lr, 1).getDisplayValues();
  var PO = ss.getRange(headerRow+1, PONumInd, lr, 1).getDisplayValues();
  
  for (var i=0; i<range.length; i++) {
    if (REQNum[i][0].length >= 6 && (PO[i][0] !== "Pending" || PO[i][0] !== "Processing")) {
      REQs.push(REQNum[i][0]);
    };
  };
  
  return REQs;
  
}; 

// ********* End: Get Requisition Info
