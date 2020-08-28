function POSummary() {
  
  var PODetails = getPOfromMasterPPPMREQFile();
  
  // Add PO information to a separate PO summary sheet
  var POList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO List");
  POList.getRange(1, 1, 1, 3).setValues([["REQ#","PO Number","PO Issue Date"]]);
  POList.getRange(2, 1, PODetails.length, 3).setValues(PODetails);
  
}
