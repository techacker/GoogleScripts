function copyRFQForm(destination, suppcode) {
  
  // Master RFQ Form: ID - 1jfr3hYIL7oWdePtXF1WE7CFOwyhsS3dKZBPCCyLImKg/edit#gid=1831052769
  
  const RFQFile = SpreadsheetApp.openById("1jfr3hYIL7oWdePtXF1WE7CFOwyhsS3dKZBPCCyLImKg");
  const lr = RFQFile.getLastRow();
  const lc = RFQFile.getLastColumn();
  
  // Copy RFQ Form sheet to the tracker file and rename the sheet to 'Supplier Code'
  RFQFile.getSheetByName("RFQ Form").copyTo(destination).setName(suppcode);
  
}

function createRFQForms() {
  
  var destination = SpreadsheetApp.getActiveSpreadsheet();
  var SuppCodes = UpdateSupplierInfo();
  
  for (var i=0; i<5; i++) { // Replace 5 with "SuppCodes.length"
    copyRFQForm(destination, SuppCodes[i]);
  }
}
