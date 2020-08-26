// ********* Start: Retrieve Supplier Data from Master Supplier Contacts Database Google Sheet

function getMasterSupplierData() {
  
  // Master Supplier Contacts File: ID - 1rx2kKuDqgE9gjyJ2UrZOdXPYneJ4zFXTv_icexl_154
  
  var masterSuppFile = SpreadsheetApp.openById("1rx2kKuDqgE9gjyJ2UrZOdXPYneJ4zFXTv_icexl_154");
  var fileName = masterSuppFile.getName();
  var numsheets = masterSuppFile.getSheets();
  var sheets = {};
  
  // To go through number of sheets in Master File in case few other sheets are added in future
  for (let i=0; i<numsheets.length; i++) {
    sheets[numsheets[i].getName()] = (numsheets[i].getSheetId()).toString();
  };
    
  var sheetNames = Object.keys(sheets);
  var sheetIDs = Object.values(sheets);
  var SupplierData = [];
  
  // Store Supplier Info from all the sheets into Array
  for (let i= 0; i<sheetNames.length; i++) {
    SupplierData.push(masterSuppFile.getSheetByName(sheetNames[i]).getDataRange().getValues());
  };
  
  return SupplierData;  
  
}


// ********* End: Retrieve Supplier Data from Master Supplier Contacts Database Google Sheet
