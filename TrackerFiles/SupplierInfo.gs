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
  var SupplierDataDump = [];
  //var SupplierData = [];
  
  // Store Supplier Info from all the sheets into Array
  
  // Get Sales Managers Data
  var SalesManagerSheet = masterSuppFile.getSheetByName(sheetNames[0]);
  var SalesLR = SalesManagerSheet.getLastRow();
  var SalesLC = SalesManagerSheet.getLastColumn();
  var SalesManagerData = SalesManagerSheet.getRange(2, 1, SalesLR, SalesLC).getDisplayValues();
  
  // Get Engg Managers Data
  var EnggManagerSheet = masterSuppFile.getSheetByName(sheetNames[1]);
  var EnggLR = EnggManagerSheet.getLastRow();
  var EnggLC = EnggManagerSheet.getLastColumn();
  var EnggManagerData = EnggManagerSheet.getRange(2, 1, EnggLR, EnggLC).getDisplayValues();
  
  
  for (let i= 0; i<sheetNames.length; i++) {
     
    SupplierDataDump.push(masterSuppFile.getSheetByName(sheetNames[i]).getDataRange().getValues());
  };
  
  //SupplierData.push([SalesManagerData, EnggManagerData]);
  
  //Logger.log(SupplierData);
  
  return [SalesManagerData, EnggManagerData];  
  
}


// ********* End: Retrieve Supplier Data from Master Supplier Contacts Database Google Sheet
