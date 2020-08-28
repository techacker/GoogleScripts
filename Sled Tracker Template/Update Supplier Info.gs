// ********* Start: Update Supplier Info in Sled Series Tracker Sheet

function UpdateSupplierInfo(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkTemplate() function to get the required headerRow and column indexes.
  var headerRow = checkSledTemplate()[0];  
  var indices = checkSledTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
  
  var SupplierData = getMasterSupplierData();
  var SalesManagersData = SupplierData[0];
  var EnggManagersData = SupplierData[1];
     
  // Get the index values of all necessary columns for Supplier Info.
  // Logger.log(indices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 9.0, 11.0, 13.0, 16.0, 18.0, 21.0, 23.0, 26.0, 28.0, 31.0, 33.0, 36.0, 38.0, 41.0, 43.0, 46.0, 48.0, 51.0, 53.0, 56.0, 60.0] 
  
  var SCodeInd = indices[1]+1;
  var CompInd = indices[2]+1;
  var SuppNameInd = indices[3]+1;
  var SEmailInd = indices[4]+1;
  var AEmailInd = indices[5]+1;
  var PhoneInd = indices[6]+1;
  
  
  var SuppCodes = ss.getRange(headerRow+1, SCodeInd, lr, 1).getDisplayValues();
  var CompanyNames = ss.getRange(headerRow+1, CompInd, lr, 1).getDisplayValues();
  var SalesMngrEmails = ss.getRange(headerRow+1, SEmailInd, lr, 1).getDisplayValues();
  var EnggMngrEmails = ss.getRange(headerRow+1, AEmailInd, lr, 1).getDisplayValues();
  
  // Logger.log(SalesManagersData[0]) &  Logger.log(EnggManagersData[0]) returns ->
  // [Supplier Code, Region/Sector, Supplier Name, Role Name, Contact First Name, Contact Last Name, Phone, Mobile, Email, Created Date, Created By, Updated Date, Updated By]
 
  
  // Update Sales Manager Details in the tracker sheet.
  for (var i=0; i<lr-headerRow; i++) {
    for (var j=0; j<SalesManagersData.length; j++) {
      if (SalesMngrEmails[i][0] === "" && SuppCodes[i][0] === SalesManagersData[j][0]) {
        ss.getRange(i + (headerRow + 1), CompInd, 1, 1).setValue(SalesManagersData[j][2]);
        ss.getRange(i + (headerRow + 1), SuppNameInd, 1, 1).setValue(SalesManagersData[j][4] + " " + SalesManagersData[j][5]);
        ss.getRange(i + (headerRow + 1), SEmailInd, 1, 1).setValue(SalesManagersData[j][8]);
        ss.getRange(i + (headerRow + 1), PhoneInd, 1, 1).setValue(SalesManagersData[j][6]);
      };
    };
  };
  
  
  // Update Engg Manager Details in the tracker sheet.
  for (var i=0; i<lr-headerRow; i++) {
    for (var j=0; j<EnggManagersData.length; j++) {
      // Only if Company Names was not detected in above Sales Mangers lookup
      if (SalesMngrEmails[i][0] === "" && EnggMngrEmails[i][0] === "" && SuppCodes[i][0] === EnggManagersData[j][0]) {
        ss.getRange(i + (headerRow + 1), CompInd, 1, 1).setValue(EnggManagersData[j][2]);
        //ss.getRange(i + (headerRow + 1), SuppNameInd, 1, 1).setValue(EnggManagersData[j][4] + " " + EnggManagersData[j][5]);
        ss.getRange(i + (headerRow + 1), AEmailInd, 1, 1).setValue(EnggManagersData[j][8]);
        //ss.getRange(i + (headerRow + 1), PhoneInd, 1, 1).setValue(EnggManagersData[j][6]);
      } else if (SuppCodes[i][0] === EnggManagersData[j][0]) {
        ss.getRange(i + (headerRow + 1), AEmailInd, 1, 1).setValue(EnggManagersData[j][8]);
      }; 
    };
  };
  
  Browser.msgBox("Supplier Info has been updated in the current sheet based on the data from GYPSIS!");
  
}; 

// ********* End: Update Supplier Info in Tracker Sheet

