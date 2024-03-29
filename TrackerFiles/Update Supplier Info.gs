// ********* Start: Update Supplier Info in Tracker Sheet

function UpdateSupplierInfo(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkTemplate() function to get the required headerRow and column indexes.
  var headerRow = checkTemplate()[0];  
  var indices = checkTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
  
  var SupplierData = getMasterSupplierData();
  var SalesManagersData = SupplierData[0];
  var EnggManagersData = SupplierData[1];
     
  // Get the index values of all necessary columns for Supplier Info.
  // Logger.log(indices);
  // [1.0, 2.0, 4.0, 5.0, 9.0, 11.0, 13.0, 14.0, 17.0, 0.0, 15.0, 16.0, 19.0, 6.0]
  
  
  var SCodeInd = indices[0]+1;
  var CompInd = indices[1]+1;
  var SEmailInd = indices[2]+1;
  var AEmailInd = indices[3]+1;
  var PhoneInd = indices[13]+1;
  var SuppNameInd = indices[14]+1;
  
  var SuppCodes = ss.getRange(headerRow+1, SCodeInd, lr, 1).getDisplayValues();
  var CompanyNames = ss.getRange(headerRow+1, CompInd, lr, 1).getDisplayValues();
  var SalesMngrEmails = ss.getRange(headerRow+1, SEmailInd, lr, 1).getDisplayValues();
  var EnggMngrEmails = ss.getRange(headerRow+1, AEmailInd, lr, 1).getDisplayValues();
  
  var uniqueSupplierCodes = [];
  
  for (var i=0; i < SuppCodes.length; i++) {
    if (uniqueSupplierCodes.indexOf(SuppCodes[i][0].trim()) === -1.0 && SuppCodes[i][0].trim() !== "") {
      uniqueSupplierCodes.push(SuppCodes[i][0].trim());
    };
  };
  
  // Collect Sales Managers info for Unique Supplier Codes
  var SalesManagers = [];
  for (var i=0; i<uniqueSupplierCodes.length; i++) {
    for (var j=0; j<SalesManagersData.length; j++) {
      if (uniqueSupplierCodes[i] === SalesManagersData[j][0].toString()) {
        SalesManagers.push([SalesManagersData[j][0].toString(), SalesManagersData[j][2],SalesManagersData[j][4] + " " + SalesManagersData[j][5],
                           SalesManagersData[j][8], SalesManagersData[j][6]]);
      }
    }
  }
  
  // Collect Engg Managers info for Unique Supplier Codes
  var EnggManagers = [];
  for (var i=0; i<uniqueSupplierCodes.length; i++) {
    for (var j=0; j<EnggManagersData.length; j++) {
      if (uniqueSupplierCodes[i] === EnggManagersData[j][0].toString()) {
        EnggManagers.push([EnggManagersData[j][0].toString(), EnggManagersData[j][2],EnggManagersData[j][4] + " " + EnggManagersData[j][5],
                           EnggManagersData[j][8], EnggManagersData[j][6]]);
      }
    }
  }
  
  // [[12624, LS Automotive Technologies Wuxi Cor, HS Park, hsparkwuxi@lsautomotive.com, 704467752], 
  //  [28242, Bitron De Mexico SA DE CV, Juan Chavez, juan.chavez@bitron.com.mx, 4421011600], 
  //  [28242, Bitron De Mexico SA DE CV, Noemi Rodriguez, noemi.rodriguez@mex.bitron-ind.com, 4421011614-1614], 
  //  [59305 B, TA YIH INDUSTRIAL CO LTD, CHIN-WEN CHEN, jchen@tayih-ind.com.tw, 88662615151-278]] 
  
  
  // Update Sales Manager Details in the tracker sheet.
  for (var i=0; i<lr-headerRow; i++) {
    for (var j=0; j<SalesManagers.length; j++) {
      if (SuppCodes[i][0].trim() !== "" && SalesMngrEmails[i][0] === "" && SuppCodes[i][0].trim() === SalesManagers[j][0]) {
        ss.getRange(i + (headerRow + 1), CompInd, 1, 1).setValue(SalesManagers[j][1]);
        ss.getRange(i + (headerRow + 1), SuppNameInd, 1, 1).setValue(SalesManagers[j][2]);
        ss.getRange(i + (headerRow + 1), SEmailInd, 1, 1).setValue(SalesManagers[j][3]);
        ss.getRange(i + (headerRow + 1), PhoneInd, 1, 1).setValue(SalesManagers[j][4]);
      };
    };
  };
  
  
  // Update Engg Manager Details in the tracker sheet.
  for (var i=0; i<lr-headerRow; i++) {
    for (var j=0; j<EnggManagers.length; j++) {
      // Only if Company Names was not detected in above Sales Mangers lookup
      if (SalesMngrEmails[i][0] === "" && EnggMngrEmails[i][0] === "" && SuppCodes[i][0].trim() === EnggManagers[j][0].toString()) {
        ss.getRange(i + (headerRow + 1), CompInd, 1, 1).setValue(EnggManagers[j][1]);
        ss.getRange(i + (headerRow + 1), SuppNameInd, 1, 1).setValue(EnggManagers[j][2]);
        ss.getRange(i + (headerRow + 1), AEmailInd, 1, 1).setValue(EnggManagers[j][3]);
        ss.getRange(i + (headerRow + 1), PhoneInd, 1, 1).setValue(EnggManagers[j][4]);
      } 
      /*
      else if (SuppCodes[i][0] === EnggManagers[j][0]) {
        ss.getRange(i + (headerRow + 1), AEmailInd, 1, 1).setValue(EnggManagers[j][3]);
      };
      */
    };
  };
  
  //Browser.msgBox("Supplier Info has been updated in the current sheet based on the data from GYPSIS!");
  
  return uniqueSupplierCodes;
  
}; 

// ********* End: Update Supplier Info in Tracker Sheet
