function getUniqueSupplierCodes() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  var headerRow = checkTemplate()[0];
  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
  
  var indices = checkTemplate()[1]; 
  var SCodeInd = indices[0]+1;
  
  var SCodes = ss.getRange(headerRow+1, SCodeInd, lr, 1).getDisplayValues();
  var uniqueSupplierCodes = [];
  
  for (var i=0; i < SCodes.length; i++) {
    if (uniqueSupplierCodes.indexOf(SCodes[i][0]) === -1.0 && SCodes[i][0] !== "") {
      uniqueSupplierCodes.push(SCodes[i][0].toString());
    };
  };
  
  return uniqueSupplierCodes;
  
}


function getRFQForms() {
  
  const RFQFile = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1jfr3hYIL7oWdePtXF1WE7CFOwyhsS3dKZBPCCyLImKg/");
  // Master RFQ Form: ID - 1jfr3hYIL7oWdePtXF1WE7CFOwyhsS3dKZBPCCyLImKg/edit#gid=1831052769
  // const RFQFile = SpreadsheetApp.openById("1jfr3hYIL7oWdePtXF1WE7CFOwyhsS3dKZBPCCyLImKg");
  
  // Destination is current active sheet
  var destination = SpreadsheetApp.getActiveSpreadsheet(); 
  var existingSheetNames = getSheetNames();
  
  // Get unique supplier codes
  var SuppCodes = getUniqueSupplierCodes();
  var RFQFormNames = [];
  
  for (var i=0; i<SuppCodes.length; i++) { 
    if (existingSheetNames.indexOf(SuppCodes[i]) === -1) {
      RFQFile.getSheetByName("RFQ Form").copyTo(destination).setName(SuppCodes[i]);
      RFQFormNames.push(SuppCodes[i]);
    } else {
      RFQFormNames.push(SuppCodes[i]);
    }
  }
  
  return RFQFormNames;
}

function getUserInfo(info) {
  var ui = SpreadsheetApp.getUi();
  var Response = ui.prompt("Information Required", "Enter your " + info, ui.ButtonSet.OK_CANCEL); 
  
  //Confirm user input
  if (Response.getSelectedButton() == ui.Button.CANCEL || Response.getSelectedButton() == ui.Button.CLOSE) {
    var requestorInfo = "";
    ui.alert("No info was provided. Requestor " + info + " will be blank in RFQ Forms");
  }
  else {
    var requestorInfo = Response.getResponseText();
  }
  
  return requestorInfo;
}


// ********* Start: Draft Email Function

function createRFQForms(){
  
  var name = getUserInfo("name");
  var TID = getUserInfo("TID");
  var phone = getUserInfo("phone");
  var today = new Date();
  
  var RFQFormNames = getRFQForms();
  
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  var EventTitle = summarySheet.getRange(1, 3).getValue();
  var MYProgram = EventTitle.split(" ");
  var MY = MYProgram[0];
  var VF = MYProgram[1];
  
  var ShiptoCode = summarySheet.getRange(5, 3).getDisplayValue();
  var attn = summarySheet.getRange(7, 3).getValue();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  var headerRow = checkTemplate()[0];
  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
  var indices = checkTemplate()[1];  
  
  // Get the index values of all necessary columns for Email.
 
  var SCodeInd = indices[0]+1;
  var CompInd = indices[1]+1;
  var PNInd = indices[4]+1;
  var PDescInd = indices[5]+1;
  var QtyInd = indices[6]+1;
  var ShipInd = indices[7]+1;
  var mrdInd = indices[8]+1; 
  
  // Get the array of values for these indexes
  
  var SCodes = ss.getRange(headerRow+1, SCodeInd, lr, 1).getDisplayValues();
  var CompNames = ss.getRange(headerRow+1, CompInd, lr, 1).getValues();
  var PN = ss.getRange(headerRow+1, PNInd, lr, 1).getValues();
  var PDesc = ss.getRange(headerRow+1, PDescInd, lr, 1).getValues();
  var Qty = ss.getRange(headerRow+1, QtyInd, lr, 1).getValues();
  var Ship = ss.getRange(headerRow+1, ShipInd, lr, 1).getValues();
  var mrd = ss.getRange(headerRow+1, mrdInd, lr, 1).getDisplayValues();
  
  //Constant Info
  var PartDetails = [];
  var uniqueSupplierCodes = getUniqueSupplierCodes();
  
  for (var j=0; j<uniqueSupplierCodes.length; j++) {
    for (var i=0; i<lr-headerRow; i++) {
      if (SCodes[i][0] === uniqueSupplierCodes[j]) {
        var reqdDate = mrd[i][0];
        var shipAddress = Ship[i][0];
        // Add part number, description and qty to PartDetails array
        PartDetails.push([SCodes[i][0], PN[i][0], PDesc[i][0], Qty[i][0], mrd[i][0], Ship[i][0]]);
      }
    }
  }
  
  
  // Fill the form with appropriate information
  for (var j=0; j<PartDetails.length; j++) {
    var suppCode = PartDetails[j][0];
    for (var i=0; i<RFQFormNames.length; i++) {
      if (RFQFormNames[i] === suppCode) {
        var RFQForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RFQFormNames[i]);
        var RFQlc = RFQForm.getLastColumn();
        
        // Set Generic Event Info
        RFQForm.getRange(12, 2).setValue(today); // Today's Date
        RFQForm.getRange(12, 8).setValue(MY); // Model Year
        RFQForm.getRange(13, 2).setValue(name); // Requestor
        RFQForm.getRange(13, 8).setValue(VF); // Program
        RFQForm.getRange(14, 2).setValue(TID); // Program
        RFQForm.getRange(15, 2).setValue(phone); // Program
        RFQForm.getRange(16, 8).setValue(PartDetails[j][4]); // MRD
        RFQForm.getRange(20, 2).setValue(ShiptoCode); // Ship to Code
        RFQForm.getRange(21, 2).setValue(attn); // Attention to
        RFQForm.getRange(22, 2).setValue(PartDetails[j][5]); // Ship to Code
        RFQForm.getRange(22, 2).setVerticalAlignment("middle");
        
        // Fill part information
        RFQForm.getRange(26, 1, 1, RFQlc).insertCells(SpreadsheetApp.Dimension.ROWS);
        RFQForm.getRange(26, 3, 1, 3).merge();
        RFQForm.getRange(26, 1).setValue(PartDetails[j][1]);
        RFQForm.getRange(26, 3).setValue(PartDetails[j][2]);
        RFQForm.getRange(26, 6).setValue(PartDetails[j][3]);
        RFQForm.getRange(26, 8).setFormula(`=IFERROR(F26*G26,0)`);
      }
    }
  }
  
}; //Close RFQFormDrafter function
