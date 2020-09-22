// ********* Start: Update Status Column Function

function updateStatusColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkTemplate() function to get the required headerRow and column indexes.
  var headerRow = checkTemplate()[0];  
  var indices = checkTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];  
  
  var PNInd = indices[4]+1;
  var StatusInd = indices[9]+1;
  var REQNumInd = indices[10]+1;
  var PONumInd = indices[11]+1;
  var PartRecdInd = indices[15]+1;
  
  var PartNumbers = ss.getRange(headerRow + 1, PNInd, lr-headerRow, 1).getDisplayValues();
  var StatusRange = ss.getRange(headerRow + 1, StatusInd, lr-headerRow, 1).getValues();
  var REQNumbers = ss.getRange(headerRow + 1, REQNumInd, lr-headerRow, 1).getDisplayValues();
  var PONumbers = ss.getRange(headerRow + 1, PONumInd, lr-headerRow, 1).getDisplayValues();
  var PartRecd = ss.getRange(headerRow + 1, PartRecdInd, lr-headerRow, 1).getDisplayValues();
  
  // Based on values in Parts Received, PO Number and REQ no. column, update Status column
  
  for (var i=0; i<PartNumbers.length; i++) {
    // Only if there is a part number in Part Number col
    if (PartNumbers[i][0] !== "") {
      // If there is data in Parts Received, PO Number and REQ No cols (all three)
      if (PartRecd[i][0] !== "" && PONumbers[i][0] !== "" && REQNumbers[i][0] !== "") {
        ss.getRange(i + headerRow + 1, StatusInd, 1, 1).setValue("PARTS RECEIVED").setBackground("GREEN").setFontColor("White");
      }
      // If there is data in both PO Number and REQ No cols 
      else if (PONumbers[i][0] !== "" && REQNumbers[i][0] !== "") {
        ss.getRange(i + headerRow + 1, StatusInd, 1, 1).setValue("PO ISSUED").setBackground("YELLOW").setFontColor("Black");
      }
      // If there is data in REQ No cols 
      else if (REQNumbers[i][0] !== "") {
        ss.getRange(i + headerRow + 1, StatusInd, 1, 1).setValue("REQ SUBMITTED").setBackground("CYAN").setFontColor("Black");
      }
    }
  }
  
  
  //updateProjectStatus();
}

// ********* End: Update Status Column Function

/*
// ********* Start: Update Project Status Function

function updateProjectStatus() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkTemplate() function to get the required headerRow and column indexes.
  var headerRow = checkTemplate()[0];  
  var indices = checkTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];  
  var StatusInd = indices[9]+1;
  var REQNumInd = indices[10]+1;
  
  var StatusRange = ss.getRange(headerRow + 1, StatusInd, lr-headerRow, 1).getValues();
  var REQRange = ss.getRange(headerRow + 1, REQNumInd, lr-headerRow, 1).getValues();
  
  // Define a class for using the status column entries
  
  class Status {
    constructor(name, times) {
      this.name = name;
      this.times = times;
    }
  }
  
  // Initialize the status entries using class
  
  const PreQuote = new Status("PRE QUOTE", 0);
  const RFQSent = new Status("RFQ SENT", 0);
  const REQSubmitted = new Status("REQ SUBMITTED", 0);
  const POIssued = new Status("PO ISSUED", 0);
  const PartsReceived = new Status("PARTS RECEIVED", 0);
  
  // Loop through the Status column to count each entry
  
  for (var i=0; i<REQRange.length; i++) {
    if (REQRange[i][0] !== "") {
      REQSubmitted.times += 1;
    }
  }
  
  for (var i=0; i<StatusRange.length; i++) {
    if (StatusRange[i][0]) {
      switch (StatusRange[i][0].toUpperCase()){
        case RFQSent.name:
          RFQSent.times += 1;
          break;
        case POIssued.name:
          POIssued.times += 1;
          break;
        case PartsReceived.name:
          PartsReceived.times += 1;
          break;
        default:
          PreQuote.times += 1;
      }
    } 
  }
  
  ss.getRange(1, 16).setValue(PreQuote.times + RFQSent.times + REQSubmitted.times  + POIssued.times + PartsReceived.times);
  ss.getRange(1, 17).setValue(REQSubmitted.times);
  ss.getRange(1, 18).setValue(POIssued.times);
  ss.getRange(1, 19).setValue(PartsReceived.times);
  //ss.getRange(1, 21).setValue(PreQuote.times);
  ss.getRange(1, 22).setValue(RFQSent.times);
  
  
}

// ********* End: Update Project Status Function
*/
