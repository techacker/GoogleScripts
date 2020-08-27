// ------------------------------------------------------------------------- RFQ Draft Email Creations Program ------------------------------------------------------------------
// -------------------------------------------------------------------------      Author: Anurag Bansal        ------------------------------------------------------------------
// -------------------------------------------------------------------------          Version: 1.4.0           ------------------------------------------------------------------
// -------------------------------------------------------------------------      Only for PPPM Programs       ------------------------------------------------------------------
// -----------Further Development Ideas & Change Log:
// -----------1. Validate the Tracker by finding the header            ----- Completed 8/21/2020
// -----------2. Update Status on Summary Sheet                        ----- Completed 8/25/2020
// -----------3. Update Supplier Info from a Master Sheet              ----- Completed 8/26/2020
//------------4. Update PO information from a Master Sheet             ----- Completed 8/25/2020
//------------5. Auto Create RFQ Forms
//------------6. Save RFQ forms in a Drive Folder with Program Name
//------------7. Update GMail draft to include developed RFQ Forms
//------------8. Send REQ Emails to suppliers
//------------9. Send PO Emails to suppliers
//-----------10. Send Reminder emails to suppliers
//-----------11. Create New Sheets based on Master Template            ----- Completed 8/26/2020
//-----------12. Add new row with New Sheet Name in Summary Sheet      ----- Completed 8/27/2020
//-----------13. Update Master Supplier Info from Trackers


// ********* Start: Main Function

function createRFQEmails() {
  var emailCount = DraftEmails();
  
  // Show a browser confirmation
  if (emailCount > 0){
    Browser.msgBox("Success!!!", emailCount + ' RFQ emails were saved in Gmail drafts. Please review before sending emails.', Browser.Buttons.OK);
  } else {
    Browser.msgBox("Oops!!!", 'It seems RFQ Emails were previously sent. Please clear Status column before trying again.', Browser.Buttons.OK);
  };
};

// ********* End: Main Function

// ********* Start: Find Header Function

function findHeader() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = ss.getDataRange().getValues();
  
  //Logger.log(data.indexOf("Status"));
  
  var headerRow = 0;
  
  // Find the header row that contains 'Status'
  for (var i=0; i<data.length; i++) {
    if (data[i][1] === "Supplier Code"){
      headerRow = i+1;
    };
  };
  
  // Check if header contain "Supplier Code"
  if (headerRow === 0) {
    Browser.msgBox("Supplier Code column is not found. Make sure it is called 'Supplier Code', then try again.", Browser.Buttons.OK);
    return;
  }
  else {
    return headerRow;
  }
  
};

// ********* End: Find Header Function


// ********* Start: Verify Template Function

function checkTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  var headerRow = findHeader();
  
  // Header Columns in Template:
  // ['Status',	'Supplier Code', 'Company Name', 'Supplier Contact', 'Sales Manager Email', 'Engg Manager Email', 'Phone  - Sales Manager', 
  // 'PR/PC/RE', 'VSC:', 'Part Number', 'C/L', 'Part Description', '"Line Up Code"', 'Quantity Ordered', 'Ship To Location', 
  // 'Purchase Requisition #', 'PO # ', 'Part MRD', 'Supplier Promise Date', 'PO Issue Date', '"PARTS RECEIVED Y/ Blank"', '"PSAP GR Y/ Blank"', 'Special Means', '"Quoted Piece Cost"',
  // 'Misc. Costs (ie, Shipping, Set-up)', 'Tooling Cost', '"Production Piece"', 'Total REQ Cost', 'COMMENTS']
  
  var lookupValues = ["Supplier Code",
                      "Company Name",
                      "Sales Manager Email",
                      "Engg Manager Email",
                      "Part Number",
                      "Part Description",
                      "Quantity Ordered",
                      "Ship To Location",
                      "Part MRD", 
                      "Status", 
                      "Purchase Requisition #", 
                      "PO # ", 
                      "PO Issue Date", 
                      "Phone  - Sales Manager", 
                      "Supplier Contact"]; 
  
  var lookupInd = [];
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];   
  
  // Get indexes of all required columns
  if (header.length >= 0) {
    for (var i = 0; i<lookupValues.length; i++) {
      lookupInd.push(header.indexOf(lookupValues[i]));
    };
  };
  
  // If something is missing, then let the user know what is missing
  for (var i = 0; i<lookupInd.length; i++) {
    if (lookupInd[i] === -1.0) {
      Browser.msgBox(lookupValues[i] + " is missing.", Browser.Buttons.OK);
    };
  };
  
  // lookupInd = [1.0, 2.0, 4.0, 5.0, 9.0, 11.0, 13.0, 14.0, 17.0, 0.0, 15.0, 16.0, 19.0] 
  // Sample that specifies where the respective columns are.

  return [headerRow, lookupInd];
};

// ********* End: Check Template Function



// ********* Start: Draft Email Function

function DraftEmails(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  var headerRow = checkTemplate()[0];
  
  var lookupValues = ["Supplier Code",
                      "Company Name",
                      "Sales Manager Email",
                      "Engg Manager Email",
                      "Part Number",
                      "Part Description",
                      "Quantity Ordered",
                      "Ship To Location",
                      "Part MRD", 
                      "Status"]; 
  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
  
  var indices = checkTemplate()[1];  
  
  // Get the index values of all necessary columns for Email.
 
  var SCodeInd = indices[0]+1;
  var CompInd = indices[1]+1;
  var SEmailInd = indices[2]+1;
  var AEmailInd = indices[3]+1;
  var PNInd = indices[4]+1;
  var PDescInd = indices[5]+1;
  var QtyInd = indices[6]+1;
  var ShipInd = indices[7]+1;
  var mrdInd = indices[8]+1;
  var StatusInd = indices[9]+1;  
  
  // Get the array of values for these indexes
  
  var SCodes = ss.getRange(headerRow+1, SCodeInd, lr, 1).getValues();
  var CompNames = ss.getRange(headerRow+1, CompInd, lr, 1).getValues();
  var SEmail = ss.getRange(headerRow+1, SEmailInd, lr, 1).getValues();
  var AEmail = ss.getRange(headerRow+1, AEmailInd, lr, 1).getValues();
  var PN = ss.getRange(headerRow+1, PNInd, lr, 1).getValues();
  var PDesc = ss.getRange(headerRow+1, PDescInd, lr, 1).getValues();
  var Qty = ss.getRange(headerRow+1, QtyInd, lr, 1).getValues();
  var Ship = ss.getRange(headerRow+1, ShipInd, lr, 1).getValues();
  var mrd = ss.getRange(headerRow+1, mrdInd, lr, 1).getDisplayValues();
  var Status = ss.getRange(headerRow+1, StatusInd, lr, 1).getValues();
  
  //Constant Info
  
  var Prog = ss.getRange(1, 12).getValue();
  
  // Get all the unique supplier codes
  
  var uniqueSupplierCodes = [];
  
  for (i=0; i < SCodes.length; i++) {
    if (uniqueSupplierCodes.indexOf(SCodes[i][0]) === -1.0 && SCodes[i][0] !== "") {
      uniqueSupplierCodes.push(SCodes[i][0]);
    };
  };
  
  // Logger.log(uniqueSupplierCodes) returns
  // [64964 H, 16180 A, 90550 K, 70280 D, 23079.0, 12279.0, 61177 M, 51649 O, 64642 E, 40732AA, 40732 R, 38201 B, 21193AJ, 62010.0, 62283.0, 23758 O, 38303 N, 26285 L]  
  
  //Email Draft standard text
  
  var sub = {initial: "FCA : RFQ ", 
                 program: Prog, 
                 supplier: ""
                };
  
  var miscText = {initial: "Hello from FCA team,", 
                  part1: "Can you please provide a quote for the following components with an MRD of ", 
                  part2: "",
                  part3: "The parts will need to be shipped to the following address:", 
                  part4: "Please let me know if there are any questions or concerns.\n\nWe expect the quotes back within 3 to 5 business days."
                 };
  
  // Store all the Part information (Part No., Description and Qty)
  var PartDetails = [];
  
  var emailCount = 0;  // To count number of emails created.
  
  // Loop through the parts list and match with unique supplier codes to draft their emails
  
  for (var j=0; j<uniqueSupplierCodes.length; j++) {
    for (var i=0; i<lr-6; i++) {
      if (SCodes[i][0] === uniqueSupplierCodes[j] && SEmail[i][0] !== "" && Status[i][0] === "") {
        var reqdDate = mrd[i][0];
        var shipAddress = Ship[i][0];
        var toEmail = SEmail[i][0];
        var ccEmail = AEmail[i][0];
        sub['supplier'] = CompNames[i][0];
        
        // Add part number, description and qty to PartDetails array
        PartDetails.push(PN[i][0] + " - " + PDesc[i][0] + " - Qty: " + Qty[i][0]);
        
        // Update Status Column to say 'RFQ SENT'
        ss.getRange(i+headerRow+1, StatusInd).setValue("RFQ SENT");
        Status.splice(i, 1,"RFQ SENT");
      };
    };
    
    // Draft email only if Part Details for a supplier has data
    
    if (PartDetails.length > 0){
      
      // Print the parts on separate line in the email body
      var PrintPNs = PartDetails.join('\n');
      
      // Crete Email subject combining all the information
      var fullSubject = `${sub['initial']}: ${sub['program']} - ${sub['supplier']}`;
      
      // Create email body using above information
      var fullBody = `${miscText['initial']}\n${miscText['part1']} ${reqdDate}:\n\n${PrintPNs}\n\n${miscText['part3']}\n\n${shipAddress}\n\n${miscText['part4']}`;
      
      // Create Email Draft
      GmailApp.createDraft(toEmail, fullSubject, fullBody,{cc: ccEmail});
      
      emailCount++;
      
    };
    
    //Rest PNs to blank after the loop to store information for next supplier
    PartDetails = [];
  };
  
  return emailCount;    
  
}; //Close DraftEmails function


// ********* End: Draft Email Function


function getSheetNames() {
  var App = SpreadsheetApp;
  var Sheet = App.getActiveSpreadsheet(); 
  var sheetsArray = Sheet.getSheets();
  var sheetNames = [];
  
  for (var i=0; i<sheetsArray.length; i++) {
    sheetNames.push(sheetsArray[i].getSheetName());
  }
  
  return sheetNames;
}

