function DraftHTMLEmails(){
  
  var emailTemp = HtmlService.createTemplateFromFile("email"); 
  //var file = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ZxFpnLA_pGbAa81d9eXxkri5nxKqjgy0nVWq64FV9cU/");
  //var ss = file.getSheetByName("TrailSheet");
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  var addCC = "";
  emailTemp.addText = "";
  
  var sheetNames = getSheetNames();
  
  if (sheetNames.includes("Settings")) {
    
    // Get additional settings from sheet
    var settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    var settings = settingSheet.getRange(2, 2, 8).getValues();
    
    var settingObj = {
      typeEvent: settings[0], 
      addText: settings[1], 
      addCCEmails: settings[2],  
      userName: settings[3],
      userTitle: settings[4],
      userEmail: settings[5],
      userCell: settings[6],
      userDesk: settings[7]
    }
    
    var eventType = settingObj['typeEvent'];
    
    emailTemp.addText = settingObj['addText'];
    var addCC = settingObj['addCCEmails'];
    emailTemp.userName = settingObj['userName'];
    emailTemp.userTitle = settingObj['userTitle'];
    emailTemp.userEmail = "Email: " + settingObj['userEmail'];
    emailTemp.userCell = "Cell: "+ settingObj['userCell'];
    emailTemp.userDesk = "Desk: " + settingObj['userDesk'];
  }
  else {
    emailTemp.userName = "";
    emailTemp.userTitle = "";
    emailTemp.userEmail = "";
    emailTemp.userCell = "";
    emailTemp.userDesk = "";
  }
 
  var headerRow = checkTemplate()[0];
  
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
  var NameInd = indices[14]+1; 
  
  // Get the array of values for these indexes
  
  var SCodes = ss.getRange(headerRow+1, SCodeInd, lr, 1).getValues();
  var CompNames = ss.getRange(headerRow+1, CompInd, lr, 1).getValues();
  var SEmail = ss.getRange(headerRow+1, SEmailInd, lr, 1).getValues();
  var AEmail = ss.getRange(headerRow+1, AEmailInd, lr, 1).getValues();
  var PN = ss.getRange(headerRow+1, PNInd, lr, 1).getValues();
  var PDesc = ss.getRange(headerRow+1, PDescInd, lr, 1).getValues();
  var Qty = ss.getRange(headerRow+1, QtyInd, lr, 1).getValues();
  var Ship = ss.getRange(headerRow+1, ShipInd, lr, 1).getDisplayValues();
  var mrd = ss.getRange(headerRow+1, mrdInd, lr, 1).getDisplayValues();
  var Status = ss.getRange(headerRow+1, StatusInd, lr, 1).getValues();
  var Name = ss.getRange(headerRow+1, NameInd, lr, 1).getValues();
  
  //Constant Info
  var Prog = ss.getRange(1, 1).getValue();
  var reqdDate = ss.getRange(3, 1).getDisplayValue();
  
  // Get all the unique supplier codes
  
  var uniqueSupplierCodes = [];
  
  for (var i=0; i < SCodes.length; i++) {
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
  
  var miscText = {initial: "Hello,", 
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
    for (var i=0; i<lr-headerRow; i++) {
      if (SCodes[i][0] === uniqueSupplierCodes[j] && SEmail[i][0] !== "" && Status[i][0] === "" && PN[i][0] !== "") {
        
        var shipAddress = Ship[i][0];
        var toEmail = "";
        var ccEmail = "";
        const regex = /(\w+\.*\w+\.*\w+@\w+\.*\-*\+*\w+\.*\w+)/gm;
        const SEmailstr = SEmail[i][0];
        const AEmailstr = AEmail[i][0];
        let SEmailMatches = SEmailstr.match(regex);
        let AEmailMatches = AEmailstr.match(regex);
        if (SEmailMatches !== null) {
          SEmailMatches.forEach(match => { 
                          toEmail += match + ", ";})
        }
        if (AEmailMatches !== null) {
          AEmailMatches.forEach(match => { 
                          ccEmail += match + ", ";})
        }
        emailTemp.name = Name[i][0].split(" ")[0];
        emailTemp.shipAddress = Ship[i][0];
        emailTemp.prog = Prog;
        sub['supplier'] = CompNames[i][0];
        
        // Add part number, description and qty to PartDetails array
        PartDetails.push([PN[i][0], PDesc[i][0], Qty[i][0], mrd[i][0]]);
        
        // Update Status Column to say 'RFQ SENT'
        ss.getRange(i+headerRow+1, StatusInd).setValue("RFQ SENT").setBackground("Indigo").setFontColor("White");
        Status.splice(i, 1,"RFQ SENT");
      };
    };
    
    // Draft email only if Part Details for a supplier has data
    if (PartDetails.length > 0){
      
      // Print the parts on separate line in the email body
      var PrintPNs = PartDetails.join('\n');
      emailTemp.partDetails = PartDetails;
      
      // Crete Email subject combining all the information
      var fullSubject = `${sub['initial']}: ${sub['program']} - ${sub['supplier']}`;
      
      // Create email body using above information
      var htmlMessage = emailTemp.evaluate().getContent();
      
      var fullBody = `${miscText['initial']}\n${miscText['part1']} ${reqdDate}:\n\n${PrintPNs}\n
      \n${miscText['part3']}\n\n${shipAddress}\n\n${miscText['part4']}`;

      // If there are additional people required in CC.
      if (addCC !== "") {
        var ccAddresses = addCC + ", " + ccEmail;
      } else {
        var ccAddresses = ccEmail;
      }
      
      // Create Email Draft
      GmailApp.createDraft(toEmail, fullSubject, fullBody, {cc: ccAddresses, htmlBody: htmlMessage});
      
      emailCount++;
      
    };
    
    //Rest PNs to blank after the loop to store information for next supplier
    PartDetails = [];
  };
  
  return emailCount;    
  
}; //Close DraftHTMLEmails function
