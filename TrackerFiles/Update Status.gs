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
}


function pushEventUpdates() {
  
  // Summary Sheet
  
  var SummarySheetURL = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var SummarySheetLink = PPPMWorkloadGoogleScript.parseURL(SummarySheetURL);
  
  var SummarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  var sslr = SummarySheet.getLastRow();
  var sslc = SummarySheet.getLastColumn();
  
  
  // Workload Events
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/"; // Test Workload File
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var EventSheet = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Events");
  var eslr = EventSheet.getLastRow();
  var eslc = EventSheet.getLastColumn(); 
  var headerRow = PPPMWorkloadGoogleScript.getHeaderRow(EventSheet, "Tracker URL");
  
  // Get col indexes from Events Sheet
  var colIndices = PPPMWorkloadGoogleScript.getColIndex(EventSheet, headerRow);
  
  var VFInd = colIndices[0];
  var EventTitleInd = colIndices[1];
  var EventStatusInd = colIndices[3];
  var ProgMgrInd = colIndices[4];
  var urlColInd = colIndices[11];
  
  var trackerURLArray = EventSheet.getRange(headerRow+1, urlColInd+1, eslr-headerRow, 1).getValues(); 
  
  var EventTitle, ProgMgr, EventStatus, VF, DaystoMRD;
  var SummarySheet, sslc, sslr, SummaryRange, ssheaderRow, eventTabsData;
  
  // Update Master Workload File with information
  
  for (var i=0; i<trackerURLArray.length; i++) {
    if (trackerURLArray[i][0] === SummarySheetLink) {
      
      EventTitle = EventSheet.getRange(i + headerRow+1, EventTitleInd+1, 1, 1).getValue();
      ProgMgr = EventSheet.getRange(i + headerRow+1, ProgMgrInd+1, 1, 1).getValue();
      EventStatus = EventSheet.getRange(i+ headerRow+1, EventStatusInd+1, 1, 1).getValue();
      VF = EventSheet.getRange(i+ headerRow+1, VFInd+1, 1, 1).getValue();
      
      // Get Information from Event Tracker's Summary Tab
      
      SummaryRange = SummarySheet.getRange(1, 1, sslr, sslc).getValues();
      ssheaderRow = PPPMWorkloadGoogleScript.getHeaderRow(SummarySheet, "Tab");
      
      // Event Data from individual trackers file
      eventTabsData = SummarySheet.getRange(ssheaderRow+1, 1, sslr-ssheaderRow, sslc).getDisplayValues();
      
      var infoCol = 3;
      var eventInfoData = SummarySheet.getRange(1, infoCol, ssheaderRow-1, 1).getDisplayValues();
      
      // Remove tabs that are either "MASTER" or "OVERALL EVENT STATUS"
      let usefulTabs = eventTabsData.filter(tabName => tabName[0].toUpperCase() !== "MASTER" 
      && tabName[0].toUpperCase() !== "OVERALL EVENT STATUS");
      
      PPPMWorkloadGoogleScript.addEventTitles(EventTitle, VF, ProgMgr, eventInfoData, eventTabsData);
      PPPMWorkloadGoogleScript.updateEventsData(EventTitle, eventTabsData);
    }
  }   
}

// ********* End: Update Status Column Function
