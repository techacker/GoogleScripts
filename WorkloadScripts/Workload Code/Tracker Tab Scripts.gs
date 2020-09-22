function createNewEventTracker() {
  
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var EventSheet = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Events");
  var lr = EventSheet.getLastRow();
  var lc = EventSheet.getLastColumn();
  var headerRow = getHeaderRow(EventSheet, "Tracker URL");
  var colIndices = getColIndex();
  var urlColInd = colIndices[11];
  var trackerURLArray = EventSheet.getRange(headerRow+1, urlColInd+1, lr-headerRow, 1).getValues();
  
  // Get the tracker URL from Events Tab
  
  for (var i=0; i<trackerURLArray.length; i++) {
    if (trackerURLArray[i][0] !== "") {
      var eventTrackerURL = trackerURLArray[i][0];
      var SummarySheet = SpreadsheetApp.openByUrl(eventTrackerURL).getSheetByName("Summary");
      var lc = SummarySheet.getLastColumn();
      var lr = SummarySheet.getLastRow();
      var SummaryRange = SummarySheet.getRange(1, 1, lr, lc).getValues();
      var headerRow = getHeaderRow(SummarySheet, "Tab");
      var eventTabsData = SummarySheet.getRange(headerRow+1, 1, lr-headerRow, lc).getDisplayValues();
      var infoCol = 3;
      var eventInfoData = SummarySheet.getRange(1, infoCol, headerRow-1, 1).getDisplayValues();
      if (eventInfoData[0][0] !== "<Event Title>" || eventInfoData[0][0] !== "") {
        addEventTitles(eventInfoData, eventTabsData);
        updateEventsData(eventInfoData[0][0], eventTabsData);
      }
    }
  } 
}


function addEventTitles(eventInfoData,eventTabsData) {
  
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var TrackerDataTab = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Tracker Data");
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerRow = getHeaderRow(TrackerDataTab, "Tab");
  var TrackerTabColIndices = getTrackerTabIndices(TrackerDataTab, headerRow);
  var TrackerDataTabRange = TrackerDataTab.getRange(headerRow+1, 1, lr, lc).getDisplayValues();
  
  // Logger.log(TrackerTabColIndices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0]
  // [Event Title, Tab, PPPM Engineer, MRD, Total No. of Parts, % Reqs, % PO, % Rec'd, % RFQ Pending, 
  // # RFQ Sent, # REQ Submitted, # PO Issued, # Parts Received, Cost]
  
  var EventTileCol = TrackerTabColIndices[0];
  var TabCol = TrackerTabColIndices[1];
  var PPPMEngCol = TrackerTabColIndices[2];
  var MRDInd = TrackerTabColIndices[3];
  var TotalPartsCol = TrackerTabColIndices[4];
  var PercReqCol = TrackerTabColIndices[5];
  var PercPOCol = TrackerTabColIndices[6];
  var PercRecdCol = TrackerTabColIndices[7];
  var PercRFQCol = TrackerTabColIndices[8];
  var RFQSentCol = TrackerTabColIndices[9];
  var REQCol = TrackerTabColIndices[10];
  var POCol = TrackerTabColIndices[11];
  var PartsRecdCol = TrackerTabColIndices[12];
  var CostCol = TrackerTabColIndices[13];
  
  // Get Event Titles from Workload file
  var EventTitles = [];
  var TabNames = [];
  var EventandTabs = [];
  for (var i=0; i<TrackerDataTabRange.length; i++) {
    EventTitles.push(TrackerDataTabRange[i][EventTileCol]);
    TabNames.push(TrackerDataTabRange[i][TabCol]);
    EventandTabs.push([TrackerDataTabRange[i][EventTileCol], TrackerDataTabRange[i][TabCol]]);
  }
  
  // Logger.log(eventInfoData);
  // [[2021 WS DV], [Anurag Bansal], [VD00240], [CPG], [1256], [1256 Chrylser Dr, Detroit, MI], [Mike Kaslly]]
  
  var eventInfo = [];
  for (var i=0; i<eventInfoData.length; i++) {
    eventInfo.push(eventInfoData[i][0]);
  }
  
  // Logger.log(eventInfo);
  // [2021 WS DV, Anurag Bansal, VD00240, CPG, 1256, 1256 Chrylser Dr, Detroit, MI, Mike Kaslly]

  var updatedEventData = [];
  
  // Logger.log(eventTabData);
  // [[Chassis, Anurag Bansal, 11/11/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6], 
  // [Interior, Anurag Bansal, 11/12/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6], 
  // [MASTER, Name (dropdown), 09/15/2020, 0, 0.0%, 0.0%, 0.0%, 0.0%, $0.00, 0, 0, 0, 0]]
  
  var tabCount = 0;
  // Get number of tabs in Event Tracker Sheet
  for (var i=0; i<eventTabsData.length; i++) {
    if (eventTabsData[i][0].toUpperCase() !== "MASTER") {
      // Include Event Title in the Event Data
      updatedEventData.unshift([eventInfo[0], eventTabsData[i]]);
      tabCount += 1;
    }
  }
  
  // Logger.log(updatedEventData);
  // [2021 WS DV, [Interior,Anurag Bansal,11/12/2020,25,60.0%,40.0%,8.0%,24.0%,$0.00,15,10,2,6], 
  // 2021 WS DV, [Chassis,Anurag Bansal,11/11/2020,40,62.5%,62.5%,25.0%,37.5%,$0.00,25,25,10,15]]
  
  // Count how many times the event title appears in Workload Tracker Data Tab 
  var eventRowCount = 0;
  for (var i=0; i<EventTitles.length; i++) {
    if (EventTitles[i] === eventInfo[0]) {
      eventRowCount += 1;
    }
  }
  
  // Find Event Row
  for (var i=0; i<EventTitles.length; i++) {
    if (EventTitles[i] === eventInfo[0]) {
      var EventRow = i + headerRow + 1;
      break;
    }
  }
  
  // Add rows if the count of sheets in Tracker file and workload file doesn't match
  if (tabCount === 0 && eventRowCount === 0) {
    var addRow = TrackerDataTab.getRange(headerRow+1, 1, 1,lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(headerRow+1, EventTileCol+1, 1,1).setValue(eventInfo[0]);
  }
  else if (eventRowCount === 0) {
    var addRow = TrackerDataTab.getRange(headerRow+1, 1, tabCount,lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(headerRow+1, EventTileCol+1, tabCount,1).setValue(eventInfo[0]);
  }
  else if (eventRowCount < tabCount) {
    var addRow = TrackerDataTab.getRange(EventRow, EventTileCol+1, tabCount - eventRowCount,lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(EventRow, EventTileCol+1, tabCount - eventRowCount,1).setValue(eventInfo[0]);
  }
  
}

function updateEventsData(eventInfoData, eventTabsData) {
  
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var TrackerDataTab = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Tracker Data");
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerRow = getHeaderRow(TrackerDataTab, "Tab");
  var TrackerTabColIndices = getTrackerTabIndices(TrackerDataTab, headerRow);
  var TrackerDataTabRange = TrackerDataTab.getRange(headerRow+1, 1, lr, lc).getDisplayValues();
  
  // Logger.log(TrackerTabColIndices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0]
  // [Event Title, Tab, PPPM Engineer, MRD, Total No. of Parts, % Reqs, % PO, % Rec'd, % RFQ Pending, 
  // # RFQ Sent, # REQ Submitted, # PO Issued, # Parts Received, Cost]
  
  var EventTileCol = TrackerTabColIndices[0];
  var TabCol = TrackerTabColIndices[1];
  var PPPMEngCol = TrackerTabColIndices[2];
  var MRDInd = TrackerTabColIndices[3];
  var TotalPartsCol = TrackerTabColIndices[4];
  var PercReqCol = TrackerTabColIndices[5];
  var PercPOCol = TrackerTabColIndices[6];
  var PercRecdCol = TrackerTabColIndices[7];
  var PercRFQCol = TrackerTabColIndices[8];
  var RFQSentCol = TrackerTabColIndices[9];
  var REQCol = TrackerTabColIndices[10];
  var POCol = TrackerTabColIndices[11];
  var PartsRecdCol = TrackerTabColIndices[12];
  var CostCol = TrackerTabColIndices[13];
  
  var updatedEventData = [];
  var tabCount = 0;
  
  // Get number of tabs in Event Tracker Sheet
  for (var i=0; i<eventTabsData.length; i++) {
    if (eventTabsData[i][0].toUpperCase() !== "MASTER") {
      // Include Event Title in the Event Data
      updatedEventData.unshift([eventInfoData, eventTabsData[i]]);
      tabCount += 1;
    }
  }
    
  // [[[2021 WS DV], [Interior, Anurag Bansal, 11/12/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6]], 
  // [[2021 WS DV], [Chassis, Anurag Bansal, 11/11/2020, 40, 62.5%, 62.5%, 25.0%, 37.5%, $0.00, 25, 25, 10, 15]]]
  
  for (var j=0; j<updatedEventData.length; j++) {
    if (updatedEventData.length !== 0) {
      for (var i=0; i<TrackerDataTabRange.length; i++) {
        if (updatedEventData[j][0] === TrackerDataTabRange[i][0]) {
          var Row = i+j+headerRow+1; 
          TrackerDataTab.getRange(Row, TabCol+1, 1,1).setValue(updatedEventData[j][1][0]);
          TrackerDataTab.getRange(Row, PPPMEngCol+1, 1,1).setValue(updatedEventData[j][1][1]);
          TrackerDataTab.getRange(Row, MRDInd+1, 1,1).setValue(updatedEventData[j][1][2]);
          TrackerDataTab.getRange(Row, TotalPartsCol+1, 1,1).setValue(updatedEventData[j][1][3]);
          TrackerDataTab.getRange(Row, PercReqCol+1, 1,1).setValue(updatedEventData[j][1][4]);
          TrackerDataTab.getRange(Row, PercPOCol+1, 1,1).setValue(updatedEventData[j][1][5]);
          TrackerDataTab.getRange(Row, PercRecdCol+1, 1,1).setValue(updatedEventData[j][1][6]);
          TrackerDataTab.getRange(Row, PercRFQCol+1, 1,1).setValue(updatedEventData[j][1][7]);
          TrackerDataTab.getRange(Row, CostCol+1, 1,1).setValue(updatedEventData[j][1][8]);
          TrackerDataTab.getRange(Row, RFQSentCol+1, 1,1).setValue(updatedEventData[j][1][9]);
          TrackerDataTab.getRange(Row, REQCol+1, 1,1).setValue(updatedEventData[j][1][10]);
          TrackerDataTab.getRange(Row, POCol+1, 1,1).setValue(updatedEventData[j][1][11]);
          TrackerDataTab.getRange(Row, PartsRecdCol+1, 1,1).setValue(updatedEventData[j][1][12]);
          break;
        }
      }
    }
  }
  
}
