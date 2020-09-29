function updateTrackerTab() {
  
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/"; // Test Workload File
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var EventSheet = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Events");
  var lr = EventSheet.getLastRow();
  var lc = EventSheet.getLastColumn();
  
  var headerRow = getHeaderRow(EventSheet, "Tracker URL");
  
  // Get col indexes from Events Sheet
  var colIndices = getColIndex();
  var VFInd = colIndices[0];
  var EventTitleInd = colIndices[1];
  var EventStatusInd = colIndices[3];
  var ProgMgrInd = colIndices[4];
  var urlColInd = colIndices[11];
  
  
  var trackerURLArray = EventSheet.getRange(headerRow+1, urlColInd+1, lr-headerRow, 1).getValues();  
  var BasicEventInfo = [];
  var EventTitle, ProgMgr, EventStatus, VF, DaystoMRD;
  var SummarySheet, sslc, sslr, SummaryRange, ssheaderRow, eventTabsData;
  
  // Get the tracker URL from Events Tab
  
  for (var i=0; i<trackerURLArray.length; i++) {
    if (trackerURLArray[i][0] !== "") {
      EventTitle = EventSheet.getRange(i + headerRow+1, EventTitleInd+1, 1, 1).getValue();
      ProgMgr = EventSheet.getRange(i + headerRow+1, ProgMgrInd+1, 1, 1).getValue();
      EventStatus = EventSheet.getRange(i+ headerRow+1, EventStatusInd+1, 1, 1).getValue();
      VF = EventSheet.getRange(i+ headerRow+1, VFInd+1, 1, 1).getValue();
      
      // Get Information from Event Tracker's Summary Tab
      
      SummarySheet = SpreadsheetApp.openByUrl(trackerURLArray[i][0]).getSheetByName("Summary");
      sslc = SummarySheet.getLastColumn();
      sslr = SummarySheet.getLastRow();
      SummaryRange = SummarySheet.getRange(1, 1, sslr, sslc).getValues();
      ssheaderRow = getHeaderRow(SummarySheet, "Tab");
      
      // Events Tab Data from trackers file
      eventTabsData = SummarySheet.getRange(ssheaderRow+1, 1, sslr-ssheaderRow, sslc).getDisplayValues();
      var infoCol = 3;
      var eventInfoData = SummarySheet.getRange(1, infoCol, ssheaderRow-1, 1).getDisplayValues();
      
      
      //addEventTitles(EventTitle, ProgMgr, EventStatus, eventInfoData, eventTabsData);
      
      if (eventInfoData[0][0] !== "<Event Title>" || eventInfoData[0][0] !== "" || eventInfoData[0][1] !== "" || eventInfoData[0][2] !== "" 
          || eventInfoData[0][3] !== "" || eventInfoData[0][4] !== "" || eventInfoData[0][5] !== "" || eventInfoData[0][6] !== "") {
        addEventTitles(EventTitle, VF, ProgMgr, EventStatus, eventInfoData, eventTabsData);
        updateEventsData(eventInfoData[0][0], eventTabsData);
      }
    }
  } 
  
}


function addEventTitles(EventTitle, VehFam, ProgManager, EventStatus, eventInfoData, eventTabsData) {
  
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/"; // Test Workload File
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var TrackerDataTab = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Tracker Data");
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerRow = getHeaderRow(TrackerDataTab, "Tab");
  var TrackerTabColIndices = getTrackerTabIndices(TrackerDataTab, headerRow);
  var TrackerDataTabRange = TrackerDataTab.getRange(headerRow+1, 1, lr, lc).getDisplayValues();
  
  // Logger.log(TrackerTabColIndices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 
  // 5.0, 6.0, 7.0, 8.0, 
  // 9.0, 10.0, 11.0, 12.0, 
  // 13.0, 14.0, 15.0, 16.0, 
  // 17.0, 18.0, 19.0, 20.0, 21.0]
  // [Event Title, Program Manager,	Event Status, Tab, PPPM Engineer, MRD, Total # of Parts, 
  // % REQ, % PO, % Received, Cost, % Cancelled, 
  // # REQ Submitted, # PO Issued, # Parts Received, # Cancelled, 
  // # On Time, # Exception, # Late, # Not Defined, 
  // % On Time, % Exception, % Late, % Not Defined]
  
  var EventTileCol = TrackerTabColIndices[0];
  var PrgMgrCol = TrackerTabColIndices[1];
  var EventStatusCol = TrackerTabColIndices[2];
  var TabCol = TrackerTabColIndices[3];
  var VFCol = TrackerTabColIndices[28];
  var DaystoMRDCol = TrackerTabColIndices[29];
  
  // Get Event Titles from Workload file
  var TrackerTabEventTitles = [];
  var TrackerTabSheetNames = [];
  //var EventandTabs = [];
  for (var i=0; i<TrackerDataTabRange.length; i++) {
    TrackerTabEventTitles.push(TrackerDataTabRange[i][EventTileCol]);
    TrackerTabSheetNames.push(TrackerDataTabRange[i][TabCol]);
    //EventandTabs.push([TrackerDataTabRange[i][EventTileCol], TrackerDataTabRange[i][TabCol]]);
  }
  
  // Logger.log(eventInfoData);
  // [[2022 WL74 PHEV Pre-VP QAF], [John Mark Tarbunas, Nick Antovski, Chantal Saade], [VD00194], [Romeo Technologies], [17986], [101 McLean Romeo, MI 48065], [Scott Hunter]]
  // [[2022 WL74 PHEV VPA QAF], [John Mark Tarbunas, Nick Antovski, Chantal Saade], [VD00194], [FCA Detroit Assy Complex MACK], [], [11570 E Warren Ave, Detroit, MI 48214], [John Mark Tarbunas]]
  // [[2022 WS Rollover Spare Parts], [Chris Sleiman], [VD00228_VIVP_SIMP], [DOW], [1256], [12501 Chrysler Dr, Detroit, MI], [Chris Sleiman]]
  
  // Get Event Info data from the Event Tracker Summary Sheet
  
  var eventInfo = [];
  for (var i=0; i<eventInfoData.length; i++) {
    eventInfo.push(eventInfoData[i][0]);
  }
  
  // Logger.log(eventInfo);
  // [2022 WS Rollover Spare Parts, Chris Sleiman, VD00228_VIVP_SIMP, DOW, 1256, 12501 Chrysler Dr, Detroit, MI, Chris Sleiman]

  var updatedEventData = [];
  
  // [[Chassis, Anurag Bansal, 11/11/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6], 
  // [Interior, Anurag Bansal, 11/12/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6], 
  // [MASTER, Name (dropdown), 09/15/2020, 0, 0.0%, 0.0%, 0.0%, 0.0%, $0.00, 0, 0, 0, 0]]
  
  var tabCount = 0;
  // Get number of tabs in Event Tracker Sheet
  for (var i=0; i<eventTabsData.length; i++) {
    if (eventTabsData[i][0].toUpperCase() !== "MASTER") {
      // Include Event Title in the Event Data
      updatedEventData.unshift([eventInfoData[i][0], eventTabsData[i]]);
      tabCount += 1;
    }
  }
  
  // Logger.log(updatedEventData);
  // [[2022 WS Rollover Spare Parts, [CHASSIS, ANURAG BANSAL, 09/28/2020, 43, 44.2%, 37.2%, 7.0%, 0.0%, $38,151.26, 19, 16, 3, 0]]]
  
  // Count how many times the event title appears in Workload Tracker Data Tab 
  var eventRowCount = 0;
  for (var i=0; i<TrackerTabEventTitles.length; i++) {
    if (TrackerTabEventTitles[i] === eventInfo[0]) {
      eventRowCount += 1;
    }
  }
  
  // Find Event Row
  for (var i=0; i<TrackerTabEventTitles.length; i++) {
    if (TrackerTabEventTitles[i] === eventInfo[0]) {
      var EventRow = i + headerRow + 1;
      break;
    }
  }
  
  // Logger.log(eventRowCount, EventRow);
  
  // Add rows if the count of sheets in Tracker file and workload file doesn't match
  if (tabCount === 0 && eventRowCount === 0) {
    //Logger.log("Inside if", headerRow+1, EventTitle, ProgManager, EventStatus);
    var addRow = TrackerDataTab.getRange(headerRow+1, 1, 1,lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(headerRow+1, EventTileCol+1, 1,1).setValue(EventTitle);
    TrackerDataTab.getRange(headerRow+1, VFCol+1, 1,1).setValue(VehFam);
    TrackerDataTab.getRange(headerRow+1, PrgMgrCol+1, 1,1).setValue(ProgManager);
    TrackerDataTab.getRange(headerRow+1, EventStatusCol+1, 1,1).setValue(EventStatus);
  }
  else if (eventRowCount === 0) {
    //Logger.log("Inside 1st else if", tabCount, EventTitle, ProgManager, EventStatus);
    var addRow = TrackerDataTab.getRange(headerRow+1, 1, tabCount,lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(headerRow+1, EventTileCol+1, tabCount,1).setValue(EventTitle);
    TrackerDataTab.getRange(headerRow+1, VFCol+1, tabCount,1).setValue(VehFam);
    TrackerDataTab.getRange(headerRow+1, PrgMgrCol+1, tabCount,1).setValue(ProgManager);
    TrackerDataTab.getRange(headerRow+1, EventStatusCol+1, tabCount,1).setValue(EventStatus);
  }
  else if (eventRowCount < tabCount) {
    //Logger.log("Inside 2nd Else If", EventTileCol+1, tabCount - eventRowCount, EventTitle, ProgManager, EventStatus);
    var addRow = TrackerDataTab.getRange(EventRow, EventTileCol+1, tabCount - eventRowCount,lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(EventRow, EventTileCol+1, tabCount - eventRowCount,1).setValue(EventTitle);
    TrackerDataTab.getRange(EventRow, VFCol+1, tabCount - eventRowCount,1).setValue(VehFam);
    TrackerDataTab.getRange(EventRow, PrgMgrCol+1, tabCount - eventRowCount,1).setValue(ProgManager);
    TrackerDataTab.getRange(EventRow, EventStatusCol+1, tabCount - eventRowCount,1).setValue(EventStatus);
  }  
  
}

function updateEventsData(eventInfoData, eventTabsData) {
  
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/"; // Test Workload File
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var TrackerDataTab = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Tracker Data");
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerRow = getHeaderRow(TrackerDataTab, "Tab");
  var TrackerTabColIndices = getTrackerTabIndices(TrackerDataTab, headerRow);
  var TrackerDataTabRange = TrackerDataTab.getRange(headerRow+1, 1, lr, lc).getDisplayValues();
  
  // Logger.log(TrackerTabColIndices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 
  // 5.0, 6.0, 7.0, 8.0, 
  // 9.0, 10.0, 11.0, 12.0, 
  // 13.0, 14.0, 15.0, 16.0, 
  // 17.0, 18.0, 19.0, 20.0, 21.0]
  // [Event Title, Program Manager,	Event Status, Tab, PPPM Engineer, MRD, Total # of Parts, 
  // % REQ, % PO, % Received, Cost, % Cancelled, 
  // # REQ Submitted, # PO Issued, # Parts Received, # Cancelled, 
  // # On Time, # Exception, # Late, # Not Defined, 
  // % On Time, % Exception, % Late, % Not Defined]
  
  var EventTileCol = TrackerTabColIndices[0];
  var PrgMgrCol = TrackerTabColIndices[1];
  var EventStatusCol = TrackerTabColIndices[2];
  var TabCol = TrackerTabColIndices[3];
  var PPPMEngCol = TrackerTabColIndices[4];
  var MRDInd = TrackerTabColIndices[5];
  var TotalPartsCol = TrackerTabColIndices[6];
  
  var PercReqCol = TrackerTabColIndices[7];
  var PercPOCol = TrackerTabColIndices[8];
  var PercRecdCol = TrackerTabColIndices[9];
  var CostCol = TrackerTabColIndices[10];
  var PercCancelledCol = TrackerTabColIndices[11];
  
  var REQCol = TrackerTabColIndices[12];
  var POCol = TrackerTabColIndices[13];
  var PartsRecdCol = TrackerTabColIndices[14];
  var PartsCancelledCol = TrackerTabColIndices[15];
  
  var OnTimeCol = TrackerTabColIndices[16];
  var ExcCol = TrackerTabColIndices[17];
  var LateCol = TrackerTabColIndices[18];
  var NotDefCol = TrackerTabColIndices[19];
  
  var PercOnTimeCol = TrackerTabColIndices[20];
  var PercExcCol = TrackerTabColIndices[21];
  var PercLateCol = TrackerTabColIndices[22];
  var PercNotDefCol = TrackerTabColIndices[23];
  
  var NoRFQPendInd = TrackerTabColIndices[24];
  var NoREQInd = TrackerTabColIndices[25];
  var NoPOInd = TrackerTabColIndices[26];
  var NoRecdInd = TrackerTabColIndices[27];
  
  //var VFInd = TrackerTabColIndices[28];
  var DaystoMRDInd = TrackerTabColIndices[29];
  
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
  
  // Logger.log(updatedEventData);
  // [[[2021 WS DV], [Interior, Anurag Bansal, 11/12/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6]], 
  // [[2021 WS DV], [Chassis, Anurag Bansal, 11/11/2020, 40, 62.5%, 62.5%, 25.0%, 37.5%, $0.00, 25, 25, 10, 15]]]
  
  for (var j=0; j<updatedEventData.length; j++) {
    if (updatedEventData.length !== 0) {
      for (var i=0; i<TrackerDataTabRange.length; i++) {
        if (updatedEventData[j][0] === TrackerDataTabRange[i][0]) {
          var Row = i+j+headerRow+1; 
          //#in RFQ Pending	#in REQ	#in PO	#in Rec'd
          TrackerDataTab.getRange(Row, TabCol+1, 1,1).setValue(updatedEventData[j][1][0]);
          TrackerDataTab.getRange(Row, PPPMEngCol+1, 1,1).setValue(updatedEventData[j][1][1]);
          TrackerDataTab.getRange(Row, MRDInd+1, 1,1).setValue(updatedEventData[j][1][2]);
          TrackerDataTab.getRange(Row, DaystoMRDInd+1, 1,1).setFormula(`=G${Row}-TODAY()`);
          TrackerDataTab.getRange(Row, TotalPartsCol+1, 1,1).setValue(updatedEventData[j][1][3]);
          
          TrackerDataTab.getRange(Row, PercReqCol+1, 1,1).setValue(updatedEventData[j][1][4]);
          TrackerDataTab.getRange(Row, PercPOCol+1, 1,1).setValue(updatedEventData[j][1][5]);
          TrackerDataTab.getRange(Row, PercRecdCol+1, 1,1).setValue(updatedEventData[j][1][6]);
          TrackerDataTab.getRange(Row, CostCol+1, 1,1).setValue(updatedEventData[j][1][7]);
          TrackerDataTab.getRange(Row, PercCancelledCol+1, 1,1).setValue(updatedEventData[j][1][8]);
          
          TrackerDataTab.getRange(Row, REQCol+1, 1,1).setValue(updatedEventData[j][1][9]);
          TrackerDataTab.getRange(Row, POCol+1, 1,1).setValue(updatedEventData[j][1][10]);
          TrackerDataTab.getRange(Row, PartsRecdCol+1, 1,1).setValue(updatedEventData[j][1][11]);
          TrackerDataTab.getRange(Row, PartsCancelledCol+1, 1,1).setValue(updatedEventData[j][1][12]);
          
          TrackerDataTab.getRange(Row, OnTimeCol+1, 1,1).setValue(updatedEventData[j][1][13]);
          TrackerDataTab.getRange(Row, ExcCol+1, 1,1).setValue(updatedEventData[j][1][14]);
          TrackerDataTab.getRange(Row, LateCol+1, 1,1).setValue(updatedEventData[j][1][15]);
          TrackerDataTab.getRange(Row, NotDefCol+1, 1,1).setValue(updatedEventData[j][1][16]);
          
          TrackerDataTab.getRange(Row, PercOnTimeCol+1, 1,1).setValue(updatedEventData[j][1][17]);
          TrackerDataTab.getRange(Row, PercExcCol+1, 1,1).setValue(updatedEventData[j][1][18]);
          TrackerDataTab.getRange(Row, PercLateCol+1, 1,1).setValue(updatedEventData[j][1][19]);
          TrackerDataTab.getRange(Row, PercNotDefCol+1, 1,1).setValue(updatedEventData[j][1][20]);
          
          TrackerDataTab.getRange(Row, NoRFQPendInd+1, 1,1).setFormula(`=I${Row}-O${Row}`);
          TrackerDataTab.getRange(Row, NoREQInd+1, 1,1).setFormula(`=I${Row}-P${Row}`);
          TrackerDataTab.getRange(Row, NoPOInd+1, 1,1).setFormula(`=I${Row}-Q${Row}`);
          TrackerDataTab.getRange(Row, NoRecdInd+1, 1,1).setFormula(`=I${Row}-R${Row}`);
          break;
        }
      }
    }
  }
  
}
