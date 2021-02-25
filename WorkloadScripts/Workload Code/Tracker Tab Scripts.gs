function updateTrackerTab() {
  
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/"; // Test Workload File
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var EventSheet = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Events");
  var lr = EventSheet.getLastRow();
  var lc = EventSheet.getLastColumn();
  
  var headerRow = getHeaderRow(EventSheet, "Tracker URL");
  
  // Get col indexes from Events Sheet
  var colIndices = getColIndex(EventSheet, headerRow);
  
  var VFInd = colIndices[0];
  var EventTypeInd = colIndices[1];
  var EventTitleInd = colIndices[2];
  var EventMRDInd = colIndices[3];
  var EventStatusInd = colIndices[4];
  var ProgMgrInd = colIndices[5];
  var PartCountColInd = colIndices[12];
  var urlColInd = colIndices[13];
  
  
  var trackerURLArray = EventSheet.getRange(headerRow+1, urlColInd+1, lr-headerRow, 1).getValues();  
  var BasicEventInfo = [];
  var VF, EventTitle, EventType, EventMRD, EventStatus, ProgMgr, PartCount, URL;
  var SummarySheet, sslc, sslr, SummaryRange, ssheaderRow, eventTabsData;
  
  // Get the tracker URL from Events Tab
  
  for (var i=0; i<trackerURLArray.length; i++) {
    if (trackerURLArray[i][0] !== "") {
      VF = EventSheet.getRange(i+ headerRow+1, VFInd+1, 1, 1).getValue();
      EventTitle = EventSheet.getRange(i + headerRow+1, EventTitleInd+1, 1, 1).getValue();
      EventType = EventSheet.getRange(i + headerRow+1, EventTypeInd+1, 1, 1).getValue();
      EventMRD = EventSheet.getRange(i + headerRow+1, EventMRDInd+1, 1, 1).getDisplayValue();
      EventStatus = EventSheet.getRange(i+ headerRow+1, EventStatusInd+1, 1, 1).getValue();
      ProgMgr = EventSheet.getRange(i + headerRow+1, ProgMgrInd+1, 1, 1).getValue();
      PartCount = EventSheet.getRange(i+ headerRow+1, PartCountColInd+1, 1, 1).getValue();
      URL = trackerURLArray[i][0];
      
      // Get Information from Event Tracker's Summary Tab only if the event is "In-Process"
      if (EventStatus === "In-Process") {
        SummarySheet = SpreadsheetApp.openByUrl(URL).getSheetByName("Summary");
        sslc = SummarySheet.getLastColumn();
        sslr = SummarySheet.getRange("A1").getDataRegion().getLastRow();
        SummaryRange = SummarySheet.getRange(1, 1, sslr, sslc).getValues();
        ssheaderRow = getHeaderRow(SummarySheet, "Tab");
        
        // Events Tab Data from trackers file
        eventTabsData = SummarySheet.getRange(ssheaderRow+1, 1, sslr-ssheaderRow, sslc).getDisplayValues();
        var infoCol = 3;
        var eventInfoData = SummarySheet.getRange(1, infoCol, ssheaderRow-1, 1).getDisplayValues();
        
        // Add Events to Tracker Data Tab
        addEventTitles(EventTitle, VF, ProgMgr, EventType, URL, eventInfoData, eventTabsData);
        // Update Event information in Tracker Data Tab
        updateEventsData(EventTitle, EventMRD, PartCount, eventTabsData);
      } 
    }
  } 
}


function addEventTitles(EventTitle, VehFam, ProgManager, EventType, URL, eventInfoData, eventTabsData) {
  
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/"; // Test Workload File
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var TrackerDataTab = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Tracker Data");
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerRow = getHeaderRow(TrackerDataTab, "Tab");
  var TrackerTabColIndices = getTrackerTabIndices(TrackerDataTab, headerRow);
  var TrackerDataHeaders = TrackerDataTab.getRange(headerRow, 1, 1, lc).getDisplayValues();
  var TrackerDataTabRange = TrackerDataTab.getRange(headerRow+1, 1, lr, lc).getDisplayValues();
  
  // Logger.log(TrackerTabColIndices)
  // [VF, Event Title, Tab, Event Title - Tab, Event Status, Program Manager, 
  // PPPM Engineer, MRD, Days until MRD, Total # of Parts, 
  // % REQ, % PO, % Received, Cost, % Cancelled, 
  // # REQ Submitted, # PO Issued, # Parts Received, # Cancelled, 
  // # On Time, # Exception, # Late, # Not Defined, 
  // % On Time, % Exception, % Late, % Not Defined, 
  // #in RFQ Pending, #in REQ, #in PO, #in Rec'd, Tracker URL
  
  // [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 
  // 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 
  // 21.0, 22.0, 23.0, 24.0, 25.0, 26.0, 27.0, 28.0, 29.0, 30.0, 31.0]
  
  var VFCol = TrackerTabColIndices[0];
  var EventTileCol = TrackerTabColIndices[1];
  var TabCol = TrackerTabColIndices[2];
  var TitleTabCol = TrackerTabColIndices[3];
  var EventStatusCol = TrackerTabColIndices[4];
  var PrgMgrCol = TrackerTabColIndices[5];  
  var DaystoMRDCol = TrackerTabColIndices[8];
  var EventTypeCol = TrackerTabColIndices[31];
  var TrackerURLCol = TrackerTabColIndices[32];
  
  // Get Event Titles from Workload file
  var TrackerTabEventTitles = [];
  var TrackerTabSheetNames = [];
  for (var i=0; i<TrackerDataTabRange.length; i++) {
    TrackerTabEventTitles.push(TrackerDataTabRange[i][EventTileCol]);
    TrackerTabSheetNames.push(TrackerDataTabRange[i][TabCol]);
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
  
  // [2022 WS Rollover Spare Parts, Chris Sleiman, VD00228_VIVP_SIMP, DOW, 1256, 12501 Chrysler Dr, Detroit, MI, Chris Sleiman]
  
  // Logger.log(eventTabsData);
  // [[Chassis, Anurag Bansal, 11/11/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6], 
  // [Interior, Anurag Bansal, 11/12/2020, 25, 60.0%, 40.0%, 8.0%, 24.0%, $0.00, 15, 10, 2, 6], 
  // [MASTER, Name (dropdown), 09/15/2020, 0, 0.0%, 0.0%, 0.0%, 0.0%, $0.00, 0, 0, 0, 0]]
  
  var tabCount = 0;
  
  // Get number of tabs in Event Tracker Sheet
  for (var i=0; i<eventTabsData.length; i++) {
    if (eventTabsData[i][0].toUpperCase() !== "MASTER" && eventTabsData[i][0].toUpperCase() !== "OVERALL EVENT STATUS" 
    && eventTabsData[i][0] !== "") {
      // Include Event Title in the Event Data
      tabCount += 1;
    }
  }
  
  // Count how many times the event title appears in Workload Tracker Data Tab 
  var eventRowCount = 0;
  for (var i=0; i<TrackerTabEventTitles.length; i++) {
    if (TrackerTabEventTitles[i] === EventTitle && TrackerTabEventTitles[i] !== "") {
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
  
  // 2022 MP PS QAF (QAF) 0.0 8.0 null
  // 2022 VF IP DV Test 0.0 0.0 null
  
  // Add rows if the event doesn't exist in the tracker tab
  if (eventRowCount === 0 && tabCount !== 0) {
    var addRow = TrackerDataTab.getRange(headerRow+1, 1, tabCount, lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(headerRow+1, EventTileCol+1, tabCount, 1).setValue(EventTitle);
    TrackerDataTab.getRange(headerRow+1, VFCol+1, tabCount, 1).setValue(VehFam);
    TrackerDataTab.getRange(headerRow+1, PrgMgrCol+1, tabCount, 1).setValue(ProgManager);
    TrackerDataTab.getRange(headerRow+1, EventTypeCol+1, tabCount, 1).setValue(EventType);
    TrackerDataTab.getRange(headerRow+1, TrackerURLCol+1, tabCount, 1).setValue(URL);
    //TrackerDataTab.getRange(headerRow+1, EventStatusCol+1, tabCount, 1).setValue(EventStatus);
  }
  // Add row if there are additional tabs are added
  else if (eventRowCount < tabCount) {
    var addRow = TrackerDataTab.getRange(EventRow, 1, tabCount - eventRowCount, lc).insertCells(SpreadsheetApp.Dimension.ROWS);
    TrackerDataTab.getRange(EventRow, EventTileCol+1, tabCount - eventRowCount, 1).setValue(EventTitle);
    TrackerDataTab.getRange(EventRow, VFCol+1, tabCount - eventRowCount, 1).setValue(VehFam);
    TrackerDataTab.getRange(EventRow, PrgMgrCol+1, tabCount - eventRowCount, 1).setValue(ProgManager);
    TrackerDataTab.getRange(EventRow, EventTypeCol+1, tabCount - eventRowCount, 1).setValue(EventType);
    TrackerDataTab.getRange(EventRow, TrackerURLCol+1, tabCount - eventRowCount, 1).setValue(URL);
    //TrackerDataTab.getRange(EventRow, EventStatusCol+1, tabCount - eventRowCount, 1).setValue(EventStatus);
  } 
  // Delete rows if there were additional tabs earlier but removed later on.
  else if (eventRowCount > tabCount) {
    var deleteRow = TrackerDataTab.getRange(EventRow, 1, eventRowCount - tabCount, lc).deleteCells(SpreadsheetApp.Dimension.ROWS);
  }
  // Ignore if there is only MASTER tab name
  else if (eventRowCount === 0) {
    
  }
  // Update the event status column
  else {
    //TrackerDataTab.getRange(EventRow, EventStatusCol+1, eventRowCount, 1).setValue(EventStatus);
  }

}

function updateEventsData(EventTitle, EventMRD, PartCount, eventTabsData) {
  
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/"; // Test Workload File
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  var TrackerDataTab = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Tracker Data");
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerRow = getHeaderRow(TrackerDataTab, "Tab");
  var TrackerTabColIndices = getTrackerTabIndices(TrackerDataTab, headerRow);
  var TrackerDataTabRange = TrackerDataTab.getRange(headerRow+1, 1, lr, lc).getDisplayValues();
  
  // Logger.log(TrackerTabColIndices)
  // [VF, Event Title, Tab, Event Title - Tab, Event Status, Program Manager, PPPM Engineer, MRD, Days until MRD, Total # of Parts, % REQ,
  // % PO, % Received, Cost, % Cancelled, # REQ Submitted, # PO Issued, # Parts Received, # Cancelled, # On Time, # Exception,
  // # Late, # Not Defined. % On Time, % Exception, % Late, % Not Defined, #in RFQ Pending, #in REQ, #in PO, #in Rec'd,
  // Event Type, Tracker URL]

  //[0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 
  // 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 
  // 21.0, 22.0, 23.0, 24.0, 25.0, 26.0, 27.0, 28.0, 29.0, 30.0, 
  // 31.0, 32.0]
  
  var VFInd = TrackerTabColIndices[0];
  var EventTileCol = TrackerTabColIndices[1];
  var TabCol = TrackerTabColIndices[2];
  var TitleTabCol = TrackerTabColIndices[3];
  var EventStatusCol = TrackerTabColIndices[4];
  var PrgMgrCol = TrackerTabColIndices[5];
  
  var PPPMEngCol = TrackerTabColIndices[6];
  var MRDInd = TrackerTabColIndices[7];
  var DaystoMRDInd = TrackerTabColIndices[8];
  var TotalPartsCol = TrackerTabColIndices[9];
  
  var PercReqCol = TrackerTabColIndices[10];
  var PercPOCol = TrackerTabColIndices[11];
  var PercRecdCol = TrackerTabColIndices[12];
  var CostCol = TrackerTabColIndices[13];
  var PercCancelledCol = TrackerTabColIndices[14];
  
  var REQCol = TrackerTabColIndices[15];
  var POCol = TrackerTabColIndices[16];
  var PartsRecdCol = TrackerTabColIndices[17];
  var PartsCancelledCol = TrackerTabColIndices[18];
  
  var OnTimeCol = TrackerTabColIndices[19];
  var ExcCol = TrackerTabColIndices[20];
  var LateCol = TrackerTabColIndices[21];
  var NotDefCol = TrackerTabColIndices[22];
  
  var PercOnTimeCol = TrackerTabColIndices[23];
  var PercExcCol = TrackerTabColIndices[24];
  var PercLateCol = TrackerTabColIndices[25];
  var PercNotDefCol = TrackerTabColIndices[26];
  
  var NoRFQPendInd = TrackerTabColIndices[27];
  var NoREQInd = TrackerTabColIndices[28];
  var NoPOInd = TrackerTabColIndices[29];
  var NoRecdInd = TrackerTabColIndices[30];
  
  var updatedEventData = [];
  var tabCount = 0;
  
  // Get number of tabs in Event Tracker Sheet
  for (var i=0; i<eventTabsData.length; i++) {
    if (eventTabsData[i][0].toUpperCase() !== "MASTER" && eventTabsData[i][0].toUpperCase() !== "OVERALL EVENT STATUS" 
    && eventTabsData[i][0] !== "") {
      // Include Event Title in the Event Data
      updatedEventData.unshift([EventTitle, eventTabsData[i]]);
      tabCount += 1;
    }
  }
  
  // Logger.log(updatedEventData);
  // [[2021 WD Media Spare Parts, [SpareParts, Jim Crile, 10/26/2020, 33, 100.0%, 75.8%, 6.1%, $0.00, 0.0%, 33, 25, 2, 0, 0, 0, 0, 33, 0.0%, 0.0%, 0.0%, 100.0%]]]
  // [[2022 WL74 PHEV Pre-VP QAF, [PPPM, Anurag Bansal, 09/11/2020, 21, 100%, 100%, 100%, $7,223.96, 0.0%, 21, 21, 21, 0, 21, 0, 0, 0, 100.0%, 0.0%, 0.0%, 0.0%]]]
  
  var EventTitlesInTrackerDataTab = TrackerDataTab.getRange(headerRow+1, EventTileCol+1, lr-1, 1).getValues();
  
  for (var j=0; j<updatedEventData.length; j++) {
    if (updatedEventData.length !== 0) {
      for (var i=0; i<EventTitlesInTrackerDataTab.length; i++) {
        if (updatedEventData[j][0] === EventTitlesInTrackerDataTab[i][0]) {
          var Row = i+j+headerRow+1; 
          TrackerDataTab.getRange(Row, TabCol+1, 1,1).setValue(updatedEventData[j][1][0]);
          TrackerDataTab.getRange(Row, PPPMEngCol+1, 1,1).setValue(updatedEventData[j][1][1]);
          
          // In case it is a new tracker, use the Event MRD from Event Tab
          if (updatedEventData[j][1][2] !== "<MRD>") { 
            TrackerDataTab.getRange(Row, MRDInd+1, 1,1).setValue(updatedEventData[j][1][2]);
            TrackerDataTab.getRange(Row, MRDInd+1, 1,1).setBackground("White").setFontColor("Black");
          } else {
            TrackerDataTab.getRange(Row, MRDInd+1, 1,1).setValue(EventMRD);
            TrackerDataTab.getRange(Row, MRDInd+1, 1,1).setBackground("Purple").setFontColor("White");
          }
          
          TrackerDataTab.getRange(Row, EventStatusCol+1, 1,1).setValue(updatedEventData[j][1][3]);
          
          // In case no one is assigned to this event, use the Parts Count from Event Tab
          if (updatedEventData[j][1][1] !== "<PPPM ENGINEER>") { 
            TrackerDataTab.getRange(Row, TotalPartsCol+1, 1,1).setValue(updatedEventData[j][1][4]);
            TrackerDataTab.getRange(Row, TotalPartsCol+1, 1,1).setBackground("White").setFontColor("Black");
          } else {
            TrackerDataTab.getRange(Row, TotalPartsCol+1, 1,1).setValue(PartCount);
            TrackerDataTab.getRange(Row, TotalPartsCol+1, 1,1).setBackground("Purple").setFontColor("White");
          }
          
          TrackerDataTab.getRange(Row, PercReqCol+1, 1,1).setValue(updatedEventData[j][1][5]);
          TrackerDataTab.getRange(Row, PercPOCol+1, 1,1).setValue(updatedEventData[j][1][6]);
          TrackerDataTab.getRange(Row, PercRecdCol+1, 1,1).setValue(updatedEventData[j][1][7]);
          
          TrackerDataTab.getRange(Row, CostCol+1, 1,1).setValue(updatedEventData[j][1][8]);
          TrackerDataTab.getRange(Row, PercCancelledCol+1, 1,1).setValue(updatedEventData[j][1][9]);
          TrackerDataTab.getRange(Row, REQCol+1, 1,1).setValue(updatedEventData[j][1][10]);
          TrackerDataTab.getRange(Row, POCol+1, 1,1).setValue(updatedEventData[j][1][11]);
          
          TrackerDataTab.getRange(Row, PartsRecdCol+1, 1,1).setValue(updatedEventData[j][1][12]);
          TrackerDataTab.getRange(Row, PartsCancelledCol+1, 1,1).setValue(updatedEventData[j][1][13]);
          TrackerDataTab.getRange(Row, OnTimeCol+1, 1,1).setValue(updatedEventData[j][1][14]);
          TrackerDataTab.getRange(Row, ExcCol+1, 1,1).setValue(updatedEventData[j][1][15]);
          
          TrackerDataTab.getRange(Row, LateCol+1, 1,1).setValue(updatedEventData[j][1][16]);
          TrackerDataTab.getRange(Row, NotDefCol+1, 1,1).setValue(updatedEventData[j][1][17]);
          TrackerDataTab.getRange(Row, PercOnTimeCol+1, 1,1).setValue(updatedEventData[j][1][18]);
          TrackerDataTab.getRange(Row, PercExcCol+1, 1,1).setValue(updatedEventData[j][1][19]);
          
          TrackerDataTab.getRange(Row, PercLateCol+1, 1,1).setValue(updatedEventData[j][1][20]);
          TrackerDataTab.getRange(Row, PercNotDefCol+1, 1,1).setValue(updatedEventData[j][1][21]);
   
          TrackerDataTab.getRange(Row, TitleTabCol+1, 1,1).setFormula(`=CONCATENATE(B${Row}, " - ", C${Row})`);
          TrackerDataTab.getRange(Row, DaystoMRDInd+1, 1,1).setFormula(`=IFERROR(H${Row}-TODAY(),365)`);
          TrackerDataTab.getRange(Row, NoRFQPendInd+1, 1,1).setFormula(`=J${Row}-P${Row}`);
          TrackerDataTab.getRange(Row, NoREQInd+1, 1,1).setFormula(`=P${Row}-Q${Row}`);
          TrackerDataTab.getRange(Row, NoPOInd+1, 1,1).setFormula(`=Q${Row}-R${Row}`);
          TrackerDataTab.getRange(Row, NoRecdInd+1, 1,1).setFormula(`=R${Row}`);
          break;
        }
      }
    }
  }
}
