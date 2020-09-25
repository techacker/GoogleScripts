function collectFolderNames(foldersIter) {
  
  var folderNames = [];
  // Collect all folder names in the Drive.
  while (foldersIter.hasNext()) {
    var insideFolder = foldersIter.next();
    var insideFolderName = insideFolder.getName();
    folderNames.push(insideFolderName);
  }
  return folderNames;
}



function parseURL(url) {
  
  var urlArray = url.split('/');
  var id = urlArray[5];
  var link = url.split(id);
  var linktoURL = link[0] + id + '/';
  return linktoURL;
  
}



function getTrackerURL(VehFamFileNames, EventTitle) {
  
  for (var i=0; i<VehFamFileNames.length; i++) {
    var fileName = EventTitle + " Tracker"
    if (VehFamFileNames[i] === fileName) {
      var tracker = DriveApp.getFilesByName(fileName).next();
      var url = tracker.getUrl();
      var link = parseURL(url);
      return url;
    }
  }
}


// Getting the Header Row of any sheet with a give search key
function getHeaderRow(sheet, searchKey) {
  
  //var EventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  var lr = sheet.getLastRow();
  var lc = sheet.getLastColumn();
  var headerSearchRange = sheet.getRange(1, 1, 10, lc).getValues();
  
  for (var i=0; i<headerSearchRange.length; i++) {
    for (var j=0; j<lc; j++) {
      if (headerSearchRange[i][j] === searchKey) {
        var headerRow = i+1;
        break;
      }
    }
  }
  
  // If "Event Title" is not found, show an error message.
  if (headerRow === undefined) {
    Browser.msgBox("ERROR", "Can't find a header row with " + searchKey + ".", Browser.Buttons.OK);
    return null;
  }
  
  return headerRow;
}



function getColIndex(headerRow) {
  
  var EventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  var lr = EventSheet.getLastRow();
  var lc = EventSheet.getLastColumn();
  var headerSearchRange = EventSheet.getRange(1, 1, 10, lc).getValues();
  var headerRow = getHeaderRow(EventSheet, "Event Title");
  var EventsHeader = EventSheet.getRange(headerRow, 1, 1, lc).getValues()[0];
  var colIndices = [];
  
  // Logger.log("Event Headers", EventsHeader);
  // [Date Added, MY, VF, Event Name, Event Type, Event Title, Earliest MRD, Event Status, 
  // Program Manager, Requestor, WBS Code, Location of Event, Ship-to Code, Ship-to Address, Attention-to, Notes/ Exceptions List @ MRD, 
  // Initial Part #s, Actual Part #s, # RFQ, # Reqs, # POs, # complete, Cost, % Reqs, % POs, % Complete, 
  // # Parts On Time, # Parts on Exception, # Parts Late, Timing Status Not Defined, Tracker URL]
  
  var VFInd = colIndices.push(EventsHeader.indexOf("VF"));
  var EventTitleInd = colIndices.push(EventsHeader.indexOf("Event Title"));
  var MRDColInd = colIndices.push(EventsHeader.indexOf("Earliest MRD"));
  var EventStatusColInd = colIndices.push(EventsHeader.indexOf("Event Status"));
  var ProgMgrColInd = colIndices.push(EventsHeader.indexOf("Program Manager"));
  var RequestorColInd = colIndices.push(EventsHeader.indexOf("Requestor"));
  var WBSColInd = colIndices.push(EventsHeader.indexOf("WBS Code"));
  var LocColInd = colIndices.push(EventsHeader.indexOf("Location of Event"));
  var ShiptoColInd = colIndices.push(EventsHeader.indexOf("Ship-to Code"));
  var ShipAddColInd = colIndices.push(EventsHeader.indexOf("Ship-to Address"));
  var AttnColInd = colIndices.push(EventsHeader.indexOf("Attention-to"));
  var urlColInd = colIndices.push(EventsHeader.indexOf("Tracker URL"));
  
  // Logger.log(colIndices);
  // [0.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 39.0]
  
  return colIndices;
  
}

function getTrackerTabIndices(TrackerDataTab, headerRow) {
  
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerSearchRange = TrackerDataTab.getRange(1, 1, headerRow, lc).getValues();
  var TabHeader = TrackerDataTab.getRange(headerRow, 1, 1, lc).getValues()[0];
  var TrackerTabColIndices = [];
  
  // Logger.log(TabHeader);
  // [Event Title, Program Manager,	Event Status, Tab, PPPM Engineer, MRD, Total # of Parts, 
  // % REQ, % PO, % Received, Cost, % Cancelled, 
  // # REQ Submitted, # PO Issued, # Parts Received, # Cancelled, 
  // # On Time, # Exception, # Late, # Not Defined, 
  // % On Time, % Exception, % Late, % Not Defined, 
  // #in RFQ Pending, #in REQ, #in PO, #in Rec'd]
  
  var EventTitleInd = TrackerTabColIndices.push(TabHeader.indexOf("Event Title"));
  var PrgMgrInd = TrackerTabColIndices.push(TabHeader.indexOf("Program Manager"));
  var EventStatusInd = TrackerTabColIndices.push(TabHeader.indexOf("Event Status"));
  
  var TabInd = TrackerTabColIndices.push(TabHeader.indexOf("Tab"));
  var PPPMEngrInd = TrackerTabColIndices.push(TabHeader.indexOf("PPPM Engineer"));
  var MRDInd = TrackerTabColIndices.push(TabHeader.indexOf("MRD"));
  var TotalPartsInd = TrackerTabColIndices.push(TabHeader.indexOf("Total # of Parts"));
  
  var PercREQInd = TrackerTabColIndices.push(TabHeader.indexOf("% REQ"));
  var PercPOInd = TrackerTabColIndices.push(TabHeader.indexOf("% PO"));
  var PercRecdInd = TrackerTabColIndices.push(TabHeader.indexOf("% Received"));
  var CostInd = TrackerTabColIndices.push(TabHeader.indexOf("Cost"));
  var PercCancelledInd = TrackerTabColIndices.push(TabHeader.indexOf("% Cancelled"));
  
  var REQInd = TrackerTabColIndices.push(TabHeader.indexOf("# REQ Submitted"));
  var POIssuedInd = TrackerTabColIndices.push(TabHeader.indexOf("# PO Issued"));
  var RecdInd = TrackerTabColIndices.push(TabHeader.indexOf("# Parts Received"));
  var CancelledInd = TrackerTabColIndices.push(TabHeader.indexOf("# Cancelled"));
  
  var OnTimeInd = TrackerTabColIndices.push(TabHeader.indexOf("# On Time"));
  var ExceptionInd = TrackerTabColIndices.push(TabHeader.indexOf("# Exception"));
  var LateInd = TrackerTabColIndices.push(TabHeader.indexOf("# Late"));
  var NotDefInd = TrackerTabColIndices.push(TabHeader.indexOf("# Not Defined"));
  
  var PercOnTimeInd = TrackerTabColIndices.push(TabHeader.indexOf("% On Time"));
  var PercExceptionInd = TrackerTabColIndices.push(TabHeader.indexOf("% Exception"));
  var PercLateInd = TrackerTabColIndices.push(TabHeader.indexOf("% Late"));
  var PercNotDefInd = TrackerTabColIndices.push(TabHeader.indexOf("% Not Defined"));
  
  var NoRFQPendInd = TrackerTabColIndices.push(TabHeader.indexOf("#in RFQ Pending"));
  var NoREQInd = TrackerTabColIndices.push(TabHeader.indexOf("#in REQ"));
  var NoPOInd = TrackerTabColIndices.push(TabHeader.indexOf("#in PO"));
  var NoRecdInd = TrackerTabColIndices.push(TabHeader.indexOf("#in Rec'd"));
  
  // Logger.log(TrackerTabColIndices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 21.0, 22.0, 23.0, 24.0, 25.0, 26.0, 27.0]
  
  return TrackerTabColIndices;
  
  
  
}
