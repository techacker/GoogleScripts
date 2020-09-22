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
  
  // Logger.log(EventsHeader);
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
  // [2.0, 5.0, 6.0, 7.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 30.0]
  
  return colIndices;
  
}

function getTrackerTabIndices(TrackerDataTab, headerRow) {
  
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerSearchRange = TrackerDataTab.getRange(1, 1, headerRow, lc).getValues();
  var TabHeader = TrackerDataTab.getRange(headerRow, 1, 1, lc).getValues()[0];
  var TrackerTabColIndices = [];
  
  // Logger.log(TabHeader);
  // [Event Title, Tab, PPPM Engineer, MRD, Total No. of Parts, % Reqs, % PO, % Rec'd, % RFQ Pending, 
  // # RFQ Sent, # REQ Submitted, # PO Issued, # Parts Received, Cost, % On Time, # Parts On Time, 
  // # Parts on Exception, # Parts Late, Timing Status Not Defined]
  
  var EventTitleInd = TrackerTabColIndices.push(TabHeader.indexOf("Event Title"));
  var TabInd = TrackerTabColIndices.push(TabHeader.indexOf("Tab"));
  var PPPMEngineerInd = TrackerTabColIndices.push(TabHeader.indexOf("PPPM Engineer"));
  var MRDInd = TrackerTabColIndices.push(TabHeader.indexOf("MRD"));
  var TotalPartsInd = TrackerTabColIndices.push(TabHeader.indexOf("Total No. of Parts"));
  var PercREQInd = TrackerTabColIndices.push(TabHeader.indexOf("% Reqs"));
  var PercPOInd = TrackerTabColIndices.push(TabHeader.indexOf("% PO"));
  var PercRecdInd = TrackerTabColIndices.push(TabHeader.indexOf("% Rec'd"));
  var PercRFQPendingInd = TrackerTabColIndices.push(TabHeader.indexOf("% RFQ Pending"));
  var RFQSentInd = TrackerTabColIndices.push(TabHeader.indexOf("# RFQ Sent"));
  var REQInd = TrackerTabColIndices.push(TabHeader.indexOf("# REQ Submitted"));
  var POIssuedInd = TrackerTabColIndices.push(TabHeader.indexOf("# PO Issued"));
  var RecdInd = TrackerTabColIndices.push(TabHeader.indexOf("# Parts Received"));
  var CostInd = TrackerTabColIndices.push(TabHeader.indexOf("Cost"));
  
  // Logger.log(TrackerTabColIndices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0]
  
  return TrackerTabColIndices;
  
  
  
}
