function lookupColumns(url=workloadfileURL, sheetName=mainEventsSheetName,stringToFind=keyWord,colArray=eventTabLookupArray) {
  let as = SpreadsheetApp.openByUrl(url).getSheetByName(sheetName);
  let lc = as.getLastColumn();
  let textFinder = as.getDataRange().createTextFinder(stringToFind).findNext();
  let headerRow = textFinder.getRow();
  let headers = as.getRange(headerRow,1,1,lc).getValues()[0];
  let HeaderIndexObj = {};

  // Create an object with Header Index Columns for every header entry
  headers.forEach(item => {
    HeaderIndexObj[item] = headers.indexOf(item);
  })

  // Logger.log(HeaderIndexObj)
  // {Event Status=7.0, % Exception=31.0, MY=2.0, Total # of Parts=16.0, Initial Part #s=39.0, Earliest MRD=6.0, # REQ Submitted=22.0, % Late=32.0,   Ship-to Code=12.0, Program Manager=8.0, #in Rec'd=37.0, % PO=18.0, # Not Defined=29.0, Event Title=5.0, VF=1.0, #in RFQ Pending=34.0, % Not Defined=33.0, Cost=20.0, # Exception=27.0, Date Added=38.0, # Cancelled=25.0, % On Time=30.0, % REQ=17.0, Event Name=3.0, # Parts Received=24.0, Notes=15.0, Location of Event=11.0, # PO Issued=23.0, Attention-to=14.0, % Cancelled=21.0, #in REQ=35.0, Tracker URL=40.0, WBS Code=10.0, Requestor=9.0, Ship-to Address=13.0, % Received=19.0, Event Type=4.0, # Late=28.0, #in PO=36.0, # On Time=26.0}

  let LookupObj = {};
  colArray.forEach(val => {
    if (val in HeaderIndexObj) {
      LookupObj[val] = HeaderIndexObj[val];
    }
  })
  //Logger.log(LookupObj)
  return LookupObj
  
}


function parseURL_(url) {
  
  var urlArray = url.split('/');
  var id = urlArray[5];
  var link = url.split(id);
  var linktoURL = link[0] + id + '/';
  return linktoURL;
  
}


function getTrackerURL_(VehFamFileNames, EventTitle) {
  
  for (var i=0; i<VehFamFileNames.length; i++) {
    var fileName = EventTitle + " Tracker"
    if (VehFamFileNames[i] === fileName) {
      var tracker = DriveApp.getFilesByName(fileName).next();
      var url = tracker.getUrl();
      var link = parseURL_(url);
      return url;
    }
  }
}


// Getting the Header Row of any sheet with a given search key
function getHeaderRow_(sheet, searchKey) {
  
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


function getColIndex_(EventSheet, headerRow) {
  
  var lr = EventSheet.getRange("A1").getDataRegion().getLastRow();
  var lc = EventSheet.getLastColumn();
  var headerSearchRange = EventSheet.getRange(1, 1, 10, lc).getValues();
  var EventsHeader = EventSheet.getRange(headerRow, 1, 1, lc).getValues()[0];
  var colIndices = [];
  
  // Logger.log("Event Headers", EventsHeader);
  // [VF, MY, Event Name, Event Type, Event Title, Earliest MRD, Event Status, 
  // Program Manager, Requestor, WBS Code, Location of Event, Ship-to Code, Ship-to Address, 
  // Attention-to, Notes, Total # of Parts, % REQ, % PO, % Received, Cost, % Cancelled, 
  // # REQ Submitted, # PO Issued, # Parts Received, # Cancelled, # On Time, # Exception, # Late, # Not Defined, 
  // % On Time, % Exception, % Late, % Not Defined, #in RFQ Pending, #in REQ, #in PO, #in Rec'd, 
  // Date Added, Initial Part #s, Tracker URL]
  
  var VFInd = colIndices.push(EventsHeader.indexOf("VF"));
  var EventTypeInd = colIndices.push(EventsHeader.indexOf("Event Type"));
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
  var PartCountInd = colIndices.push(EventsHeader.indexOf("Initial Part #s"));
  var urlColInd = colIndices.push(EventsHeader.indexOf("Tracker URL"));
  
  // Logger.log(colIndices);
  // [0.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 38.0, 39.0]
  
  return colIndices;
  
}


function getTrackerTabIndices_(TrackerDataTab, headerRow) {
  
  var lr = TrackerDataTab.getLastRow();
  var lc = TrackerDataTab.getLastColumn();
  var headerSearchRange = TrackerDataTab.getRange(1, 1, headerRow, lc).getValues();
  var TabHeader = TrackerDataTab.getRange(headerRow, 1, 1, lc).getValues()[0];
  var TrackerTabColIndices = [];
  
  // Logger.log(TabHeader);
  // [VF, Event Title, Tab, Event Title - Tab, Event Status, Program Manager, 
  // PPPM Engineer, MRD, Days until MRD, Total # of Parts, 
  // % REQ, % PO, % Received, Cost, % Cancelled, 
  // # REQ Submitted, # PO Issued, # Parts Received, # Cancelled, 
  // # On Time, # Exception, # Late, # Not Defined, 
  // % On Time, % Exception, % Late, % Not Defined, 
  // #in RFQ Pending, #in REQ, #in PO, #in Rec'd, Event Type, Tracker URL]
  
  var VFInd = TrackerTabColIndices.push(TabHeader.indexOf("VF"));
  var EventTitleInd = TrackerTabColIndices.push(TabHeader.indexOf("Event Title"));
  var TabInd = TrackerTabColIndices.push(TabHeader.indexOf("Tab"));
  var TitleTabInd = TrackerTabColIndices.push(TabHeader.indexOf("Event Title - Tab"));
  var EventStatusInd = TrackerTabColIndices.push(TabHeader.indexOf("Event Status"));
  var PrgMgrInd = TrackerTabColIndices.push(TabHeader.indexOf("Program Manager"));
  
  var PPPMEngrInd = TrackerTabColIndices.push(TabHeader.indexOf("PPPM Engineer"));
  var MRDInd = TrackerTabColIndices.push(TabHeader.indexOf("MRD"));
  var DaystoMRDInd = TrackerTabColIndices.push(TabHeader.indexOf("Days until MRD"));
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
  var EventTypeInd = TrackerTabColIndices.push(TabHeader.indexOf("Event Type"));
  var TrackerURLInd = TrackerTabColIndices.push(TabHeader.indexOf("Tracker URL"));
    
  // Logger.log(TrackerTabColIndices);
  // [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 
  // 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 
  // 21.0, 22.0, 23.0, 24.0, 25.0, 26.0, 27.0, 28.0, 29.0, 30.0, 31.0, 32.0]
  return TrackerTabColIndices;
  
}
