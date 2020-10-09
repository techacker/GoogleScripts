// ---------------------------------------------------------------------- PPPM Workload Events Data Update Program --------------------------------------------------------------
// -------------------------------------------------------------------------      Author: Anurag Bansal        ------------------------------------------------------------------
// -------------------------------------------------------------------------          Version: 1.1.4           ------------------------------------------------------------------
// -------------------------------------------------------------------------      Only for PPPM Programs       ------------------------------------------------------------------
// -----------Change Log:
// -----------Create Event Tracker in PPPM Shared Parts Tracker Folder            ----- Completed 09/11/2020
// -----------Update Event URL in the Events Tab                                  ----- Completed 09/14/2020
// -----------Update Trackers Basic Event Info section from Events Tab            ----- Completed 09/17/2020
// -----------Get Event Tabs information from Trackers and update Tracker Tab     ----- Completed 09/21/2020
// -----------Updated Tracker Tab Template and corresponding code                 ----- Completed 09/24/2020
// -----------Bug fixes in Tracker Tab Scripts and utility to not count extra
//            rows at the bottom in Tracker's Summary Sheet                       ----- Completed 10/09/2020


//----------------------Start : Function to update tracker url field in Event Sheet

function getNewTrackerURL() {
  
  var workloadfileURL = "https://docs.google.com/spreadsheets/d/1lwDLj82hJWXi_6r7ec7s7BXSGL2C8MJkdxLkg3OsCUA/";
  //var workloadfileURL = "https://docs.google.com/spreadsheets/d/1TpNZ-fOasSRQN6JJRI9JfqWfgvHhVIi83YYnMDuTVX0/";  // Test PPPM Workload File
  var EventSheet = SpreadsheetApp.openByUrl(workloadfileURL).getSheetByName("Events");
  //var EventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  var lr = EventSheet.getLastRow();
  var lc = EventSheet.getLastColumn();
  var headerRow = getHeaderRow(EventSheet, "Event Title");
  var colIndices = getColIndex(EventSheet, headerRow);
  
  // Logger.log(colIndices);
  // [2.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 30.0]
  // VFInd, EventTitleInd, MRDColInd, EventStatusColInd, ProgMgrColInd, RequestorColInd, 
  // WBSColInd, LocColInd, ShiptoColInd, ShipAddColInd, AttnColInd, urlColInd
  
  var EventsHeader = EventSheet.getRange(headerRow, 1, 1, lc).getValues()[0];
  var dataRange = EventSheet.getRange(headerRow+1, 1, lr-headerRow, lc).getDisplayValues();
  var PPPMFolderName = "PPPM Shared Part Trackers";
  var ArchiveFolderName = "Archive - PPPM Trackers";
  
  var VFInd = colIndices[0];
  var EventTitleInd = colIndices[1];
  var MRDColInd = colIndices[2];
  var EventStatusColInd = colIndices[3];
  var RequestorColInd = colIndices[5];
  var WBSColInd = colIndices[6];
  var LocColInd = colIndices[7];
  var ShiptoColInd = colIndices[8];
  var ShipAddColInd = colIndices[9];
  var AttnColInd = colIndices[10];
  var urlColInd = colIndices[11];
  
  for (var i=0; i<dataRange.length; i++) { 
    var VF = dataRange[i][VFInd];
    var EventTitle = dataRange[i][EventTitleInd];
    var MRD = dataRange[i][MRDColInd];
    var EventStatus = dataRange[i][EventStatusColInd];
    var Requestor = dataRange[i][RequestorColInd]
    var WBSCode = dataRange[i][WBSColInd]
    var Location = dataRange[i][LocColInd];
    var Shipto = dataRange[i][ShiptoColInd];
    var ShipAddress = dataRange[i][ShipAddColInd];
    var Attention = dataRange[i][AttnColInd];
    var url = dataRange[i][urlColInd];
    
    if (EventStatus === "In-Process" && url === "") {
      var link = createNewTracker(VF.trim(), EventTitle.trim());
      EventSheet.getRange(i+headerRow+1, urlColInd+1, 1, 1).setValue(link);
      updateEventInfo(link, EventTitle, Requestor, WBSCode, Location, Shipto, ShipAddress, Attention);
    }
    else if (EventStatus === "Complete" && url !== "") {
      var id = url.split('/')[5];
      var ArchiveFolder = DriveApp.getFoldersByName(ArchiveFolderName);
      var trackerFile = DriveApp.getFileById(id);
      trackerFile.moveTo(ArchiveFolder.next());
    }
    else {
      updateEventInfo(url, EventTitle, Requestor, WBSCode, Location, Shipto, ShipAddress, Attention);
    }
    
  }
  
}

//----------------------End : Function to update tracker url field in Event Sheet



//----------------------Start : Function to create new Event Folders and Trackers within for each Event Title
//----------------------Don't call directly, gets called from getNewTrackerURL()

function createNewTracker(VehFam, EventTitle) {
  
  var folders = DriveApp.getFolders();
  var PPPMFolderName = "PPPM Shared Part Trackers";  
  
  // https://docs.google.com/spreadsheets/d/1OQd5q8aI5jvSanxRAUe245HAgPHdX5qj3vvn1KwTmBc/edit#gid=1502867507
  var TemplateFileID = "1OQd5q8aI5jvSanxRAUe245HAgPHdX5qj3vvn1KwTmBc";
  
  // Get folder names inside PPPM Shared Part Trackers folder
  var PPPMFolders = DriveApp.getFoldersByName(PPPMFolderName);
  var PPPMFolder = PPPMFolders.next();
  var VFFolders = PPPMFolder.getFolders();
  var VFFolderNames = [];
  while (VFFolders.hasNext()) {
    var VFFolder = VFFolders.next();
    var VFFolderName = VFFolder.getName();
    VFFolderNames.push(VFFolderName);
  }
  
  // Create a Vehicle Family Folder inside if it doesn't exist
  if (VFFolderNames.indexOf(VehFam) === -1) {
    PPPMFolder.createFolder(VehFam);
  }
  
  // Get file names inside Veh Family folder
  var VehFamFolders = DriveApp.getFoldersByName(VehFam);
  var VehFamFolder = VehFamFolders.next();
  var VehFamFiles = VehFamFolder.getFiles();
  var VehFamFileNames = [];
  while (VehFamFiles.hasNext()) {
    var VehFamFile = VehFamFiles.next();
    var VehFamFileName = VehFamFile.getName();
    VehFamFileNames.push(VehFamFileName);
  }
  
  
  // Create a new tracker inside if it doesn't exist
  if (VehFamFileNames.indexOf(EventTitle + " Tracker") === -1) {
    var newTracker = DriveApp.getFileById(TemplateFileID).makeCopy(EventTitle + " Tracker", VehFamFolder);
    var url = newTracker.getUrl();
    var link = parseURL(url);
    return link;
  }
  // Show an error message that it already exists and capture its url.
  else {
    Browser.msgBox("Alert!", "Tracker for " + EventTitle + " already exists.", Browser.Buttons.OK);
    var url = getTrackerURL(VehFamFileNames, EventTitle);
    var link = parseURL(url);
    return link;
  }
}

//----------------------End : Function to create new Event Folders and Trackers within for each Event Title



//----------------------Start : Update Event Info Section in trackers
function updateEventInfo(url, EventTitle, Requestor, WBSCode, Location, Shipto, ShipAddress, Attention) {
  
  var SummarySheet = SpreadsheetApp.openByUrl(url).getSheetByName("Summary");
  var lc = SummarySheet.getLastColumn();
  var lr = SummarySheet.getLastRow();
  
  Logger.log(lc, lr);
 
  // Update Event Tracker with Basic Event Info
  SummarySheet.getRange(1, 3).setValue(EventTitle);
  SummarySheet.getRange(2, 3).setValue(Requestor);
  SummarySheet.getRange(3, 3).setValue(WBSCode);
  SummarySheet.getRange(4, 3).setValue(Location);
  SummarySheet.getRange(5, 3).setValue(Shipto);
  SummarySheet.getRange(6, 3).setValue(ShipAddress);
  SummarySheet.getRange(7, 3).setValue(Attention);
  
}
