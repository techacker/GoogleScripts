/*  
        PBE PLM Function Library
        Author: Anurag Bansal 
        Version: 1.0.2  
        Release Date: 03/26/2021        
    
 -----------Change Log Notes:
 *********** First Release for BIW Group                            - 03/26/2021
 *********** First Release for Build Management System(BMS)         - 05/02/2021
 *********** Push event status from individual tracker files        - 05/04/2021
 *********** Edit Events functionality                              - 05/18/2021

*/

// Global Constants/Default Values if nothing is passed to function
let workloadfileURL = "https://docs.google.com/spreadsheets/d/1nzH0FyOkmhqm-RRVo376ZDiikNWu9MQSbPFSeTF5IWY/";
let addNeweventSheetName = "Add New Event";
let editEventSheetName = "Edit Event";
let mainEventsSheetName = "Events";
let trackersSheetName = "Tracker Data";
let trackersFolderName = "PPPM Shared Part Trackers";
let archiveFolderName = "Archive - PPPM Trackers";
let templateFileID = "1OQd5q8aI5jvSanxRAUe245HAgPHdX5qj3vvn1KwTmBc";
let menuName = "PPPM Tools";
let keyWord = "Event Title"
let eventTabLookupArray = ["Vehicle Family", "Event Type","Event Title","Earliest MRD","Event Status","Program Manager","Requestor","WBS Code",
  "Location of Event","Ship-to Code","Ship-to Address","Attention-to","Initial Part #s","Tracker URL"];
let newEventMandatoryFields = ["Model Year","Vehicle Family","Event Name", "Earliest MRD", "Event Type", "Program Manager"]
let trackerTabLookupArray = ["Vehicle Family","Event Title","Event Status","Program Manager","Earliest MRD","Initial Part #s","Event Type","Tracker URL"]


//Run when spreadsheet loads TEST
function onOpen(){
  createMenu(menuName); 
}

//Create menu dropdown
function createMenu(menuName) {

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(menuName);
  menu.addItem("Create Tracker","createNewTracker");
  menu.addItem("Edit Event Info","editEventInfo");
  menu.addItem("Refresh Event Status","getEventStatus");
  menu.addItem("Archive Trackers","archiveTrackers");
  menu.addToUi(); 
  

}
