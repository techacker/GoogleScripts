/*  
        PBE PLM Function Library
        Author: Anurag Bansal 
        Version: 1.5.0  
        Release Date: 05/06/2021        
    
 -----------Change Log Notes:
 
*********** First Release for Build Management System(BMS)         - 05/02/2021
*********** Added functionality to edit build                      - 05/11/2021
*********** Update Status from Individual Builds                   - 05/20/2021
*********** Update Red Item Details for builds                     - 06/02/2021
*********** New Build Announcement for builds                      - 06/04/2021
*********** Change Bulletin for builds                             - 06/07/2021
*********** Update Shelly's BOM file for build updates             - 08/09/2021
*********** Update BOM file BICEE tabs with new Builds             - 08/23/2021
*********** Update Build KPI data for every build                  - 09/16/2021
*********** Update BOM file with data validation                   - 09/21/2021
*********** Edit KPI Assessment                                    - 09/28/2021


*/

// Global Constants/Default Values if nothing is passed to function
let buildsumfileURL = "https://docs.google.com/spreadsheets/d/1ANkq-NE26z76w0auBlyt-T2CYY3lhROsHOtPhVbMygs/";
let addNewBuildSheetName = "Add New Build";
let editBuildSheetName = "Edit Build Info";
let buildsListSheetName = "PBE Build List";
let buildDetailsSheetName = "Build Detail Status";
let redItemsSheetName = "Red Item Details";
let trackersFolderName = "PBE Build Trackers";
let kpiSheetName = "Build KPI Data";
let kpiAssessmentSheetName = "Build KPI Assessment";
let archiveFolderName = "Archive - PBE Build Trackers";
let templateID = "1cKA2XSCYWpnqvSEkYsFvuigLUV83oVGnarK6jqPLYr8";
let parentBOMFileID = '1LK8ef_bQ4sIQPPLDE0orgfSguD2Paburk9hf_nGAvN4'
let menuName = "PBE Tools";
let keyWord = "Build Tracker URL"
let mandatoryFields = ["Model Year","Vehicle Family","Build Type","Build Phase","Vehicle Quantity","TCF MRD","TCF Build Start"]
let buildDetailLookupArray = ["Vehicle Family","Build Title","Build Status","Build Lead","TCF MRD","Build Type","Build Tracker URL"]


//Run when spreadsheet loads TEST
function onOpen(){
  createMenu(menuName); 
}

//Create menu dropdown
function createMenu(menuName) {

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(menuName);
  menu.addItem("Add New Build","createBuildTracker");
  menu.addSeparator();

  const submenu1 = ui.createMenu("Edit Build Info");
  submenu1.addItem("Edit Build Info","editBuild");
  submenu1.addItem("Edit KPI Assessment","editKPIData");
  // Add Submenu to Menu
  menu.addSubMenu(submenu1);

  menu.addSeparator();
  // Create Submenu
  const submenu = ui.createMenu("Refresh Data");
  submenu.addItem("Builds Status","getBuildStatus");
  submenu.addItem("Red Item Report","getRedItemsStatus")
  submenu.addItem("Builds KPI","getKPIStatus")
  submenu.addItem("Everything","refreshEverything")

  // Add Submenu to Menu
  menu.addSubMenu(submenu);

  //menu.addItem("Archive Build Trackers","archiveTrackers");
  menu.addToUi(); 
  
}
