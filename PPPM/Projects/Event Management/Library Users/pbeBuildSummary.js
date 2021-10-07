// Global Constants/Default Values if nothing is passed to function
let buildsumfileURL = "https://docs.google.com/spreadsheets/d/1OAiwIHKlcl7gSdDSsC6qNhV6rlKf6CQCxmY6qejme9Q/";
let addNewBuildSheetName = "Add New Build";
let editBuildSheetName = "Edit Build Info";
let buildsListSheetName = "PBE Build List";
let redItemsSheetName = "Red Item Details";
let buildDetailsSheetName = "Build Detail Status";
let trackersFolderName = "PBE Build Trackers";
let kpiSheetName = "Build KPI Data";
let kpiAssessmentSheetName = "Build KPI Assessment";
let archiveFolderName = "Archive - PBE Build Trackers";
let templateID = "1cKA2XSCYWpnqvSEkYsFvuigLUV83oVGnarK6jqPLYr8";
let parentBOMFileID = '19y4M-vR7KWivHALzdiA0xMMhNaR7ZbxK5eMKMaLXTxQ'
let menuName = "PBE Tools";
let keyWord = "Build Tracker URL"
let mandatoryFields = ["Model Year","Vehicle Family","Build Type","Build Phase","Vehicle Quantity","TCF MRD","TCF Build Start"]
let buildDetailLookupArray = ["Vehicle Family","Build Title","Build Status","Build Lead","TCF MRD","Build Type","Build Tracker URL"]

function onOpen() {
  BuildSummaryLibrary.createMenu(menuName)
}

function createBuildTracker() {
  BuildSummaryLibrary.createBuildTracker(buildsumfileURL,addNewBuildSheetName,buildsListSheetName,keyWord,trackersFolderName,templateID,mandatoryFields,parentBOMFileID)  
}

function getEventStatus() {
  BuildSummaryLibrary.getEventStatus(buildsumfileURL,buildsListSheetName,keyWord,buildDetailLookupArray)
}

function archiveTrackers() {
  BuildSummaryLibrary.archiveTrackers(buildsumfileURL,buildsListSheetName,keyWord,archiveFolderName)
}

function editBuild() {
  BuildSummaryLibrary.editBuild(buildsumfileURL,editBuildSheetName,buildsListSheetName,keyWord,parentBOMFileID)
  
}

function getBuildStatus() {
  BuildSummaryLibrary.getBuildStatus(buildsumfileURL,buildsListSheetName,buildDetailsSheetName,keyWord)
}

function getRedItemsStatus() {
  BuildSummaryLibrary.getRedItemsStatus(buildsumfileURL,buildsListSheetName,redItemsSheetName,keyWord)
}

function getKPIStatus() {
  BuildSummaryLibrary.getKPIStatus(buildsumfileURL,buildsListSheetName,keyWord,buildDetailLookupArray,kpiSheetName)
}

function editKPIData() {
  BuildSummaryLibrary.editKPIData(buildsumfileURL,kpiAssessmentSheetName,buildsListSheetName,keyWord)
}

function refreshEverything() {
  BuildSummaryLibrary.refreshEverything(buildsumfileURL,buildsListSheetName,buildDetailsSheetName,keyWord,redItemsSheetName,kpiSheetName,buildDetailLookupArray)
}
