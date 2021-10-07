/**
 * Adds a row to the main "Events" sheet by taking relevant values from "Add New Event" sheet and initiates the tracker creation process
 * @author Anurag Bansal <anurag.bansal@fcagroup.com>
 * @param url - URL to the target Workload File. It uses default if nothing is passed.
 * @param addNewBuild - Name of the sheet where the new event data is being entered, default = "Add New Event"
 * @param pbeBuildList - Name of the sheet where the list of events/programs are in this workload file, default = "Events"
 * @param defaultHeader - Header name in the Event sheet that will never change - Something to find the header row, default = "Event Title".
 * @param trackersFolder - Name where all "New" trackers are to be stored in Google Drive, default = "PPPM Shared Part Trackers"
 * @param templateFileID - "ID" of the template file in Google Drive
 * @param mandatory - An array of mandatory fields in the New Event Sheet, default = Mandatory fields from PPPM Workload file.
 * @return {void}
 */
function createBuildTracker(url=buildsumfileURL, addNewBuild=addNewBuildSheetName, pbeBuildList=buildsListSheetName,defaultHeader=keyWord, trackersFolder=trackersFolderName, templateFileID=templateID, mandatory=mandatoryFields,parentFileID=parentBOMFileID) {

  let ui = SpreadsheetApp.getUi();
  let newBuildForm = SpreadsheetApp.openByUrl(url).getSheetByName(addNewBuild);
  let lr = newBuildForm.getLastRow();
  let newBuildData = newBuildForm.getRange(2,1,lr-1,2).getDisplayValues();
  let buildLead = getUserName()
  
  const ssbuildList = SpreadsheetApp.openByUrl(url).getSheetByName(pbeBuildList);
  const ssbuildListLR = ssbuildList.getLastRow();
  const ssbuildListLC = ssbuildList.getLastColumn();
  const headerFinder = ssbuildList.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const ssbuildListHeaders = ssbuildList.getRange(headerRow,1,1,ssbuildListLC).getDisplayValues()[0];
  const dataAdded = new Date().toLocaleDateString();
  
  let NewBuildObj = {}
  let BuildListObj = {}
  
  // For every data point in New Build Form, it will create the object with relevant values
  newBuildData.forEach(item => {
    NewBuildObj[item[0]] = item[1]
  })

  ssbuildListHeaders.forEach(item => {
    BuildListObj[item] = ssbuildListHeaders.indexOf(item)
  })

  // Check if mandatory fields exist in the New Build Form, if they don't then exit
  mandatory.forEach(item => {
    if (Object.keys(NewBuildObj).indexOf(item) === -1) {
      ui.alert("Error...", item + " is not found in " + addNewBuild + ". Please check the spelling and make sure it matches the headers in " + pbeBuildList, ui.ButtonSet.OK) 
      return false;
    }
  })

  // Check if manadatory items are provided in the New Build Form
  mandatory.forEach(item => {
    if (NewBuildObj[item] === "") {
      ui.alert("Error...", item + " is not provided. It is required to create the build. Please try again.",ui.ButtonSet.OK)
      return false;
    } 
  })
  
  // Add default values to the NewBuildObj and create Build Title
  NewBuildObj["Date Created"] = dataAdded;
  NewBuildObj["Last Modified"] = dataAdded;
  NewBuildObj["Build Lead"] = buildLead;
  NewBuildObj["Build Status"] = "Active";
  NewBuildObj["Build Title"] = NewBuildObj["Model Year"] + " " + NewBuildObj["Vehicle Family"].toUpperCase() + " " + NewBuildObj["Build Type"] + " " + NewBuildObj["Build Phase"]
  
  // Generate URL for the new tracker
  let newTrackerURL = getNewTrackerURL_(NewBuildObj["Vehicle Family"], NewBuildObj["Build Title"],trackersFolder,templateFileID);
  
  NewBuildObj["Build Tracker URL"] = newTrackerURL

  ssbuildListHeaders.forEach(heading => {
    if (NewBuildObj[heading] === null) {
      BuildListObj[heading] = ""
    } else {
      BuildListObj[heading] = NewBuildObj[heading]
    }
  })
  
  // Get the previous event details
  const lastBuildRange = ssbuildList.getRange(ssbuildListLR,1,1,ssbuildListLC);
  const lastBuildFormulas = lastBuildRange.getFormulas()[0];
  const newBuildRange = ssbuildList.getRange(ssbuildListLR+1,1,1,ssbuildListLC);

  // Get values from newEventArray and formulas from lastBuildFormulas Array and put them in new row.
  Object.keys(BuildListObj).forEach((key, ind) => {
    if (lastBuildFormulas[ind] !== "") {
      ssbuildList.getRange(ssbuildListLR,ind+1,1,1).copyTo(ssbuildList.getRange(ssbuildListLR+1,ind+1,1,1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false)
    } else if (BuildListObj[key] !== undefined) {
      ssbuildList.getRange(ssbuildListLR+1,ind+1,1,1).setValue(BuildListObj[key])
    } 
  })

  // Compile Draft Email Array with required information
  let newBuildEmailArray = [BuildListObj["Build Title"],BuildListObj["TCF MRD"],BuildListObj["TCF Build Start"],BuildListObj["Vehicle Quantity"],   BuildListObj["Build Location"],BuildListObj["Tree FC#"],BuildListObj["WBS Code"],BuildListObj["Build Tracker URL"]]

  // Create Draft Email
  draftHTMLEmails_(url,pbeBuildList,defaultHeader,newBuildEmailArray)

  // Copy the format from previous row
  lastBuildRange.copyTo(newBuildRange,SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false) 
  
  if (newTrackerURL !== "") {
    updateNewTracker_(newTrackerURL, BuildListObj)
  }
  
  // Clear form in New Event Sheet
  newBuildForm.getRange(2,2,lr-1,1).clearContent();

  // Update BOM File
  updateBOMFile_(BuildListObj, parentFileID)
  //updateBOMFileWithValidation_(BuildListObj,parentFileID)
  
  // Create KPI File
  //updateKPIFile_(NewBuildObj)
  
  ui.alert("Success...","1. A new Build Tracker file was created.\n2. An announcement email was saved in your draft folder.\n3. The BOM file was updated.",ui.ButtonSet.OK)
  
}

function updateNewTracker_(newTrackerURL, BuildListObj) {

  const summarySheet = SpreadsheetApp.openByUrl(newTrackerURL).getSheetByName("Build Summary");
  const headerFinder = summarySheet.createTextFinder("BICEEP").findNext();
  const headerRow = headerFinder.getRow();
  const infoArray = summarySheet.getRange(1,1,headerRow-1,2).getValues();

  let SummaryObj = {}

  // Create SummaryObj by looking at Summary Sheet Array
  infoArray.forEach((item, ind) => {
    SummaryObj[item[0]] = ind
  })

  // For every key in Summary Obj, set its value from passed on BuildListObj
  Object.keys(SummaryObj).forEach(key => {
    if (key in BuildListObj) {
      SummaryObj[key] = BuildListObj[key]
    }
    else {
      SummaryObj[key] = ""
    }
  })
  
  // Write these values into Summary Sheet Information Section
  Object.keys(SummaryObj).forEach((key) => {
    infoArray.forEach((item, ind) => {
      if (item[0] === key) {
        summarySheet.getRange(ind+1,2,1,1).setValue(SummaryObj[key])
      }
    })
  })
}

// Generate the tracker file using the given template in a specified location
function getNewTrackerURL_(vehFam, buildTitle,trackersFolder,templateFileID) { 
  
  // Get folder names inside Shared Part Trackers folder
  let folders = DriveApp.getFoldersByName(trackersFolder);
  let folder = folders.next();
  let vFFolders = folder.getFolders();
  let vFFolderNames = [];
  while (vFFolders.hasNext()) {
    let vFFolder = vFFolders.next();
    let vFFolderName = vFFolder.getName();
    vFFolderNames.push(vFFolderName);
  }
  
  // Create a Vehicle Family Folder inside if it doesn't exist
  if (vFFolderNames.indexOf(vehFam) === -1) {
    folder.createFolder(vehFam);
  }
  
  // Get file names inside Veh Family folder
  let vehFamFolders = DriveApp.getFoldersByName(vehFam);
  let vehFamFolder = vehFamFolders.next();
  let vehFamFiles = vehFamFolder.getFiles();
  let vehFamFileNames = [];
  while (vehFamFiles.hasNext()) {
    let vehFamFile = vehFamFiles.next();
    let vehFamFileName = vehFamFile.getName();
    vehFamFileNames.push(vehFamFileName);
  }
  
  // Create a new tracker inside if it doesn't exist
  if (vehFamFileNames.indexOf(buildTitle + " Tracker") === -1) {
    let newTracker = DriveApp.getFileById(templateFileID).makeCopy(buildTitle + " Tracker", vehFamFolder);
    let url = newTracker.getUrl();
    let link = parseURL_(url);
    return link;
  }
  // Show an error message that it already exists and capture its url.
  else {
    Browser.msgBox("Alert!", "Tracker for " + buildTitle + " already exists.", Browser.Buttons.OK);
    let url = getTrackerURL_(vehFamFileNames, buildTitle);
    let link = parseURL_(url);
    return link;
  }
}


/**
 * Archive trackers to the given Archive Folder
 * @author Anurag Bansal <anurag.bansal@fcagroup.com>
 * @param url - URL to the target Workload File. It uses default it nothing is passed.
 * @param pbeBuildList - Name of the sheet where the list of events/programs are in this workload file, default = "Events"
 * @param defaultHeader - Header name in the Event sheet that will never change - Something to find the header row, default = "Event Title".
 * @param trackersFolder - Name where all "New" trackers are to be stored in Google Drive, default = "PPPM Shared Part Trackers"
 * @param archiveFolder - Name where all "Archived" trackers are to be stored in Google Drive, default = "Archive - PPPM Trackers"
 * @return {void}
 */
function archiveTrackers(url=buildsumfileURL, pbeBuildList=buildsListSheetName,defaultHeader=keyWord, archiveFolder=archiveFolderName) {
  
  var ssbuildList = SpreadsheetApp.openByUrl(url).getSheetByName(pbeBuildList);
  const ssbuildListLR = ssbuildList.getLastRow();
  const ssbuildListLC = ssbuildList.getLastColumn();
  const headerFinder = ssbuildList.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const ssbuildListHeaders = ssbuildList.getRange(headerRow,1,1,ssbuildListLC).getDisplayValues()[0];
  const buildsArray = ssbuildList.getRange(headerRow+1,1,ssbuildListLR-headerRow,ssbuildListLC).getValues();
  var archiveFolderName = archiveFolder;

  let BuildListObj = {}

  ssbuildListHeaders.forEach(item => {
    BuildListObj[item] = ssbuildListHeaders.indexOf(item)
  })

  // *********  Work on Archive Trackers function *******
  
  // Get values from newEventArray and formulas from lastBuildFormulas Array and put them in new row.
  Object.keys(BuildListObj).forEach((key, ind) => {
    if (buildsArray[ind] !== "") {
      Logger.log(BuildListObj["Build Tracker URL"] + " " + key + " " + ind)
      //ssbuildList.getRange(ssbuildListLR,ind+1,1,1).copyTo(ssbuildList.getRange(ssbuildListLR+1,ind+1,1,1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false)
    } else if (BuildListObj[key] !== undefined) {
      //ssbuildList.getRange(ssbuildListLR+1,ind+1,1,1).setValue(BuildListObj[key])
    } 
  })

}


function editBuild(url=buildsumfileURL, editForm=editBuildSheetName, pbeBuildList=buildsListSheetName,defaultHeader=keyWord,parentFileID=parentBOMFileID) {

  let ui = SpreadsheetApp.getUi();
  let editBuildForm = SpreadsheetApp.openByUrl(url).getSheetByName(editForm);
  let lr = editBuildForm.getLastRow();
  let editBuildData = editBuildForm.getRange(4,1,lr-3,3).getDisplayValues();
  let buildTitle = editBuildForm.getRange(2,2).getValue();
  let userName = getUserName()

  const ssbuildList = SpreadsheetApp.openByUrl(url).getSheetByName(pbeBuildList);
  const ssbuildListLR = ssbuildList.getLastRow();
  const ssbuildListLC = ssbuildList.getLastColumn();
  const headerFinder = ssbuildList.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const ssbuildListHeaders = ssbuildList.getRange(headerRow,1,1,ssbuildListLC).getDisplayValues()[0];
  const dataModified = new Date().toLocaleDateString();
  
  let EditBuildObj = {}
  let BuildListObj = {}
  let OrigBuildObj = {}
  let ChangeLogObj = {"Date":dataModified,"What Changed":"", "From":"",	"To":"",	"Who": userName}

  // For every data point in edit Build Form, it will create the object with relevant values
  editBuildData.forEach(item => {
    OrigBuildObj[item[0]] = item[1]
    EditBuildObj[item[0]] = item[2]
  })
  
  // For every key in EditBuildObj if that value is not blank, use the existing value
  Object.keys(EditBuildObj).forEach(key => {
    if (EditBuildObj[key] === "" && EditBuildObj[key] !== OrigBuildObj[key]) {
      EditBuildObj[key] = OrigBuildObj[key]
    } 
  })

  ssbuildListHeaders.forEach(item => {
    BuildListObj[item] = ssbuildListHeaders.indexOf(item)
  })

  // Get event URL from the build list sheet
  const buildsArray = ssbuildList.getRange(headerRow+1,1,ssbuildListLR-headerRow,ssbuildListLC).getValues();
  
  let buildRow = 0
  let buildTrackerURL = ""
  buildsArray.filter((item, row) => {
    if (item[BuildListObj["Build Title"]] === buildTitle) {
      buildRow = row + headerRow + 1
      buildTrackerURL = item[BuildListObj["Build Tracker URL"]]
    }
  })
  
  // Add default values to the NewBuildObj and create Build Title
  EditBuildObj["Last Modified"] = dataModified;
  //EditBuildObj["Build Lead"] = BuildListObj['Build Lead'];
  EditBuildObj["Build Status"] = "Active";
  EditBuildObj["Build Title"] = EditBuildObj["Model Year"] + " " + EditBuildObj["Vehicle Family"].toUpperCase() + " " + EditBuildObj["Build Type"] + " " + EditBuildObj["Build Phase"]
  EditBuildObj["Build Tracker URL"] = buildTrackerURL

  let buildTitles = buildsArray.map(item => item[BuildListObj["Build Title"]]) 

  // To verify if the modified build name doesn't already exist, add the new name to the buildTitles array
  if (EditBuildObj["Build Title"] !== buildTitle) {
    buildTitles.push(EditBuildObj["Build Title"])
  }
  
  // If the name already exist, the filter array length will be more than one.
  if (buildTitles.filter(item => item === EditBuildObj["Build Title"]).length === 1) {
    // Create the proper entries to be used for updating the build list
    ssbuildListHeaders.forEach(heading => {
      if (EditBuildObj[heading] === null) {
        BuildListObj[heading] = ""
      } else {
        BuildListObj[heading] = EditBuildObj[heading]
      }
    })

    // Update build details in PBE Build List
    Object.keys(BuildListObj).forEach((key, ind) => {
      if (BuildListObj[key] !== undefined) {
        ssbuildList.getRange(buildRow,ind+1,1,1).setValue(BuildListObj[key])
      }
    })

    // Update data in the respective tracker file
    updateNewTracker_(buildTrackerURL, BuildListObj)

    let changeBulletinArray = []

    // Collect information for Change Log and call the function to update the log
    Object.keys(OrigBuildObj).filter(key => {
      if (EditBuildObj[key] !== "" && EditBuildObj[key] !== OrigBuildObj[key]) {
        ChangeLogObj["What Changed"] = key
        ChangeLogObj["From"] = OrigBuildObj[key]
        ChangeLogObj["To"] = EditBuildObj[key]
        updateChangeLog_(buildTrackerURL,ChangeLogObj)
        changeBulletinArray.push([ChangeLogObj["What Changed"],ChangeLogObj["From"],ChangeLogObj["To"]])
      }
    })

    //Logger.log(changeArray)
    draftChangeBulletinEmails_(url, buildTitle, buildTrackerURL, changeBulletinArray)

    SpreadsheetApp.openByUrl(buildTrackerURL).rename(EditBuildObj["Build Title"] + " Tracker")

    // Update BOM file
    updateBOMFile_(EditBuildObj, parentFileID)
    
    // Reset edit form
    editBuildForm.getRange(2,2).clearContent()
    editBuildForm.getRange(4,3,lr-3,1).clearContent()
    // Show success message
    ui.alert("Success...","1. The build was updated successfully.\n2. A change announcement email was saved in your Gmail Draft folder.\n3. The BOM file was updated with new information.",ui.ButtonSet.OK)
  } else {
    // Show error message
    ui.alert("Error...","The build with similar name already exist. Please check and try again.",ui.ButtonSet.OK)
  }  
  
}

function updateChangeLog_(buildTrackerURL,ChangeLogObj) {

  const changeLogSheet = SpreadsheetApp.openByUrl(buildTrackerURL).getSheetByName("Change Log");
  const lr = changeLogSheet.getLastRow()
  const lc = changeLogSheet.getLastColumn()
  const headerFinder = changeLogSheet.createTextFinder("What Changed").findNext();
  const headerRow = headerFinder.getRow();
  const headerArray = changeLogSheet.getRange(headerRow,1,1,lc).getValues()[0];

  let ChangeObj = {}
  // Create ChangeObj from Change Log
  headerArray.forEach((item, ind) => {
    ChangeObj[item] = ind
  })
  
  // Write values in the respective field
  Object.keys(ChangeObj).forEach((key,ind) => {
    ChangeObj[key] = ChangeLogObj[key]
    changeLogSheet.getRange(lr+1,ind+1,1,1).setValue(ChangeObj[key])
  })
  
}
