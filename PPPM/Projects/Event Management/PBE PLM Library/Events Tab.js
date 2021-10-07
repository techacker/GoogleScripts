/**
 * Adds a row to the main "Events" sheet by taking relevant values from "Add New Event" sheet and initiates the tracker creation process
 * @author Anurag Bansal <anurag.bansal@fcagroup.com>
 * @param url - URL to the target Workload File. It uses default if nothing is passed.
 * @param neweventSheetName - Name of the sheet where the new event data is being entered, default = "Add New Event"
 * @param maineventSheet - Name of the sheet where the list of events/programs are in this workload file, default = "Events"
 * @param defaultHeader - Header name in the Event sheet that will never change - Something to find the header row, default = "Event Title".
 * @param trackersFolder - Name where all "New" trackers are to be stored in Google Drive, default = "PPPM Shared Part Trackers"
 * @param templateFile - "ID" of the template file in Google Drive, default is PPPM Template
 * @param mandatoryFields - An array of mandatory fields in the New Event Sheet, default = Mandatory fields from PPPM Workload file.
 * @return {void}
 */
function createNewTracker(url=workloadfileURL, neweventSheetName=addNeweventSheetName, maineventSheet=mainEventsSheetName,defaultHeader=keyWord, trackersFolder=trackersFolderName, templateFile=templateFileID, mandatoryFields=newEventMandatoryFields) {

  let newEventSheet = SpreadsheetApp.openByUrl(url).getSheetByName(neweventSheetName);
  let lr = newEventSheet.getLastRow();
  let newEventData = newEventSheet.getRange(2,1,lr-1,2).getDisplayValues();
  const eventSheet = SpreadsheetApp.openByUrl(url).getSheetByName(maineventSheet);
  const eventSheetLR = eventSheet.getLastRow();
  const eventSheetLC = eventSheet.getLastColumn();
  const headerFinder = eventSheet.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const eventSheetHeaders = eventSheet.getRange(headerRow,1,eventSheetLR,eventSheetLC).getDisplayValues()[0];
  const dataAdded = new Date().toLocaleDateString();
  
  let NewEventObj = {}
  
  // For every data point in New Event Sheet, it will create the object with relevant values
  newEventData.forEach(item => {
    NewEventObj[item[0]] = item[1]
  })

  // Check if mandatory fields exist in the New Event Sheet, if they don't then exit
  mandatoryFields.forEach(item => {
    if (Object.keys(NewEventObj).indexOf(item) === -1) {
      Browser.msgBox("Alert",item + " is not found in" + neweventSheetName + ". Please check the spelling and make sure it matches the " + maineventSheet, Browser.Buttons.OK); 
      return false
    }
  })
  
  // Check if manadatory items are provided in the New Event Sheet
  mandatoryFields.forEach(item => {
      if (NewEventObj[item] === "") {
        Browser.msgBox("Alert",item + " is not provided. But it is mandatory.", Browser.Buttons.OK); 
        //Logger.log(item + " is not provided. But it is mandatory.")
        return false
      } 
  })

  let EventSheetObj = lookupColumns(url,maineventSheet,defaultHeader,eventTabLookupArray);

  // Add default values to the NewEventObj and create event title
  NewEventObj["Event Status"] = "In-Process";
  NewEventObj["Date Added"] = dataAdded;
  NewEventObj["Event Title"] = NewEventObj["Model Year"] + " " + NewEventObj["Vehicle Family"].toUpperCase() + " " + NewEventObj["Event Name"]
  NewEventObj["Tracker URL"] = "";
  
  // Define each header with its corresponding values
  eventSheetHeaders.forEach(heading => {
    if (NewEventObj[heading] === null) {
      EventSheetObj[heading] = ""
    } else {
      EventSheetObj[heading] = NewEventObj[heading]
    }
  })
  
  // Get the previous event details
  const lastEventRange = eventSheet.getRange(eventSheetLR,1,1,eventSheetLC);
  const lastEventFormulas = lastEventRange.getFormulas()[0];
  const newEventRange = eventSheet.getRange(eventSheetLR+1,1,1,eventSheetLC);

  // Get values from newEventArray and formulas from lastEventFormulas Array and put them in new row.
  eventSheetHeaders.forEach((heading,ind) => {
    if (EventSheetObj[heading] !== undefined) {
      eventSheet.getRange(eventSheetLR+1,ind+1,1,1).setValue(EventSheetObj[heading])
    }
  })

  // Copy the formulas from previous row
  lastEventFormulas.filter((item,ind)=> {
    if (item !== "") {
      eventSheet.getRange(eventSheetLR,ind,1,1).copyTo(eventSheet.getRange(eventSheetLR+1,ind,1,1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false)
    }
  }) 
  
  // Copy the format from previous row
  lastEventRange.copyTo(newEventRange,SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false) 

  // Create the URL
  getNewTrackerURL(url, maineventSheet, defaultHeader, trackersFolder, templateFile);
  
  newEventSheet.getRange(2,2,lr-1,1).clearContent();
  
  Browser.msgBox("Success", "A new tracker has been created.", Browser.Buttons.OK);
  
}

/**
 * Creates a new tracker using the template and data from "Events" sheet
 * @author Anurag Bansal <anurag.bansal@fcagroup.com>
 * @param url - URL to the target Workload File. It uses default it nothing is passed.
 * @param maineventSheet - Name of the sheet where the list of events/programs are in this workload file, default = "Events"
 * @param keyWord - A single column heading that will probably never change, default = "Event Title"
 * @param trackersFolder - Name where all "New" trackers are to be stored in Google Drive, default = "PPPM Shared Part Trackers"
 * @param templateFileID - "ID" of the template file in Google Drive, default is PPPM Template
 * @param archiveFolderName - Name where the "Archived" trackers are to be stored in Google Drive, default = "Archive - PPPM Trackers"
 * @return true
 */
function getNewTrackerURL(url=workloadfileURL, maineventSheet=mainEventsSheetName, lookupWord=keyWord, trackersFolder=trackersFolderName, tempFileID=templateFileID, lookupArray=eventTabLookupArray) {
  
  let eventSheet = SpreadsheetApp.openByUrl(url).getSheetByName(maineventSheet);
  let lr = eventSheet.getLastRow();
  let lc = eventSheet.getLastColumn();
  let textFinder = eventSheet.getDataRange().createTextFinder(lookupWord).findNext();
  let headerRow = textFinder.getRow();  
  let ColIndices = lookupColumns(url,maineventSheet,lookupWord,lookupArray);;
  let dataRange = eventSheet.getRange(headerRow+1, 1, lr-headerRow, lc).getDisplayValues();
  const eventSheetHeaders = eventSheet.getRange(headerRow,1,1,lc).getDisplayValues()[0];
  let EventListObj = {}
  eventSheetHeaders.forEach(item => {
    EventListObj[item] = eventSheetHeaders.indexOf(item)
  })

  //Logger.log(ColIndices)
  //{Event Title=5.0, Event Type=4.0, Ship-to Code=12.0, Location of Event=11.0, Requestor=9.0, Ship-to Address=13.0, Attention-to=14.0, Event  Status=7.0, Program Manager=8.0, Initial Part #s=39.0, WBS Code=10.0, VF=1.0, Tracker URL=40.0, Earliest MRD=6.0}
    
  for (let i=0; i<dataRange.length; i++) { 
    let vf = dataRange[i][ColIndices["Vehicle Family"]];
    let eventTitle = dataRange[i][ColIndices["Event Title"]];
    //let mRD = dataRange[i][ColIndices["Earliest MRD"]];
    let eventStatus = dataRange[i][ColIndices["Event Status"]];
    let requestor = dataRange[i][ColIndices["Requestor"]]
    let wBSCode = dataRange[i][ColIndices["WBS Code"]]
    let location = dataRange[i][ColIndices["Location of Event"]];
    let shipto = dataRange[i][ColIndices["Ship-to Code"]];
    let shipAddress = dataRange[i][ColIndices["Ship-to Address"]];
    let attention = dataRange[i][ColIndices["Attention-to"]];
    let trackerUrl = dataRange[i][ColIndices["Tracker URL"]];
    
    if (eventStatus === "In-Process" && trackerUrl === "") {
      let link = generateNewTrackerFile_(vf.trim(), eventTitle.trim(),trackersFolder,tempFileID);
      eventSheet.getRange(i+headerRow+1, ColIndices["Tracker URL"]+1, 1, 1).setValue(link);
      //updateNewTracker_(trackerUrl,EventListObj)
      pushEventInfo_(url, maineventSheet, lookupWord, link, eventTitle, requestor, wBSCode, location, shipto, shipAddress, attention,lookupArray);
    }
    
  }

  return true;
  
}

// Generate the tracker file using the given template in a specified location
function generateNewTrackerFile_(vehFam, eventTitle,trackersFolder,templateFileID) {
  
  let folders = DriveApp.getFolders();
  let pPPMFolderName = trackersFolder;  
  
  // Get folder names inside Shared Part Trackers folder
  let pPPMFolders = DriveApp.getFoldersByName(pPPMFolderName);
  let pPPMFolder = pPPMFolders.next();
  let vFFolders = pPPMFolder.getFolders();
  let vFFolderNames = [];
  while (vFFolders.hasNext()) {
    let vFFolder = vFFolders.next();
    let vFFolderName = vFFolder.getName();
    vFFolderNames.push(vFFolderName);
  }
  
  // Create a Vehicle Family Folder inside if it doesn't exist
  if (vFFolderNames.indexOf(vehFam) === -1) {
    pPPMFolder.createFolder(vehFam);
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
  if (vehFamFileNames.indexOf(eventTitle + " Tracker") === -1) {
    let newTracker = DriveApp.getFileById(templateFileID).makeCopy(eventTitle + " Tracker", vehFamFolder);
    let url = newTracker.getUrl();
    let link = parseURL_(url);
    return link;
  }
  // Show an error message that it already exists and capture its url.
  else {
    Browser.msgBox("Alert!", "Tracker for " + eventTitle + " already exists.", Browser.Buttons.OK);
    let url = getTrackerURL_(vehFamFileNames, eventTitle);
    let link = parseURL_(url);
    return link;
  }
}


/**
 * Push event info into the newly created tracker
 * @author Anurag Bansal <anurag.bansal@fcagroup.com>
 * @param url - URL to the target Workload File. It uses default it nothing is passed.
 * @param defaultHeader - Header name in the Event sheet that will never change - Something to find the header row, default = "Event Title".
 * @param maineventSheet - Name of the sheet where the list of events/programs are in this workload file, default = "Events"
 * @param trackersFolder - Name where all "New" trackers are to be stored in Google Drive, default = "PPPM Shared Part Trackers"
 * @param templateFileID - "ID" of the template file in Google Drive, default is PPPM Template
 * @return {void}
 */
function pushEventInfo_(url=workloadfileURL, maineventSheet=mainEventsSheetName, lookupWord=keyWord, link, eventTitle, requestor, wBSCode, location, shipto, shipAddress, attention,lookupArray=eventTabLookupArray) {
  
  //Event Sheet Info
  let eventSheet = SpreadsheetApp.openByUrl(url).getSheetByName(maineventSheet);
  let lr = eventSheet.getLastRow();
  let lc = eventSheet.getLastColumn();
  let textFinder = eventSheet.getDataRange().createTextFinder(lookupWord).findNext();
  let headerRow = textFinder.getRow();  
  let ColIndices = lookupColumns(url,maineventSheet,lookupWord,lookupArray);
  let dataRange = eventSheet.getRange(headerRow+1, 1, lr-headerRow, lc).getDisplayValues();
  
  if (link) {
    var summarySheet = SpreadsheetApp.openByUrl(link).getSheetByName("Summary");
    summarySheet.getRange(1, 3).setValue(eventTitle);
    summarySheet.getRange(2, 3).setValue(requestor);
    summarySheet.getRange(3, 3).setValue(wBSCode);
    summarySheet.getRange(4, 3).setValue(location);
    summarySheet.getRange(5, 3).setValue(shipto);
    summarySheet.getRange(6, 3).setValue(shipAddress);
    summarySheet.getRange(7, 3).setValue(attention);
    return true;
  }
  
  for (var i=0; i<dataRange.length; i++) {
    let vf = dataRange[i][ColIndices["Vehicle Family"]];
    let eventTitle = dataRange[i][ColIndices["Event Title"]];
    let requestor = dataRange[i][ColIndices["Requestor"]]
    let wBSCode = dataRange[i][ColIndices["WBS Code"]]
    let location = dataRange[i][ColIndices["Location of Event"]];
    let shipto = dataRange[i][ColIndices["Ship-to Code"]];
    let shipAddress = dataRange[i][ColIndices["Ship-to Address"]];
    let attention = dataRange[i][ColIndices["Attention-to"]];
    let trackerUrl = dataRange[i][ColIndices["Tracker URL"]];
    
    if (eventStatus === "In-Process" && trackerUrl !== "") {
      var summarySheet = SpreadsheetApp.openByUrl(url).getSheetByName("Summary");
      summarySheet.getRange(1, 3).setValue(eventTitle);
      summarySheet.getRange(2, 3).setValue(requestor);
      summarySheet.getRange(3, 3).setValue(wBSCode);
      summarySheet.getRange(4, 3).setValue(location);
      summarySheet.getRange(5, 3).setValue(shipto);
      summarySheet.getRange(6, 3).setValue(shipAddress);
      summarySheet.getRange(7, 3).setValue(attention);
    }
  }
  
}


/**
 * Archive trackers to the given Archive Folder
 * @author Anurag Bansal <anurag.bansal@fcagroup.com>
 * @param url - URL to the target Workload File. It uses default it nothing is passed.
 * @param maineventSheet - Name of the sheet where the list of events/programs are in this workload file, default = "Events"
 * @param defaultHeader - Header name in the Event sheet that will never change - Something to find the header row, default = "Event Title".
 * @param trackersFolder - Name where all "New" trackers are to be stored in Google Drive, default = "PPPM Shared Part Trackers"
 * @param archiveFolder - Name where all "Archived" trackers are to be stored in Google Drive, default = "Archive - PPPM Trackers"
 * @return {void}
 */
function archiveTrackers(url, maineventSheet, defaultHeader, archiveFolder=archiveFolderName) {
  
  var eventSheet = SpreadsheetApp.openByUrl(url).getSheetByName(maineventSheet);
  var lr = eventSheet.getLastRow();
  var lc = eventSheet.getLastColumn();
  var headerRow = getHeaderRow_(eventSheet, defaultHeader);
  var colIndices = getColIndex_(eventSheet, headerRow);
  
  // Logger.log(colIndices);
  // [2.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 30.0]
  // VFInd, EventTitleInd, MRDColInd, EventStatusColInd, ProgMgrColInd, RequestorColInd, 
  // WBSColInd, LocColInd, ShiptoColInd, ShipAddColInd, AttnColInd, urlColInd
  
  //var eventsHeader = eventSheet.getRange(headerRow, 1, 1, lc).getValues()[0];
  var dataRange = eventSheet.getRange(headerRow+1, 1, lr-headerRow, lc).getDisplayValues();
  var archiveFolderName = archiveFolder;
  
  var eventStatusColInd = colIndices[4];
  var urlColInd = colIndices[13];
  
  for (var i=0; i<dataRange.length; i++) { 
    var eventStatus = dataRange[i][eventStatusColInd];
    var url = dataRange[i][urlColInd];
    
    if (eventStatus === "Complete" && url !== "") {
      var id = url.split('/')[5];
      var archiveFolder = DriveApp.getFoldersByName(archiveFolderName);
      var trackerFile = DriveApp.getFileById(id);
      trackerFile.moveTo(archiveFolder.next());
      eventSheet.getRange(i+headerRow+1, eventStatusColInd+1, 1, 1).setValue("Archived");
    }
  }
  
  return true;
}


function editEventInfo(url=workloadfileURL, editForm=editEventSheetName, maineventSheet=mainEventsSheetName,defaultHeader=keyWord) {

  let ui = SpreadsheetApp.getUi();
  let editEventForm = SpreadsheetApp.openByUrl(url).getSheetByName(editForm);
  let lr = editEventForm.getLastRow();
  let editEventData = editEventForm.getRange(4,1,lr-3,3).getDisplayValues();
  let eventTitle = editEventForm.getRange(2,2).getValue();
  //let buildLead = getUserName()

  const ssEventSheet = SpreadsheetApp.openByUrl(url).getSheetByName(maineventSheet);
  const ssEventSheetLR = ssEventSheet.getLastRow();
  const ssEventSheetLC = ssEventSheet.getLastColumn();
  const headerFinder = ssEventSheet.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const ssEventSheetHeaders = ssEventSheet.getRange(headerRow,1,1,ssEventSheetLC).getDisplayValues()[0];
  const dataModified = new Date().toLocaleDateString();
  
  let EditEventObj = {}
  let EventListObj = {}
  let OrigEventObj = {}
  //let ChangeLogObj = {"Date":dataModified,"What Changed":"", "From":"",	"To":"",	"Who": buildLead}

  // For every data point in edit event Form, it will create the object with relevant values
  editEventData.forEach(item => {
    OrigEventObj[item[0]] = item[1]
    EditEventObj[item[0]] = item[2]
  })
  
  // For every key in EditEventObj if that value is not blank, use the existing value
  Object.keys(EditEventObj).forEach(key => {
    if (EditEventObj[key] === "" && EditEventObj[key] !== OrigEventObj[key]) {
      EditEventObj[key] = OrigEventObj[key]
    } 
  })

  ssEventSheetHeaders.forEach(item => {
    EventListObj[item] = ssEventSheetHeaders.indexOf(item)
  })

  // Get event URL from the event list sheet
  const eventsArray = ssEventSheet.getRange(headerRow+1,1,ssEventSheetLR-headerRow,ssEventSheetLC).getValues();
  
  let eventRow = 0
  let eventTrackerURL = ""
  eventsArray.filter((item, row) => {
    if (item[EventListObj["Event Title"]] === eventTitle) {
      eventRow = row + headerRow + 1
      eventTrackerURL = item[EventListObj["Tracker URL"]]
    }
  })
  
  // Add default values to the NewBuildObj and create Event Title
  EditEventObj["Last Modified"] = dataModified;
  EditEventObj["Event Status"] = "Active";
  EditEventObj["Event Title"] = EditEventObj["Model Year"]+" "+EditEventObj["Vehicle Family"].toUpperCase()+" "+EditEventObj["Event Name"]
  EditEventObj["Tracker URL"] = eventTrackerURL

  let eventTitles = eventsArray.map(item => item[EventListObj["Event Title"]]) 

  // To verify if the modified event name doesn't already exist, add the new name to the eventTitles array
  if (EditEventObj["Event Title"] !== eventTitle) {
    eventTitles.push(EditEventObj["Event Title"])
  }
  
  // If the name already exist, the filter array length will be more than one.
  if (eventTitles.filter(item => item === EditEventObj["Event Title"]).length === 1) {
    // Create the proper entries to be used for updating the event list
    ssEventSheetHeaders.forEach(heading => {
      if (EditEventObj[heading] === null) {
        EventListObj[heading] = ""
      } else {
        EventListObj[heading] = EditEventObj[heading]
      }
    })

    // Update event details in PBE event List
    Object.keys(EventListObj).forEach((key, ind) => {
      if (EventListObj[key] !== undefined) {
        ssEventSheet.getRange(eventRow,ind+1,1,1).setValue(EventListObj[key])
      }
    })

    // Update data in the respective tracker file
    updateNewTracker_(eventTrackerURL, EventListObj)
    /*
    // Collect information for Change Log and call the function to update the log
    Object.keys(OrigEventObj).filter(key => {
      if (EditEventObj[key] !== "" && EditEventObj[key] !== OrigEventObj[key]) {
        ChangeLogObj["What Changed"] = key
        ChangeLogObj["From"] = OrigEventObj[key]
        ChangeLogObj["To"] = EditEventObj[key]
        updateChangeLog_(eventTrackerURL,ChangeLogObj)
      }
    })
    */
    SpreadsheetApp.openByUrl(eventTrackerURL).rename(EditEventObj["Event Title"] + " Tracker")
    
    // Reset edit form
    editEventForm.getRange(2,2).clearContent()
    editEventForm.getRange(4,3,lr-3,1).clearContent()
    // Show success message
    ui.alert("Success...","The event was updated successfully.",ui.ButtonSet.OK)
  } else {
    // Show error message
    ui.alert("Error...","The event with similar name already exist. Please check and try again.",ui.ButtonSet.OK)
  }  
  
}

function updateChangeLog_(eventTrackerURL,ChangeLogObj) {

  const changeLogSheet = SpreadsheetApp.openByUrl(eventTrackerURL).getSheetByName("Change Log");
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


function updateNewTracker_(newTrackerURL, EventListObj) {

  const summarySheet = SpreadsheetApp.openByUrl(newTrackerURL).getSheetByName("Summary");
  const headerFinder = summarySheet.createTextFinder("Tab").findNext();
  const headerRow = headerFinder.getRow();
  const infoArray = summarySheet.getRange(1,2,headerRow-1,2).getValues();

  let SummaryObj = {}

  // Create SummaryObj by looking at Summary Sheet Array
  infoArray.forEach((item, ind) => {
    SummaryObj[item[0]] = ind
  })

  // For every key in Summary Obj, set its value from passed on BuildListObj
  Object.keys(SummaryObj).forEach(key => {
    if (key in EventListObj) {
      SummaryObj[key] = EventListObj[key]
    }
    else {
      SummaryObj[key] = ""
    }
  })
  
  // Write these values into Summary Sheet Information Section
  Object.keys(SummaryObj).forEach((key) => {
    infoArray.forEach((item, ind) => {
      if (item[0] === key) {
        summarySheet.getRange(ind+1,3,1,1).setValue(SummaryObj[key])
      }
    })
  })
  
}
