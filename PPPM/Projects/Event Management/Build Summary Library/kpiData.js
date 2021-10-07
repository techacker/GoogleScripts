function getKPIStatus(wlurl=buildsumfileURL,pbeBuildList=buildsListSheetName,defaultHeader=keyWord,lookupArray=buildDetailLookupArray,kpiDataSheet=kpiSheetName) {

  // Event Sheet
  const eventSheet = SpreadsheetApp.openByUrl(wlurl).getSheetByName(pbeBuildList);
  const eventSheetLR = eventSheet.getLastRow();
  const eventSheetLC = eventSheet.getLastColumn();
  const headerFinder = eventSheet.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const eventSheetHeaders = eventSheet.getRange(headerRow,1,eventSheetLR,eventSheetLC).getDisplayValues()[0];
  const eventSheetArray = eventSheet.getRange(headerRow+1,1,eventSheetLR-headerRow,eventSheetLC).getDisplayValues();  
  const todaysDate = new Date()
  
  // Headers in Event Sheet
  let BuildListObj = lookupColumns(wlurl,pbeBuildList,defaultHeader,lookupArray);
  let SummaryHeaderObj = {}
  let trackerURLArray = eventSheet.getRange(headerRow+1, BuildListObj["Build Tracker URL"]+1, eventSheetLR-headerRow, 1).getValues(); 
  
  // Tracker Sheet
  const kpiSheet = SpreadsheetApp.openByUrl(wlurl).getSheetByName(kpiDataSheet);
  const ttlr = kpiSheet.getLastRow();
  const ttlc = kpiSheet.getLastColumn();
  const ttheaderFinder = kpiSheet.createTextFinder(defaultHeader).findNext();
  let ttheaderRow = ttheaderFinder.getRow();
  const ttHeaders = kpiSheet.getRange(ttheaderRow,1,1,ttlc).getDisplayValues()[0];
  const ttArray = kpiSheet.getRange(ttheaderRow+1,1,ttlr,ttlc).getValues();
  
  let KPIObj = lookupColumns(wlurl,kpiDataSheet,defaultHeader,ttHeaders);
  
  // Compile an array of KPI Data from every tracker
  trackerURLArray.forEach((trackerURL,row) => {
    // Correspond every value from Event Sheet Obj with the values from the array
    if (trackerURL[0] !== "") {
      let eventRow = eventSheetArray[row]
      eventSheetHeaders.forEach((heading, ind) => {
        BuildListObj[heading] = eventRow[ind]
      })
    }

    let eventRowInfo = getEventNum_(wlurl,trackerURL[0],kpiDataSheet)
    let eventRow = eventRowInfo[0]
    let currentEventRowCount = eventRowInfo[1]
    
    // Go through every event that is "In-Process" and add them to KPI Tab
    if (BuildListObj["Event Status"] !== "Archived") {
      let kpiTracker = SpreadsheetApp.openByUrl(trackerURL[0]);
      let trackerSheetsArray = kpiTracker.getSheets()
      let trackerSheetNames = []
      trackerSheetsArray.map(sheet => {
        trackerSheetNames.push(sheet.getSheetName())
      })
      // Check if the tracker has 'Build KPI Data' sheet, if not, ignore that tracker
      if (trackerSheetNames.includes("Build KPI Data")) {
        let kpiTrackerSheet = SpreadsheetApp.openByUrl(trackerURL[0]).getSheetByName(kpiDataSheet);
        let sslc = kpiTrackerSheet.getLastColumn();
        let sslr = kpiTrackerSheet.getRange("A1").getDataRegion().getLastRow();
        let ssheaderFinder = kpiTrackerSheet.createTextFinder("SEQ").findNext();
        let ssheaderRow = ssheaderFinder.getRow(); 
        let ssHeaderArray = kpiTrackerSheet.getRange(ssheaderRow,1,1,sslc).getValues()[0];
        let ssDetailsArray = kpiTrackerSheet.getRange(ssheaderRow+1,1,sslr-ssheaderRow,sslc).getDisplayValues();

        // Define every col location based on Summary Sheet Header Array
        ssHeaderArray.forEach((key,ind) => {
          SummaryHeaderObj[key] = ind
        })

        // Remove items that have no SEQ#
        ssDetailsArray.some((item,ind) => {
          if (item[SummaryHeaderObj["SEQ"]] === "") {
            ssDetailsArray.splice(ind)
          }
        })

        //Check how many rows need to be added
        let reqdRows = ssDetailsArray.length

        // Add/Delete rows based on what is already there in Tracker Data Tab
        if (reqdRows > 0) {
          if (currentEventRowCount === 0) {
            kpiSheet.insertRows(ttheaderRow+1,reqdRows)
            eventRow = ttheaderRow + 1
          } 
          else if (currentEventRowCount < reqdRows) {
            let lessRows = reqdRows - currentEventRowCount
            kpiSheet.insertRows(eventRow,lessRows)
          } 
          else if (currentEventRowCount > reqdRows) {
            let extraRows = currentEventRowCount - reqdRows
            kpiSheet.deleteRows(eventRow,extraRows)
          } 
        }
        
        // For every row, correspond every value from Summary Sheet array to Summary Sheet Object
        ssDetailsArray.forEach((row, arrInd) => {
          //Bind the Headers with the corresponding value from Array
          Object.keys(SummaryHeaderObj).forEach((key,ind) => {
            SummaryHeaderObj[key] = row[ind]
          }) // Close SummaryHeaderObj loop
          
          //Combine values from Summary Sheet and Event Sheet into Tracker Obj
          Object.keys(KPIObj).forEach(key => {
            if (key in SummaryHeaderObj) {
              KPIObj[key] = SummaryHeaderObj[key]
            } else if (key in BuildListObj) {
              KPIObj[key] = BuildListObj[key]
            }
          }) // Close KPIObj loop

          // Write everything into the respective cells
          kpiSheet.getRange(eventRow+arrInd,1,1,ttlc).setValues([Object.values(KPIObj)])
        })  
      }   
    }
    
  })
  
} // Close function

function getEventNum_(wlurl,trackerURL,kpiDataSheet) {

  const kpiSheet = SpreadsheetApp.openByUrl(wlurl).getSheetByName(kpiDataSheet);
  const ttlr = kpiSheet.getLastRow();
  const ttlc = kpiSheet.getLastColumn();
  const ttheaderFinder = kpiSheet.createTextFinder("SEQ").findNext();
  let ttheaderRow = ttheaderFinder.getRow();
  const ttArray = kpiSheet.getRange(ttheaderRow,ttlc,ttlr,1).getValues();
  let eventRowNum = 0
  let eventRowCount = 0
  let urlArray = []
  ttArray.forEach(item => {
    urlArray.push(item[0])
    if (item[0] === trackerURL) {
      eventRowCount++
    } 
  })
  eventRowNum = urlArray.indexOf(trackerURL) + 1
  
  return [eventRowNum,eventRowCount]

}

function editKPIData(url=buildsumfileURL, editKPIForm=kpiAssessmentSheetName, pbeBuildList=buildsListSheetName,defaultHeader=keyWord) {

  let ui = SpreadsheetApp.getUi();
  let editKPIDataForm = SpreadsheetApp.openByUrl(url).getSheetByName(editKPIForm);
  let lr = editKPIDataForm.getLastRow();
  let editBuildData = editKPIDataForm.getRange(4,1,lr-3,4).getDisplayValues();
  let buildTitle = editKPIDataForm.getRange(2,2).getValue();

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

  // Update KPI details in PBE Build List
  Object.keys(BuildListObj).forEach((key, ind) => {
    if (EditBuildObj[key] !== undefined) {
      ssbuildList.getRange(buildRow,ind+1,1,1).setValue(EditBuildObj[key])
    }
  })

  // Reset KPI Assessment form
  editKPIDataForm.getRange(2,2).clearContent()
  editKPIDataForm.getRange(4,3,lr-3,1).clearContent()

  // Show success message
  ui.alert("Success...","The KPI Info was updated successfully.",ui.ButtonSet.OK)

} 
