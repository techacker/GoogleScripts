function refreshEverything(url=buildsumfileURL, pbeBuildList=buildsListSheetName,detailSheet=buildDetailsSheetName,defaultHeader=keyWord,redItemSheet=redItemsSheetName,kpiDataSheet=kpiSheetName,lookupArray=buildDetailLookupArray) {
    getBuildStatus(url,pbeBuildList,detailSheet,defaultHeader)
    getRedItemsStatus(url,pbeBuildList,redItemSheet,defaultHeader)
    getKPIStatus(url,pbeBuildList,defaultHeader,lookupArray,kpiDataSheet)
}

function getBuildStatus(url=buildsumfileURL,pbeBuildList=buildsListSheetName,detailSheet=buildDetailsSheetName,defaultHeader=keyWord) {

  // Build Sheet
  const buildList = SpreadsheetApp.openByUrl(url).getSheetByName(pbeBuildList);
  const buildListLR = buildList.getLastRow();
  const buildListLC = buildList.getLastColumn();
  const headerFinder = buildList.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const buildListHeaders = buildList.getRange(headerRow,1,1,buildListLC).getDisplayValues()[0];

  // Status Sheet
  const statusList = SpreadsheetApp.openByUrl(url).getSheetByName(detailSheet);
  const statusListLC = statusList.getLastColumn();
  const statusheaderFinder = statusList.createTextFinder("Designs-G-#").findNext();
  const statusheaderRow = statusheaderFinder.getRow();
  const statusListHeaders = statusList.getRange(statusheaderRow,1,1,statusListLC).getDisplayValues()[0];

  const dataAdded = new Date().toLocaleDateString();

  // Headers in Build List
  let BuildListObj = {}
  // Headers in Status List
  let StatusListObj = {}
  // Headers in Event Sheet
  let SummaryHeaderObj = {}

  // Create Build List Object with relevant column indexes
  buildListHeaders.forEach(item => {
    BuildListObj[item] = buildListHeaders.indexOf(item)
  })

  // Create Status List Object with relevant column indexes
  statusListHeaders.forEach(item => {
    StatusListObj[item] = statusListHeaders.indexOf(item)
  })

  // Get event URL from the build list sheet
  const buildsArray = buildList.getRange(headerRow+1,1,buildListLR-headerRow,buildListLC).getValues();
  let buildURLs = buildsArray.map(item => item[BuildListObj["Build Tracker URL"]]);
  
  // Compile an array without MASTER and OVERALL EVENT STATUS
  buildURLs.forEach((buildURL,row) => {
    // Correspond every value from Build List Obj with the values from the array
    if (buildURL[0] !== "") {
      let buildRow = buildsArray[row]
      buildListHeaders.forEach((heading, ind) => {
        BuildListObj[heading] = buildRow[ind]
      })
    }

    let buildRowInfo = getBuildRowNum_(url,buildURL,detailSheet)
    let buildRowNum = buildRowInfo[0]
    let buildRowCount = buildRowInfo[1]    

    // Go through every event that is "In-Process" and add them to Tracker Data Tab
    if (BuildListObj["Build Status"] === "Active") {
      let summarySheet = SpreadsheetApp.openByUrl(buildURL).getSheetByName("Build Summary");
      let sslc = summarySheet.getLastColumn();
      let sslr = summarySheet.getRange("A1").getDataRegion().getLastRow();
      let ssheaderFinder = summarySheet.createTextFinder("Designs-G-#").findNext();
      let ssheaderRow = ssheaderFinder.getRow(); 
      let ssHeaderArray = summarySheet.getRange(ssheaderRow,1,1,sslc).getValues()[0];
      let ssDetailsArray = summarySheet.getRange(ssheaderRow+2,1,sslr-ssheaderRow-1,sslc).getDisplayValues();

      // Define every col location based on Summary Sheet Header Array
      ssHeaderArray.forEach((key,ind) => {
        SummaryHeaderObj[key] = ind
      })
    
      // Remove items that have "MASTER" or "Overall Event Status" Tabs
      ssDetailsArray.some((item,ind) => {
        if (item[SummaryHeaderObj["BICEEP"]] === "DEFAULT" || item[SummaryHeaderObj["BICEEP"]] === "" ) {
          ssDetailsArray.splice(ind)
        }
      })
      
      let reqdRows = ssDetailsArray.length
      
      // Add/Delete rows based on what is already there in Tracker Data Tab
      if (reqdRows > 0) {
        if (buildRowCount === 0) {
          statusList.insertRows(statusheaderRow+1,reqdRows)
          buildRowNum = statusheaderRow + 1
        } 
        else if (buildRowCount < reqdRows) {
          let lessRows = reqdRows - buildRowCount
          statusList.insertRows(buildRowNum,lessRows)
        } 
        else if (buildRowCount > reqdRows) {
          let extraRows = buildRowCount - reqdRows
          statusList.deleteRows(buildRowNum,extraRows)
        } 
      }

      // For every row, correspond every value from Summary Sheet array to Summary Sheet Object
      ssDetailsArray.forEach((row, arrInd) => {
        //Bind the Headers with the corresponding value from Array
        Object.keys(SummaryHeaderObj).forEach((key,ind) => {
          SummaryHeaderObj[key] = row[ind]
        }) // Close SummaryHeaderObj loop
        
        //Combine values from Summary Sheet and Event Sheet into Tracker Obj
        Object.keys(StatusListObj).forEach(key => {
          if (key in SummaryHeaderObj) {
            StatusListObj[key] = SummaryHeaderObj[key]
          } else if (key in BuildListObj) {
            StatusListObj[key] = BuildListObj[key]
          } else {
            StatusListObj[key] = ""
          }
        }) // Close BuildListObj loop
        
        //Write the values to respective rows
        statusList.getRange(buildRowNum+arrInd,1,1,statusListLC).setValues([Object.values(StatusListObj)])
        
      }) 
    }
  })
}

function getBuildRowNum_(url,trackerURL,detailSheet) {

  const statusSheet = SpreadsheetApp.openByUrl(url).getSheetByName(detailSheet);
  const sslr = statusSheet.getLastRow();
  const sslc = statusSheet.getLastColumn();
  const ssheaderFinder = statusSheet.createTextFinder("Designs-G-%").findNext();
  let ssheaderRow = ssheaderFinder.getRow();
  const ssArray = statusSheet.getRange(ssheaderRow+1,sslc,sslr-ssheaderRow,1).getValues();

  let buildRowNum = 0
  let buildRowCount = 0
  let urlArray = []

  ssArray.forEach(item => {
    urlArray.push(item[0])
    if (item[0] === trackerURL) {
      buildRowCount++
    } 
  })
  
  buildRowNum = urlArray.indexOf(trackerURL) + ssheaderRow + 1  

  return [buildRowNum,buildRowCount]

}

