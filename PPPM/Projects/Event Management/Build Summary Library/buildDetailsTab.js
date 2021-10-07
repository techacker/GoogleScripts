function getEventStatus(wlurl=buildsumfileURL,pbeBuildList=buildsListSheetName,defaultHeader=keyWord,lookupArray=buildDetailLookupArray,trackerSheet=buildDetailsSheetName) {

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
  const trackerTabSheet = SpreadsheetApp.openByUrl(wlurl).getSheetByName(trackerSheet);
  const ttlr = trackerTabSheet.getLastRow();
  const ttlc = trackerTabSheet.getLastColumn();
  const ttheaderFinder = trackerTabSheet.createTextFinder("Tab").findNext();
  let ttheaderRow = ttheaderFinder.getRow();
  const ttHeaders = trackerTabSheet.getRange(ttheaderRow,1,1,ttlc).getDisplayValues()[0];
  const ttArray = trackerTabSheet.getRange(ttheaderRow+1,1,ttlr,ttlc).getValues();
  let TrackerTabObj = lookupColumns(wlurl,trackerSheet,"Tab",ttHeaders);

  // Compile an array without MASTER and OVERALL EVENT STATUS
  trackerURLArray.forEach((trackerURL,row) => {
    // Correspond every value from Event Sheet Obj with the values from the array
    if (trackerURL[0] !== "") {
      let eventRow = eventSheetArray[row]
      eventSheetHeaders.forEach((heading, ind) => {
        BuildListObj[heading] = eventRow[ind]
      })
    }

    let eventRowInfo = getEventRowNum_(wlurl,trackerURL[0],trackerSheet)
    let eventRow = eventRowInfo[0]
    let currentEventRowCount = eventRowInfo[1]
  
    // Go through every event that is "In-Process" and add them to Tracker Data Tab
    if (BuildListObj["Event Status"] !== "Archived") {
      let summarySheet = SpreadsheetApp.openByUrl(trackerURL[0]).getSheetByName("Summary");
      let sslc = summarySheet.getLastColumn();
      let sslr = summarySheet.getRange("A1").getDataRegion().getLastRow();
      let ssheaderFinder = summarySheet.createTextFinder("Total # of Parts").findNext();
      let ssheaderRow = ssheaderFinder.getRow(); 
      let ssHeaderArray = summarySheet.getRange(ssheaderRow,1,1,sslc).getValues()[0];
      let ssDetailsArray = summarySheet.getRange(ssheaderRow+1,1,sslr-ssheaderRow,sslc).getDisplayValues();

      // Define every col location based on Summary Sheet Header Array
      ssHeaderArray.forEach((key,ind) => {
        SummaryHeaderObj[key] = ind
      })

      // Remove items that have "MASTER" or "Overall Event Status" Tabs
      ssDetailsArray.some((item,ind) => {
        if (item[SummaryHeaderObj["Tab"]] === "MASTER" || item[SummaryHeaderObj["Tab"]] === "Overall Event Status" ) {
          ssDetailsArray.splice(ind)
        }
      })

      let reqdRows = ssDetailsArray.length

      // Add/Delete rows based on what is already there in Tracker Data Tab
      if (reqdRows > 0) {
        if (currentEventRowCount === 0) {
          trackerTabSheet.insertRows(ttheaderRow+1,reqdRows)
          eventRow = ttheaderRow + 1
        } 
        else if (currentEventRowCount < reqdRows) {
          let lessRows = reqdRows - currentEventRowCount
          trackerTabSheet.insertRows(eventRow,lessRows)
        } 
        else if (currentEventRowCount > reqdRows) {
          let extraRows = currentEventRowCount - reqdRows
          trackerTabSheet.deleteRows(eventRow,extraRows)
        } 
      }
      
      // For every row, correspond every value from Summary Sheet array to Summary Sheet Object
      ssDetailsArray.forEach((row, arrInd) => {
        //Bind the Headers with the corresponding value from Array
        Object.keys(SummaryHeaderObj).forEach((key,ind) => {
          SummaryHeaderObj[key] = row[ind]
        }) // Close SummaryHeaderObj loop
        
        //Combine values from Summary Sheet and Event Sheet into Tracker Obj
        Object.keys(TrackerTabObj).forEach(key => {
          if (key in SummaryHeaderObj) {
            TrackerTabObj[key] = SummaryHeaderObj[key]
          } else if (key in BuildListObj) {
            TrackerTabObj[key] = BuildListObj[key]
          }
        }) // Close TrackerTabObj loop
      
        // The one which is not captured by above function is "Event Title - Tab" and "Days until MRD", since these are unique
        TrackerTabObj["Event Title - Tab"] = TrackerTabObj["Event Title"] + " - " + TrackerTabObj["Tab"]
        // Get the days difference between MRD and today
        let mrd = new Date(TrackerTabObj["MRD"])
        let dateDiff = mrd - todaysDate   // It is in milliseconds
        let daysUntilMRD = Math.floor(dateDiff/(1000*60*60*24))     // Convert to days and round off

        // Other cells that have specific formulas
        TrackerTabObj["Days until MRD"] = daysUntilMRD
        TrackerTabObj["#in RFQ Pending"] = TrackerTabObj["Total # of Parts"] - TrackerTabObj["# REQ Submitted"]
        TrackerTabObj["#in REQ"] = TrackerTabObj["# REQ Submitted"] - TrackerTabObj["# PO Issued"]
        TrackerTabObj["#in PO"] = TrackerTabObj["# PO Issued"] - TrackerTabObj["# Parts Received"]
        TrackerTabObj["#in Rec'd"] = TrackerTabObj["# Parts Received"]

        // Write everything into the respective cells
        trackerTabSheet.getRange(eventRow+arrInd,1,1,ttlc).setValues([Object.values(TrackerTabObj)])
      })      
    }
  })

} // Close function

function getEventRowNum_(wlurl,trackerURL,trackerSheet) {

  const trackerTabSheet = SpreadsheetApp.openByUrl(wlurl).getSheetByName(trackerSheet);
  const ttlr = trackerTabSheet.getLastRow();
  const ttlc = trackerTabSheet.getLastColumn();
  const ttheaderFinder = trackerTabSheet.createTextFinder("Tab").findNext();
  let ttheaderRow = ttheaderFinder.getRow();
  const ttArray = trackerTabSheet.getRange(ttheaderRow,ttlc,ttlr,1).getValues();
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
