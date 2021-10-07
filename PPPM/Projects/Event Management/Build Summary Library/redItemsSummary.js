function getRedItemsStatus(url=buildsumfileURL,pbeBuildList=buildsListSheetName,redItemSheet=redItemsSheetName,defaultHeader=keyWord) {

  // Build Sheet
  const buildList = SpreadsheetApp.openByUrl(url).getSheetByName(pbeBuildList);
  const buildListLR = buildList.getLastRow();
  const buildListLC = buildList.getLastColumn();
  const headerFinder = buildList.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const buildListHeaders = buildList.getRange(headerRow,1,1,buildListLC).getDisplayValues()[0];

  // Red Items Sheet
  const redItemsList = SpreadsheetApp.openByUrl(url).getSheetByName(redItemSheet);
  const redItemsListLC = redItemsList.getLastColumn();
  const redItemsheaderFinder = redItemsList.createTextFinder("Issue Category").findNext();
  const redItemsheaderRow = redItemsheaderFinder.getRow();
  const fontType = redItemsList.getRange(redItemsheaderRow,1,1,1).getFontFamily();
  const redItemsListHeaders = redItemsList.getRange(redItemsheaderRow,1,1,redItemsListLC).getDisplayValues()[0];
  
  const dataAdded = new Date().toLocaleDateString();

  // Headers in Build List
  let BuildListObj = {}
  // Headers in Status List
  let RedItemsListObj = {}
  // Headers in Tracker Sheet
  let SummaryObj = {}

  // Create Build List Object with relevant column indexes
  buildListHeaders.forEach(item => {
    BuildListObj[item] = buildListHeaders.indexOf(item)
  })

  // Create Status List Object with relevant column indexes
  redItemsListHeaders.forEach(item => {
    RedItemsListObj[item] = redItemsListHeaders.indexOf(item)
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
    
    let buildRowInfo = getRedItemBuildRowNum_(url,buildURL,redItemSheet)
    let buildRowNum = buildRowInfo[0]
    let buildRowCount = buildRowInfo[1]

    // Go through every event that is "In-Process" and add them to Tracker Data Tab
    if (BuildListObj["Build Status"] === "Active") {
      let summarySheet = SpreadsheetApp.openByUrl(buildURL).getSheetByName("Red Item Details");
      let sslc = summarySheet.getLastColumn();
      let sslr = summarySheet.getRange("A1").getDataRegion().getLastRow();
      let ssheaderFinder = summarySheet.createTextFinder("Comments to Report").findNext();
      let ssheaderRow = ssheaderFinder.getRow();
      let designIssuesFinder = summarySheet.createTextFinder("Design Issues").findNext();
      let designHeaderRow = designIssuesFinder.getRow();
      let orderIssuesFinder = summarySheet.createTextFinder("Order Issues").findNext();
      let orderHeaderRow = orderIssuesFinder.getRow();
      let deliveryIssuesFinder = summarySheet.createTextFinder("Delivery Issues").findNext();
      let deliveryHeaderRow = deliveryIssuesFinder.getRow();
      let bomIssuesFinder = summarySheet.createTextFinder("BOM Issues").findNext();
      let bomHeaderRow = bomIssuesFinder.getRow();
      let designIssuesArray = []
      let orderIssuesArray = []
      let deliveryIssuesArray = []
      let bomIssuesArray = []
      let BuildIssues = {}

      let ssHeaderArray = summarySheet.getRange(ssheaderRow,1,1,sslc).getValues()[0];

      // Define every col location based on Summary Sheet Header Array
      ssHeaderArray.forEach((key,ind) => {
        SummaryObj[key] = ind
      })

      let designReqdRows = 0
      let orderReqdRows = 0
      let deliveryReqdRows = 0
      let bomReqdRows = 0
      let reqdRows = 0

      // 12 14 19 20 21

      // For every type of issue, compile an array if there are issues.
      if (orderHeaderRow-designHeaderRow > 1) {
        designIssuesArray = summarySheet.getRange(designHeaderRow+1,1,orderHeaderRow-designHeaderRow-1,sslc).getDisplayValues();
        designReqdRows = designIssuesArray.length
      }
      if (deliveryHeaderRow-orderHeaderRow > 1) {
        orderIssuesArray = summarySheet.getRange(orderHeaderRow+1,1,deliveryHeaderRow-orderHeaderRow-1,sslc).getDisplayValues();
        orderReqdRows = orderIssuesArray.length
      }
      
      if (bomHeaderRow-deliveryHeaderRow > 1) {
        deliveryIssuesArray = summarySheet.getRange(deliveryHeaderRow+1,1,bomHeaderRow-deliveryHeaderRow-1,sslc).getDisplayValues();
        deliveryReqdRows = deliveryIssuesArray.length
      }
      if (sslr-bomHeaderRow >= 1) {
        bomIssuesArray = summarySheet.getRange(bomHeaderRow+1,1,sslr-bomHeaderRow,sslc).getDisplayValues();
        bomReqdRows = bomIssuesArray.length
      }

      reqdRows = designReqdRows + orderReqdRows + deliveryReqdRows + bomReqdRows
      
      BuildIssues["Design Issues"] = designIssuesArray
      BuildIssues["Order Issues"] = orderIssuesArray
      BuildIssues["Delivery Issues"] = deliveryIssuesArray
      BuildIssues["BOM Issues"] = bomIssuesArray      
      
      // Add/Delete rows based on what is already there in Tracker Data Tab
      if (reqdRows > 0) {
        if (buildRowCount === 0) {
          redItemsList.insertRows(redItemsheaderRow+1,reqdRows)
          buildRowNum = redItemsheaderRow + 1
        } 
        else if (buildRowCount < reqdRows) {
          let lessRows = reqdRows - buildRowCount
          redItemsList.insertRows(buildRowNum,lessRows)
        } 
        else if (buildRowCount > reqdRows) {
          let extraRows = buildRowCount - reqdRows
          redItemsList.deleteRows(buildRowNum,extraRows)
        } 
      }
      

      let rowInd = 0
      Object.keys(BuildIssues).forEach(key => {
        let issuesCount = BuildIssues[key].length
        if (issuesCount !== 0) {
          BuildIssues[key].forEach(item => {
            Object.keys(SummaryObj).forEach((header, ind) => {
              SummaryObj[header] = item[ind]
            })
            Object.keys(RedItemsListObj).forEach(heading => {
              if (heading in SummaryObj) {
                RedItemsListObj[heading] = SummaryObj[heading]
              } else if (heading in BuildListObj) {
                RedItemsListObj[heading] = BuildListObj[heading]
              } else {
                RedItemsListObj[heading] = ""
              }
            })
            RedItemsListObj["Issue Category"] = key
            redItemsList.getRange(buildRowNum+rowInd,1,1,redItemsListLC)
              .setValues([Object.values(RedItemsListObj)])
              .setFontFamily(fontType)
              .setHorizontalAlignment("center")
              .setFontWeight("normal")
              .setFontSize(10)
            rowInd++
          })
        }
      })
    }
  })
}

function getRedItemBuildRowNum_(url,buildURL,redItemSheet) {

  const statusSheet = SpreadsheetApp.openByUrl(url).getSheetByName(redItemSheet);
  const sslr = statusSheet.getLastRow();
  const sslc = statusSheet.getLastColumn();
  const ssheaderFinder = statusSheet.createTextFinder("Issue Category").findNext();
  let ssheaderRow = ssheaderFinder.getRow();
  const ssArray = statusSheet.getRange(ssheaderRow+1,sslc,sslr-ssheaderRow,1).getValues();
  
  let buildRowNum = 0
  let buildRowCount = 0
  let buildURLArray = []

  ssArray.forEach(item => {
    buildURLArray.push(item[0])
    if (item[0] === buildURL) {
      buildRowCount++
    } 
  })
  
  buildRowNum = buildURLArray.indexOf(buildURL) + ssheaderRow + 1  

  return [buildRowNum,buildRowCount]
  
}

