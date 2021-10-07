function getRedItemDetails() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  let confirmation = getErrorLog_()
  
  if (confirmation) {
    const sheetNames = getSheetNames_()
    const tabDetails = getTabDetails_()
    let msg = ""
    
    clearRedItemSheet_("BOM Issues","Error Log from last run")
    clearRedItemSheet_("Design Issues","Order Issues")
    clearRedItemSheet_("Order Issues","Delivery Issues")
    clearRedItemSheet_("Delivery Issues","BOM Issues")
    
    const redItemSheet = ss.getSheetByName("Red Item Details")
    let redlc = redItemSheet.getLastColumn();
    let redheaderFinder = redItemSheet.createTextFinder("Comments to Report").findNext();
    let redHeaderRowNum = redheaderFinder.getRow(); 
    let fontType = redItemSheet.getRange(redHeaderRowNum,1,1,1).getFontFamily();
    let redItemHeaders = redItemSheet.getRange(redHeaderRowNum,1,1,redlc).getValues()[0];
    let errorRowFinder = redItemSheet.createTextFinder("Error Log from last run").findNext();
    let errorRow = errorRowFinder.getRow(); 
    let RedItemsObj = {}
    redItemHeaders.forEach((item, ind) => {
      RedItemsObj[item] = ind
    })

    // Check if all tabs exist, else write an error msg
    tabDetails.forEach(item => {
      let tabName = item.split(",")[0]
      let engName = item.split(",")[1]
      let biceep = tabName.split(" - ")[0]
      let system = tabName.split(" - ")[1]
      if (sheetNames.includes(tabName)) {
        const trkrSheet = ss.getSheetByName(tabName)
        let lc = trkrSheet.getLastColumn();
        let lr = trkrSheet.getLastRow();
        let headerFinder = trkrSheet.createTextFinder("Enable Reporting").findNext();
        let headerRow = headerFinder.getRow(); 
        let headerArray = trkrSheet.getRange(headerRow,1,1,lc).getValues()[0];
        let dataArray = trkrSheet.getRange(headerRow+1,1,lr,lc).getDisplayValues();
        let TrackerObj = {}

        headerArray.forEach((item, ind) => {
          TrackerObj[item] = ind
        })
        
        Object.keys(RedItemsObj).forEach(key => {
          if (key === "BICEEP" || key === "PBE Engineer" || key === "System" || key in TrackerObj) {
          } 
          else {
            msg = "Sheet: '" + tabName + "'" + " for '" + engName + "' : '" + key + "' was not found."
            let modifiedErrorRowFinder = redItemSheet.createTextFinder("Error Log from last run").findNext();
            let modifiedheaderRow = modifiedErrorRowFinder.getRow(); 
            redItemSheet.insertRowAfter(modifiedheaderRow)
            redItemSheet.getRange(modifiedheaderRow+1,1,1,1)
              .setValue(msg)
              .setFontFamily(fontType)
              .setFontWeight('normal')
              .setFontSize(8)
          }
        })
        
        
        let designIssues = []
        let orderIssues = []
        let deliveryIssues = []
        let bomIssues = []

        // Compile arrary of issue items from respective categories
        dataArray.filter(item => {
          if (item[TrackerObj["Enable Reporting"]] === "TRUE") {
            let partDesc = item[TrackerObj["Part Description"]]
            let partNum = item[TrackerObj["Part Number"]]
            let qty = item[TrackerObj["Total Quantity"]]
            let comments = item[TrackerObj["Comments to Report"]]
            let reqNum = item[TrackerObj["REQ #"]]
            let poNum = item[TrackerObj["PO #"]]
            let suppName = item[TrackerObj["Supplier Name"]]
            if (item[TrackerObj["Designs"]].toUpperCase() === "RED") {
              designIssues.push([biceep,system,engName,partDesc,partNum,qty,comments,reqNum,poNum,suppName])
            }
            if (item[TrackerObj["Order"]].toUpperCase() === "RED") {
              orderIssues.push([biceep,system,engName,partDesc,partNum,qty,comments,reqNum,poNum,suppName])
            }
            if (item[TrackerObj["Delivery"]].toUpperCase() === "RED") {
              deliveryIssues.push([biceep,system,engName,partDesc,partNum,qty,comments,reqNum,poNum,suppName])
            }
            if (item[TrackerObj["BOM"]].toUpperCase() === "RED") {
              bomIssues.push([biceep,system,engName,partDesc,partNum,qty,comments,reqNum,poNum,suppName])
            }
          }
        })

        // Call the function to write issues
        if (designIssues.length > 0) {
          writeRedItemDetails_("Design Issues",designIssues)
        }
        if (orderIssues.length > 0) {
          writeRedItemDetails_("Order Issues",orderIssues)
        }
        if (deliveryIssues.length > 0) {
          writeRedItemDetails_("Delivery Issues",deliveryIssues)
        }
        if (bomIssues.length > 0) {
          writeRedItemDetails_("BOM Issues",bomIssues)
        }
      }
    })
    ui.alert("Success...", "Red Item Details report has been updated.", ui.ButtonSet.OK)
    
  }
  else {
    ui.alert("Error...", "Red Item Details sheet is not found.", ui.ButtonSet.OK)
  }
  
}

function writeRedItemDetails_(keyword,issues) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const redItemSheet = ss.getSheetByName("Red Item Details")
  let redlc = redItemSheet.getLastColumn();
  let redheaderFinder = redItemSheet.createTextFinder("Comments to Report").findNext();
  let redHeaderRowNum = redheaderFinder.getRow(); 
  let fontType = redItemSheet.getRange(redHeaderRowNum,1,1,1).getFontFamily();
  let redItemHeaders = redItemSheet.getRange(redHeaderRowNum,1,1,redlc).getValues()[0];
  let issueRowFinder = redItemSheet.createTextFinder(keyword).findNext();
  let issueRow = issueRowFinder.getRow(); 
  let RedItemsObj = {}

  redItemHeaders.forEach((item, ind) => {
    RedItemsObj[item] = ind
  })
  
  issues.forEach(item => {
    Object.keys(RedItemsObj).forEach((key,ind) => {
      RedItemsObj[key] = item[ind]
    })
    redItemSheet.insertRowAfter(issueRow)
    redItemSheet.getRange(issueRow+1,1,1,redlc).clearFormat()
      .setBorder(true,true,true,true,true,true)
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setFontFamily(fontType)
    redItemSheet.getRange(issueRow+1,1,1,redlc).setValues([Object.values(RedItemsObj)])
  })

}

function getErrorLog_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = getSheetNames_()
  const tabDetails = getTabDetails_()
  let msg = ""
  let ui = SpreadsheetApp.getUi();

  // If Red Item Details Sheet is available, then write the appropriate error msg
  if (!sheetNames.includes("Red Item Details")) {
    return false
  }
  else {
    const redItemSheet = ss.getSheetByName("Red Item Details")
    let redlr = redItemSheet.getLastRow();
    let errorRowFinder = redItemSheet.createTextFinder("Error Log from last run").findNext();
    let errorRow = errorRowFinder.getRow(); 
    if (errorRow < redlr) {
      redItemSheet.getRange(errorRow+1,1,redlr-errorRow,1).clearContent()
    } 
    // Check if all tabs exist, else write an error msg
    tabDetails.forEach(item => {
      let tabName = item.split(",")[0]
      let engName = item.split(",")[1]
      if (!sheetNames.includes(tabName)) {
        msg = "Sheet: '" + tabName + "' for '" + engName + "' was not found." 
        redItemSheet.insertRowAfter(errorRow)
        redItemSheet.getRange(errorRow+1,1,1,1).setValue(msg)
      }
    })
    return true
  }
  
}


function clearRedItemSheet_(passedHeader,nextHeader) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = getSheetNames_()

  if (sheetNames.includes("Red Item Details")) {
    const redItemSheet = ss.getSheetByName("Red Item Details")
    let lc = redItemSheet.getLastColumn();
    let passedHeaderFinder = redItemSheet.createTextFinder(passedHeader).findNext();
    let passedHeaderRow = passedHeaderFinder.getRow(); 
    let nextHeaderFinder = redItemSheet.createTextFinder(nextHeader).findNext();
    let nextHeaderRow = nextHeaderFinder.getRow(); 

    // LOGIC: Since BOM Issues is the last header, to keep it clean, it was separated out from logic.
    // This adds 2 extra rows after "BOM Issues" and clears its format.
    if (passedHeader !== "BOM Issues" && nextHeaderRow-passedHeaderRow > 1) {
      redItemSheet.getRange(passedHeaderRow+1,1,nextHeaderRow-passedHeaderRow-1,lc).deleteCells(SpreadsheetApp.Dimension.ROWS)
    }
    else if (passedHeader === "BOM Issues") {
      redItemSheet.getRange(passedHeaderRow+1,1,nextHeaderRow-passedHeaderRow-1,lc).deleteCells(SpreadsheetApp.Dimension.ROWS)
      redItemSheet.insertRowsAfter(passedHeaderRow,2)
      redItemSheet.getRange(passedHeaderRow+1,1,2,lc).clearFormat()
    }
  }

}
