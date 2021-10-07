function updateKPIFile_(BuildObj) {

  let kpiTempID = '18wgsDOuAG7KVrt2bdW7TdVleuaY9F1MxAz1zJ1jtiNs'
  let parentFileID = '1tGpOxu9t1SvfCUlDbTjXWdlZDzebna6cHr_8VYEYNHM'
  let parentFile = SpreadsheetApp.openById(parentFileID)

  let sheetNames = parentFile.getSheets().map(sheets => sheets.getSheetName())
  let kpiTempFileHeaders = []
  if (!sheetNames.includes(BuildObj['Build Title'])) {
    SpreadsheetApp.openById(kpiTempID).getSheetByName("TEMPLATE").copyTo(parentFile).setName(BuildObj['Build Title'])
    let kpiTempFile = parentFile.getSheetByName(BuildObj['Build Title'])
    kpiTempFileHeaders = kpiTempFile.getRange(1,1,1,22).getDisplayValues()[0]
  } else {
    let kpiTempFile = parentFile.getSheetByName(BuildObj['Build Title'])
    kpiTempFileHeaders = kpiTempFile.getRange(1,1,1,22).getDisplayValues()[0]
  }

  let KPIHeaderObj = {}
  kpiTempFileHeaders.forEach(header => {
    KPIHeaderObj[header] = kpiTempFileHeaders.indexOf(header)
  })

  kpiTempFileHeaders.forEach(header =>{
    if (KPIHeaderObj[header] === null) {
      KPIHeaderObj[header] = ""
    } else {
      KPIHeaderObj[header] = BuildObj[header]
    }
  })

  kpiTempFileHeaders.forEach((item,ind) => {
    if (KPIHeaderObj[item] !== undefined) {
      parentFile.getSheetByName(BuildObj['Build Title']).getRange(2,ind+1,1,1).setValue(KPIHeaderObj[item])
    }
  })
  
}


function createKPIFiles(wlurl=buildsumfileURL,pbeBuildList=buildsListSheetName,defaultHeader=keyWord) {

  // Event Sheet
  const eventSheet = SpreadsheetApp.openByUrl(wlurl).getSheetByName(pbeBuildList);
  const eventSheetLR = eventSheet.getLastRow();
  const eventSheetLC = eventSheet.getLastColumn();
  const headerFinder = eventSheet.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const eventSheetHeaders = eventSheet.getRange(headerRow,1,eventSheetLR,eventSheetLC).getDisplayValues()[0];
  const eventSheetArray = eventSheet.getRange(headerRow+1,1,eventSheetLR-headerRow,eventSheetLC).getDisplayValues();  
  
  // Headers in Event Sheet
  let BuildListObj = {}
  eventSheetHeaders.forEach(header => {
    BuildListObj[header] = eventSheetHeaders.indexOf(header)
  })
  let buildNamesArray = eventSheet.getRange(headerRow+1, BuildListObj["Build Title"]+1, eventSheetLR-headerRow, 1).getValues(); 
  
  // KPI file
  let kpiTempID = '18wgsDOuAG7KVrt2bdW7TdVleuaY9F1MxAz1zJ1jtiNs'
  let parentFileID = '1YZb0ZPNnUMJlolG5ffBIK9mcLAjPE0wkodlaNAPkQ3I'
  let parentFile = SpreadsheetApp.openById(parentFileID)

  let sheetNames = parentFile.getSheets().map(sheets => sheets.getSheetName())
  let kpiTempFileHeaders = []

  buildNamesArray.forEach((name,ind) => {
    if (!sheetNames.includes(name[0])) {
      SpreadsheetApp.openById(kpiTempID).getSheetByName("TEMPLATE").copyTo(parentFile).setName(name[0])
      let kpiTempFile = parentFile.getSheetByName(name[0])
      kpiTempFileHeaders = kpiTempFile.getRange(1,1,1,22).getDisplayValues()[0]
    } 
    else {
      let kpiTempFile = parentFile.getSheetByName(name[0])
      kpiTempFileHeaders = kpiTempFile.getRange(1,1,1,22).getDisplayValues()[0]
    }

    eventSheetArray.forEach((event, row) => {
      let eventRow = eventSheetArray[row]
      eventSheetHeaders.forEach((header,ind) => {
        BuildListObj[header] = eventRow[ind]
      })
      let KPIHeaderObj = {}
      kpiTempFileHeaders.forEach((header,ind) => {
        KPIHeaderObj[header] = kpiTempFileHeaders.indexOf(header)
        if (KPIHeaderObj[header] === null) {
          KPIHeaderObj[header] = ""
        } else {
          KPIHeaderObj[header] = BuildListObj[header]
        }
      }) 
      Logger.log(KPIHeaderObj)
    })  
  })
  /*
  kpiTempFileHeaders.forEach((item,ind) => {
    if (KPIHeaderObj[item] !== undefined) {
      parentFile.getSheetByName(BuildListObj['Build Title']).getRange(2,ind+1,1,1).setValue(KPIHeaderObj[item])
    }
  })
  */
}
