function getSheetNames_() {
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet(); 
  var sheetsArray = sheet.getSheets();
  var sheetNames = [];
  
  for (var i=0; i<sheetsArray.length; i++) {
    sheetNames.push(sheetsArray[i].getSheetName());
  }
  
  return sheetNames;
}


function getUserName() {

  let userEmail = Session.getUser().getEmail()
  let userSplit = userEmail.split('@')[0]
  let userName = userSplit.split('.')

  let name = ""
  let fullName = []
  userName.forEach(val => {
    name = val.substr(0,1).toUpperCase() + val.substr(1,val.length)
    fullName.push(name)
  })

  name = fullName.join(" ")
  
  return name

}

function getTabDetails_() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Build Summary");
  const lc = summarySheet.getLastColumn();
  let lr = summarySheet.getLastRow();
  let range = summarySheet.getRange(1, 1, lr, lc).getDisplayValues();
  let headerRowNum = 0;
  
  // Get "Header" Row number
  range.filter((item, rowInd) => {
    item.filter(colname => {
      if (colname === "Designs-G-#") {
        headerRowNum = rowInd + 1
      }
    })
  })

  const headerArray = summarySheet.getRange(headerRowNum, 1, 1, lc).getValues()[0];
  const dataArray = summarySheet.getRange(headerRowNum+2, 1, lr-headerRowNum-1, lc).getValues();
  let SheetNameObj = {}

  headerArray.forEach((item, ind) => {
    SheetNameObj[item] = ind
  })

  // Get Tab Names from the Summary Table
  let tabDetails = []
  dataArray.forEach(item => {
    if (item[SheetNameObj["BICEEP"]] !== "") {
      let tabName = item[SheetNameObj["BICEEP"]] + " - " + item[SheetNameObj["System"]] + "," + item[SheetNameObj["PBE Engineer"]]
      //let engName = item[SheetNameObj["PBE Engineer"]]
      if (!tabName.startsWith("DEFAULT")) {
        tabDetails.push(tabName)
      }
    }
  })

  return tabDetails
  
}
