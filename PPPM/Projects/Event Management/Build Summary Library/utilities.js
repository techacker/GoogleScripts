function lookupColumns(url=buildsumfileURL, sheetName=buildsListSheetName,stringToFind=keyWord,colArray=buildListLookupArray) {
  let as = SpreadsheetApp.openByUrl(url).getSheetByName(sheetName);
  let lc = as.getLastColumn();
  let textFinder = as.getDataRange().createTextFinder(stringToFind).findNext();
  let headerRow = textFinder.getRow();
  let headers = as.getRange(headerRow,1,1,lc).getValues()[0];
  let HeaderIndexObj = {};

  // Create an object with Header Index Columns for every header entry
  headers.forEach(item => {
    HeaderIndexObj[item] = headers.indexOf(item);
  })

  // Logger.log(HeaderIndexObj)
  // {Event Status=7.0, % Exception=31.0, MY=2.0, Total # of Parts=16.0, Initial Part #s=39.0, Earliest MRD=6.0, # REQ Submitted=22.0, % Late=32.0,   Ship-to Code=12.0, Program Manager=8.0, #in Rec'd=37.0, % PO=18.0, # Not Defined=29.0, Event Title=5.0, VF=1.0, #in RFQ Pending=34.0, % Not Defined=33.0, Cost=20.0, # Exception=27.0, Date Added=38.0, # Cancelled=25.0, % On Time=30.0, % REQ=17.0, Event Name=3.0, # Parts Received=24.0, Notes=15.0, Location of Event=11.0, # PO Issued=23.0, Attention-to=14.0, % Cancelled=21.0, #in REQ=35.0, Tracker URL=40.0, WBS Code=10.0, Requestor=9.0, Ship-to Address=13.0, % Received=19.0, Event Type=4.0, # Late=28.0, #in PO=36.0, # On Time=26.0}

  let LookupObj = {};
  colArray.forEach(val => {
    if (val in HeaderIndexObj) {
      LookupObj[val] = HeaderIndexObj[val];
    }
  })
  //Logger.log(LookupObj)
  return LookupObj
  
}

function getUserName() {

  let userEmail = Session.getUser().getEmail()
  const regex = /^([a-zA-Z\.]*)/g
  const userName = userEmail.match(regex)[0]
  let name = userName.split('.')
  let completeName = ""
  let fullName = []

  name.forEach(val => {
    completeName = val.substr(0,1).toUpperCase() + val.substr(1,val.length)
    fullName.push(completeName)
  })
  
  completeName = fullName.join(" ")
  return completeName

}

function parseURL_(url) {
  
  var urlArray = url.split('/');
  var id = urlArray[5];
  var link = url.split(id);
  var linktoURL = link[0] + id + '/';
  return linktoURL;
  
}


function getTrackerURL_(VehFamFileNames, EventTitle) {
  
  for (var i=0; i<VehFamFileNames.length; i++) {
    var fileName = EventTitle + " Tracker"
    if (VehFamFileNames[i] === fileName) {
      var tracker = DriveApp.getFilesByName(fileName).next();
      var url = tracker.getUrl();
      var link = parseURL_(url);
      return url;
    }
  }
}


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


function addNewDataRow_(parentFileID,sheetName,keyWord,PassedObj) {
  
  let parentFile = SpreadsheetApp.openById(parentFileID)
  let tempFile = parentFile.getSheetByName(sheetName)
  const tempLR = tempFile.getLastRow();
  const tempLC = tempFile.getLastColumn();
  const headerFinder = tempFile.createTextFinder(keyWord).findNext();
  const headerRow = headerFinder.getRow();
  const headerCol = headerFinder.getColumn();
  const headers = tempFile.getRange(headerRow,1,1,tempLC).getDisplayValues()[0];
  let keyColArray = tempFile.getRange(headerRow+1, headerCol, tempLR-headerRow, 1).getValues(); 

  let SheetHeaderObj = {}
  headers.forEach(header => {
    SheetHeaderObj[header] = headers.indexOf(header)
  })

  headers.forEach(header =>{
    if (SheetHeaderObj[header] === null) {
      SheetHeaderObj[header] = ""
    } else {
      SheetHeaderObj[header] = PassedObj[header]
    }
  })
  
  // If the build is not already there, it will be added as a last row, otherwise, the row with the Tracker URL will be updated.
  let newDataRow = tempLR+1
  keyColArray.filter((trackerURL, ind) => {
    if (trackerURL.includes(PassedObj[keyWord])) {
      newDataRow = ind + headerRow + 1
    }
  })
  
  // Get the previous event details
  const lastRowRange = tempFile.getRange(tempLR,1,1,tempLC);
  const lastRowFormulas = lastRowRange.getFormulas()[0];
  const newRowRange = tempFile.getRange(newDataRow,1,1,tempLC);

  // Get values from newEventArray and formulas from lastRowFormulas Array and put them in new row.
  Object.keys(SheetHeaderObj).forEach((key, ind) => {
    if (lastRowFormulas[ind] !== "") {
      tempFile.getRange(tempLR,ind+1,1,1).copyTo(tempFile.getRange(newDataRow,ind+1,1,1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false)
    } else if (SheetHeaderObj[key] !== undefined) {
      tempFile.getRange(newDataRow,ind+1,1,1).setValue(SheetHeaderObj[key])
    } 
  })

  // Copy the format from previous row
  lastRowRange.copyTo(newRowRange,SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false)
}


function addNewDataWithRules_(parentFileID,sheetName,keyWord,PassedObj) {

  let parentFile = SpreadsheetApp.openById(parentFileID)
  let tempFile = parentFile.getSheetByName(sheetName)
  const tempLR = tempFile.getLastRow();
  const tempLC = tempFile.getLastColumn();
  const headerFinder = tempFile.createTextFinder(keyWord).findNext();
  const headerRow = headerFinder.getRow();
  const headerCol = headerFinder.getColumn();
  const headers = tempFile.getRange(headerRow,1,1,tempLC).getDisplayValues()[0];
  let keyColArray = tempFile.getRange(headerRow+1, headerCol, tempLR-headerRow, 1).getValues(); 

  let SheetHeaderObj = {}
  headers.forEach(header => {
    SheetHeaderObj[header] = headers.indexOf(header)
  })

  headers.forEach(header =>{
    if (SheetHeaderObj[header] === null) {
      SheetHeaderObj[header] = ""
    } else {
      SheetHeaderObj[header] = PassedObj[header]
    }
  })
  
  // If the build is not already there, it will be added as a last row, otherwise, the row with the Tracker URL will be updated.
  let newDataRow = tempLR+1
  keyColArray.filter((trackerURL, ind) => {
    if (trackerURL.includes(PassedObj[keyWord])) {
      newDataRow = ind + headerRow + 1
    }
  })
  
  // Get the previous event details
  const lastRowRange = tempFile.getRange(tempLR,1,1,tempLC);
  const lastRowFormulas = lastRowRange.getFormulas()[0];
  const newRowRange = tempFile.getRange(newDataRow,1,1,tempLC);

  // Get values from newEventArray and formulas from lastRowFormulas Array and put them in new row.
  Object.keys(SheetHeaderObj).forEach((key, ind) => {
    let cellRule = lastRowRange.getDataValidations()[0]
    if (lastRowFormulas[ind] !== "") {
      tempFile.getRange(tempLR,ind+1,1,1).copyTo(tempFile.getRange(newDataRow,ind+1,1,1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false)
    } 
    else if (SheetHeaderObj[key] !== undefined) {
      tempFile.getRange(newDataRow,ind+1,1,1).setValue(SheetHeaderObj[key])
    } 
    else if (cellRule[ind] !== null) {
      let criteria = cellRule[ind].getCriteriaType()
      let args = cellRule[ind].getCriteriaValues()
      tempFile.getRange(newDataRow,ind+1,1,1).setDataValidation(cellRule[ind])
    } 
  })

  // Copy the format from previous row
  lastRowRange.copyTo(newRowRange,SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false)
}
