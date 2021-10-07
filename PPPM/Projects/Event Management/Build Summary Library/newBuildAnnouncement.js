function draftHTMLEmails_(wlurl,pbeBuildList,defaultHeader,lastEventArray){
  
  // Event Sheet
  const es = SpreadsheetApp.openByUrl(wlurl)
  const eventSheet = es.getSheetByName(pbeBuildList);
  const eventSheetLR = eventSheet.getLastRow();
  const eventSheetLC = eventSheet.getLastColumn();
  const headerFinder = eventSheet.createTextFinder(defaultHeader).findNext();
  const headerRow = headerFinder.getRow();
  const eventSheetHeaders = eventSheet.getRange(headerRow,1,eventSheetLR,eventSheetLC).getDisplayValues()[0];
  const lastEvent = [lastEventArray[0],lastEventArray[1],lastEventArray[2],lastEventArray[3],lastEventArray[4],lastEventArray[5],lastEventArray[6],lastEventArray[7]]

  const sheetNames = getSheetNames_();
  var emailTemp = HtmlService.createTemplateFromFile("newBuildEmail"); 
  let userEmail = Session.getUser().getEmail()

  let Roles = {"role":"","email": ""}

  if (sheetNames.includes("Email Distribution")) {
    const distsheet = es.getSheetByName("Email Distribution")
    const distheaderFinder = distsheet.createTextFinder("Role").findNext();
    const distheaderRow = distheaderFinder.getRow();
    const distlr = distsheet.getLastRow()
    const distlc = distsheet.getLastColumn()
    const distHeaders = distsheet.getRange(distheaderRow,1,1,distlc).getDisplayValues()[0];
    
    let DistHeaderObj = {}
    distHeaders.forEach((val, ind) => {
      DistHeaderObj[val] = ind
    })
    
    let distribution = []
    if (distlr - distheaderRow > 1) {
      const distRange = distsheet.getRange(distheaderRow+1,1,distlr-distheaderRow,distlc).getDisplayValues()
      distRange.forEach(item => {
        //DistHeaderObj[item[0]] = item[1]
        Object.keys(DistHeaderObj).forEach((key,ind) => {
          DistHeaderObj[key] = item[ind]
          Roles.role = DistHeaderObj["Role"]
          Roles.email = DistHeaderObj[key]
        })
        distribution.push(Roles.email)
      })
      
      let subject = "New Build Announcement: " + lastEventArray[0]
      emailTemp.lastEvent = lastEvent
      emailTemp.lastEventTitle = lastEventArray[0]
      
      // Create email body using above information
      let htmlMessage = emailTemp.evaluate().getContent();

      //Logger.log(htmlMessage)
      // Create Email Draft
      GmailApp.createDraft(userEmail,subject,"", {cc: distribution.join(","), htmlBody: htmlMessage})
    }
  }
}; //Close DraftHTMLEmails function
