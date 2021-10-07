function draftChangeBulletinEmails_(wlurl, buildTitle, buildTrackerURL, changeBulletinArray){
  
  // Event Sheet
  const es = SpreadsheetApp.openByUrl(wlurl)
  const sheetNames = getSheetNames_();
  var emailTemp = HtmlService.createTemplateFromFile("buildChangeEmail"); 
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
      
      let subject = "Build Change Notification: " + buildTitle
      emailTemp.buildTitle = buildTitle
      emailTemp.buildTrackerURL = buildTrackerURL
      emailTemp.changeBulletinArray = changeBulletinArray
      
      // Create email body using above information
      let htmlMessage = emailTemp.evaluate().getContent();

      //Logger.log(htmlMessage)
      // Create Email Draft
      GmailApp.createDraft(userEmail,subject,"", {cc: distribution.join(","), htmlBody: htmlMessage})
    }
  }
}; //Close DraftHTMLEmails function
