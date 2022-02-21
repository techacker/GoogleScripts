function onOpen() {
  createMenu()
}

function createMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Gmail Utility");
  menu.addItem("Get Label Info","readLabels");
  menu.addItem("Clean Emails","deleteEmails");
  menu.addItem("Clean Inbox","inboxCleanup");
  menu.addToUi();
}

function readLabels() {

  const sheetID = "1f9dFPYSvm_UU2VwvHn2XK9zrs1Q9BRr-n5x2ApyO4l8"
  const utilitysheet = SpreadsheetApp.openById(sheetID)
  const infoss = utilitysheet.getSheetByName("Info")
  infoss.getRange(1,1,1,2).setValues([["Labels", "Message Count"]])
  infoss.getRange(2,1,infoss.getLastRow(),2).clearContent()
  const userLabels = GmailApp.getUserLabels()
  
  let LabelObj = {}

  // Get all user labels
  userLabels.forEach(label => {
    let name = label.getName()
    let count = label.getThreads().length
    if (count === 0) {
      label.deleteLabel()
    } else {
      LabelObj[name] = count
      infoss.getRange(infoss.getLastRow()+1,1,1,2).setValues([[name,count]])
    }
  })
}

function deleteEmails() {

  const sheetID = "1f9dFPYSvm_UU2VwvHn2XK9zrs1Q9BRr-n5x2ApyO4l8"
  const utilitysheet = SpreadsheetApp.openById(sheetID)
  const selectionss = utilitysheet.getSheetByName("Selections")

  let selectionLR = selectionss.getLastRow()
  let delLabels = selectionss.getRange(4,1,selectionLR,1).getValues()
  let startYear = selectionss.getRange(1,2,1,1).getValue()
  let endYear = selectionss.getRange(2,2,1,1).getValue()
  const archive = selectionss.getRange(1,8,1,1).getValue()
  const trash = selectionss.getRange(2,8,1,1).getValue()

  delLabels.forEach((item,ind) => {
    if (item[0] !== "") {
      const getUserLabel = GmailApp.getUserLabelByName(item[0])
      if (getUserLabel) {
        const threads = getUserLabel.getThreads()
        const hasThreads = Array.isArray(threads) && threads.length > 0
        let deletedThreads = 0
        if (hasThreads) {
          threads.forEach(thread => {
            if (!thread.isUnread()) {
              const msgDate = thread.getLastMessageDate()
              const year = msgDate.getFullYear()
              if (year >= startYear && year <= endYear) {
                if (archive) {
                  GmailApp.moveThreadsToArchive(thread)
                  getUserLabel.removeFromThread(thread)
                  deletedThreads++
                } 
                else if (trash) {
                  GmailApp.moveThreadToTrash(thread)
                  getUserLabel.removeFromThread(thread)
                  deletedThreads++
                }
              }
            }
          })
          selectionss.getRange(4+ind,2,1,1).setValue(deletedThreads)
        }
      } 
      else {
        selectionss.getRange(4+ind,2,1,1).setValue("Label doesn't exist")
      }
    }
  })

}

function inboxCleanup() {

  const sheetID = "1f9dFPYSvm_UU2VwvHn2XK9zrs1Q9BRr-n5x2ApyO4l8"
  const utilitysheet = SpreadsheetApp.openById(sheetID)
  const selectionss = utilitysheet.getSheetByName("Selections")
  const startYear = selectionss.getRange(1,5,1,1).getValue()
  const endYear = selectionss.getRange(2,5,1,1).getValue()
  const archive = selectionss.getRange(1,8,1,1).getValue()
  const trash = selectionss.getRange(2,8,1,1).getValue()

  const inbox = GmailApp.getInboxThreads()
  const hasThreads = Array.isArray(inbox) && inbox.length > 0
  let deletedThreads = 0
  if (hasThreads) {
    inbox.forEach(thread => {
      if (!thread.isUnread()) {
        const msgDate = thread.getLastMessageDate()
        const year = msgDate.getFullYear()
        if (year >= startYear && year <= endYear) {
          if (archive) {
            GmailApp.moveThreadToArchive(thread)
            deletedThreads++
          }
          else if (trash) {
            GmailApp.moveThreadToTrash(thread)
            deletedThreads++
          } 
        }
      }
    })
    selectionss.getRange(4,4,1,2).setValues([["Inbox",deletedThreads]])
  }

}
