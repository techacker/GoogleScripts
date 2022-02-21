function archiveReadMail() {
  
  const userLabels = GmailApp.getUserLabels()
  let LabelObj = {}

  userLabels.forEach(label => {
    let name = label.getName()
    let count = label.getThreads().length
    LabelObj[name] = count
  })

  Logger.log(LabelObj)

  // Logger.log(JSON.stringify(LabelObj))
  // {"Notifications":500,"Sent Items":7,"TRAVEL":18,"Finance/House":210,"Unroll.me":500,"Career":6,"Family":10,"School":500,"Finance":500,"MISC":34,"Newsletter":500}

  //const spamCount = GmailApp.getSpamThreads().length
  //Logger.log(spamCount)

  /*
  const getUserLabelByName = GmailApp.getUserLabelByName('VDC')
  const threads = getUserLabelByName.getThreads()
  const hasThreads = Array.isArray(threads) && threads.length > 0
 
  if (hasThreads) {
    threads.forEach(thread => {
      if (!thread.isUnread()) {
        GmailApp.moveMessageToTrash(thread)
        getUserLabelByName.removeFromThread(thread)
      }
    })
  }
  */
}
