function eventFunction_() {

  let summaryArray1 = [["Tab Name 1", "AB","05/25/2021","Active","$10.00","0.0%"],["Tab Name 2", "Mike","07/02/2021","Complete","$25.00","16.0%"]]
  let summaryArray2 = [["Tab Name 4", "XXXX","01/25/2021","In-Process","$4110.00","50.0%"]]
  let SSHeaderObj = {"Tab":1, "Engineer":2, "MRD":3, "Status":4,"Cost":5,"Completion":6}
  let events = [["2021 WL75 Grand Cherokee","http://www.google.com"],["2022 WS Grand Wagoneeer"]]
  let TrackerTab = {"Event Name":1,"Tab":2,"Engineer":3,"MRD":4,"Status":5,"Cost":6,"Completion":7, "Tracker URL":8, "Random":9}

  Object.keys(TrackerTab).forEach(key => {
    if (!Object.keys(SSHeaderObj).includes(key)) {
      SSHeaderObj[key] = Object.keys(SSHeaderObj).length + 1
    }
  })
  //SSHeaderObj["Event Name"] = Object.keys(SSHeaderObj).length + 1
  //SSHeaderObj["Tracker URL"] = Object.keys(SSHeaderObj).length + 1

  //Object.keys(TrackerTab).forEach(key => {
  //  if (Object.keys(SSHeaderObj).indexOf(key) === -1) {
  //    SSHeaderObj[key] = Object.keys(SSHeaderObj).length + 1
  //  }
  //})

  Logger.log(SSHeaderObj)
  summaryArray2.forEach(item => {
    Object.entries(SSHeaderObj).forEach((key, val) => {
      Logger.log(key + " " + item[val])
    })
  })

  /*
  summaryArray1.forEach(row=> {
    Object.entries(SSHeaderObj).forEach((key, value) => {
      Logger.log(row[value])
    })
  })
  */
}

function objFunction_() {

  let HeaderObj = {"VF":"Vehicle Family",
                  "Event Title": "Event Title",
                  "MRD": "MRD"}

  let EventObj = {}

  for (let key in HeaderObj) {
      EventObj[key] = HeaderObj[key]
  }

  let eventArray = [["WL74","WL74 Event 1","05/25/2021"],
                    ["WS","WS Sled Event 1","06/26/2021"]]

  
  for (let i=0; i<eventArray.length; i++) {
    EventObj["VF"] = eventArray[i][0]
    EventObj["Event Title"] = eventArray[i][1]
    EventObj["MRD"] = eventArray[i][2]
    Logger.log(EventObj)
  }
  
}

function testObj_() {
  let HeaderObj = {"VF":"Vehicle Family",
                    "MRD":"MRD"
                  }

  let events = [["WL74","05/25/2021"],["WS","06/26/2021"]]
  let newEvent = []

  events.forEach(row => {
    //Logger.log(row)
    row.forEach(item => {
      Logger.log(item)
    })
  })

  
  for (let key in HeaderObj) {
    events.forEach(row => {
      row.forEach(val => {
        HeaderObj[key] = val
        Logger.log(HeaderObj[key])
      })
    })
  }

}
