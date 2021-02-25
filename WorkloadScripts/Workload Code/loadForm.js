//Run when spreadsheet loads TEST
    function onOpen(){
      createMenu();  
    }

//Create menu dropdown
    function createMenu() {  
      const ui = SpreadsheetApp.getUi();
      const menu = ui.createMenu("PPPM Tools");
      //menu.addItem("Add New Event","loadModifiedNewForm");
      //menu.addItem("Create New Events Tracker","getNewTrackerURL");
      //menu.addItem("Manage Trackers","getNewTrackerURL");
      menu.addItem("Create Tracker","createTracker");
      menu.addItem("Push Event Info","pushEventInfo");
      menu.addItem("Refresh Event Status","updateTrackerTab");
      menu.addItem("Archive Trackers","archiveTrackers");
      //menu.addItem("Add Modified New Event","loadModifiedNewForm");
      menu.addToUi();  
    }


//Load New Form
  function loadNewForm() {
      //Create HTML Service
        const htmlForSidebar = HtmlService.createTemplateFromFile("addEvent")
      //Get output of HTML  
         const htmlOutput = htmlForSidebar.evaluate();
    htmlOutput.setTitle("Add New Event");
      //Place output in Sidebar  
        const ui = SpreadsheetApp.getUi();
        ui.showSidebar(htmlOutput)
  }

  //Load New Form
  function loadModifiedNewForm() {
      //Create HTML Service
        const htmlForSidebar = HtmlService.createTemplateFromFile("addEventNew")
      //Get output of HTML  
         const htmlOutput = htmlForSidebar.evaluate();
    htmlOutput.setTitle("Add New Event");
      //Place output in Sidebar  
        const ui = SpreadsheetApp.getUi();
        ui.showSidebar(htmlOutput)
  }


//Load Modify Form
  function loadModifyForm() {
      //Create HTML Service
        const htmlForSidebar = HtmlService.createTemplateFromFile("ModifyEvent")
      //Get output of HTML  
         const htmlOutput = htmlForSidebar.evaluate();
      //Place output in Sidebar  
        const ui = SpreadsheetApp.getUi();
        ui.showSidebar(htmlOutput)
  }
