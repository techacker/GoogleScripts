//Run when spreadsheet loads TEST
    function onOpen(){
      createMenu();  
    }

//Create menu dropdown
    function createMenu() {  
      const ui = SpreadsheetApp.getUi();
      const menu = ui.createMenu("PPPM Tools");
      menu.addItem("Add New Event","loadNewForm");
      //menu.addItem("Create New Events Tracker","getNewTrackerURL");
      //menu.addItem("Manage Trackers","getNewTrackerURL");
      menu.addItem("Push Event Info","pushEventInfo");
      menu.addItem("Refresh Event Status","updateTrackerTab");
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
