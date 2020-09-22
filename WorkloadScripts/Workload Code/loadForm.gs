//Run when spreadsheet loads TEST
    function onOpen(){
      createMenu();  
    }

//Create menu dropdown
    function createMenu() {  
      const ui = SpreadsheetApp.getUi();
      const menu = ui.createMenu("PPPM Tools");
      menu.addItem("Add Event","loadNewForm");
      menu.addItem("Create New Events Tracker","createNewEventTracker");
      menu.addItem("Manage Trackers","getNewTrackerURL");
      menu.addItem("Update Events Data","createNewEventTracker");
      menu.addToUi();  
    }


//Load New Form
  function loadNewForm() {
      //Create HTML Service
        const htmlForSidebar = HtmlService.createTemplateFromFile("addEvent")
      //Get output of HTML  
         const htmlOutput = htmlForSidebar.evaluate();
      //Place output in Sidebar  
        const ui = SpreadsheetApp.getUi();
        ui.showSidebar(htmlOutput)
  }
