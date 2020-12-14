function onOpen(){
  PPPMWorkloadGS.createMenu();  
  
}

//Load New Form
function loadNewForm() {
  //Create HTML Service
  const htmlForSidebar = HtmlService.createTemplateFromFile("addEvent");
  //Get output of HTML  
  const htmlOutput = htmlForSidebar.evaluate();
  htmlOutput.setTitle("Add New Event");
  //Place output in Sidebar  
  const ui = SpreadsheetApp.getUi();
  ui.showSidebar(htmlOutput)
}

