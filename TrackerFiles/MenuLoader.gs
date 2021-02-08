//Run when spreadsheet loads
    function onOpen(){
      createMenu();  
    }

//Create menu dropdown
    function createMenu() {  
      const ui = SpreadsheetApp.getUi();
      const menu = ui.createMenu("PPPM Tools");
      
      // Add other items to menu
      menu.addItem("Create New Tracker Tab","addNewRowInSummarySheet");
      menu.addItem("Get Supplier Info","UpdateSupplierInfo");
      menu.addItem("Refresh Order Details","UpdatePO");
      
      menu.addSeparator();
      menu.addItem("Draft RFQ Emails","createRFQEmails");
      //menu.addItem("Generate RFQ Forms","createRFQForms");
      
      // Push status menu
      menu.addSeparator();
      menu.addItem("Push Program Status","pushEventUpdates");
      menu.addSeparator();
      
      // Create Submenu
      var submenu = ui.createMenu("Beta Modules");
      submenu.addItem("Generate RFQ Forms","createRFQForms");
      
      // Add Submenu to Menu
      menu.addSubMenu(submenu);
      
      // Add Menu to UI
      menu.addToUi();
    }
