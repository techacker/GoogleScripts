//Run when spreadsheet loads
    function onOpen(){
      createMenu();  
    }

//Create menu dropdown
    function createMenu() {  
      const ui = SpreadsheetApp.getUi();
      const menu = ui.createMenu("PPPM Tools");
      
      // Create Submenu
      //var submenu = ui.createMenu("Fetch Info");
      //submenu.addItem("Get Supplier Info","UpdateSupplierInfo");
      //submenu.addItem("Get PO Details","UpdatePO");
      //menu.addSubMenu(submenu);
      //menu.addSeparator();
      
      // Add other items to menu
      menu.addItem("Create New Tracker Sheet","addNewRowInSummarySheet");
      menu.addItem("Get Supplier Info","UpdateSupplierInfo");
      menu.addItem("Refresh Order Details","UpdatePO");
      
      menu.addSeparator();
      menu.addItem("Draft RFQ Emails","createRFQEmails");
      menu.addItem("Generate RFQ Forms","createRFQForms");
      
      // Push status menu
      menu.addSeparator();
      menu.addItem("Push Program Status","pushEventUpdates");
      
      
      // Add Menu to UI
      menu.addToUi();
    }
