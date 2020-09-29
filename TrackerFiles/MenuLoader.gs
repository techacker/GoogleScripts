//Run when spreadsheet loads
    function onOpen(){
      createMenu();  
    }

//Create menu dropdown
    function createMenu() {  
      const ui = SpreadsheetApp.getUi();
      const menu = ui.createMenu("PPPM Tools");
      
      // Create Submenu
      var submenu = ui.createMenu("Fetch Info");
      submenu.addItem("Get Supplier Info","UpdateSupplierInfo");
      submenu.addItem("Get PO Details","UpdatePO");
      submenu.addItem("Update Program Status","updateStatusColumn");
      menu.addSubMenu(submenu);
      menu.addSeparator();
      
      // Add other items to menu
      menu.addItem("Create New Tracker Sheet","addNewRowInSummarySheet");
      menu.addItem("Generate RFQ Emails","createRFQEmails");
      menu.addItem("Generate RFQ Forms","createRFQForms");
      
      
      // Add Menu to UI
      menu.addToUi();
    }
