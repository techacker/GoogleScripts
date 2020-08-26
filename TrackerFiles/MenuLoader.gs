//Run when spreadsheet loads
    function onOpen(){
      createMenu();  
    }

//Create menu dropdown
    function createMenu() {  
      const ui = SpreadsheetApp.getUi();
      const menu = ui.createMenu("Simplify");
      menu.addItem("Get Supplier Info","UpdateSupplierInfo");
      menu.addItem("Create RFQ Emails","createRFQEmails");
      menu.addItem("Update POs","UpdatePO");
      menu.addToUi();
    }
