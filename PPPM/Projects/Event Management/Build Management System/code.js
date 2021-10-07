// ---------- PBE Build Management System (BMS)Library  ----------------------------------------
// ----------      Architect: Anurag Bansal             ----------------------------------------
// ----------          Version: 1.1.3                 ----------------------------------------
// ----------      To Suport Prototype Builds in PBE    ----------------------------------------
// ---------------------------------------------------------------------------------------------
// -----------Further Development Ideas & Change Log:
// -----------1. Create Menu Function with various templates            04/28/2021
// -----------2. Create Red Item Report                                 05/26/2021

//****** GLOBAL CONSTANTS ******

const tempID = "1zlwCBOn7C2pWJc537t-EabR2Z7aIzlGYw3YqsgEJn6w"
const biwEndItemTempID = "1avQqSc1nBIJYuwR-92hEGSX_3JGCP8rEi3n7jeKW8Gk"
const biwSheetMetalTempID = "1APNCDQGrfs9wmr5oL0wzMzt0PcpovUwr507lJ_IMWYk"
const chassisTempID = "1v3g32mLaU0xnwQxicxN5XyH2FdlZRmAHpQXX4tnYv3Q"
const intExtTempID = "1JbNzbDmgTqt5Eu-zUsZYzIBXxaQhP_-w1oZU6sH7hgM"
const engSysTempID = "128d_iRDWLM3zUOfMuvd-u4U91wpEaiBslctxGN17-n4"
const elecHardTempID = "1JLbi7TRRcg5b3rekpr4wNXp7sJYon2ifXZ8NiBbOk-w"
const elecModTempID = "1t4k70EyXmZyAvrcnGfuIljxY6cSiqBmtKAt7x46FvwI"
const elecWiringTempID = "1-OponKZfTkmO2qkb4qW1CicFPAffWznWnWn1alf0zj8"
const pwtTempID = "1buAn2BUmi1YqibIgSo9R2IJxvBKVsfI96TEeOBsHdGY"

//Run when spreadsheet loads
  function onOpen(){
    createMenu();  
  }

//Create menu dropdown
  function createMenu() {  
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("PBE Tools");

    // Create BIW Submenu
    const biwSubmenu = ui.createMenu("BIW Trackers");
    biwSubmenu.addItem("Use End Item Template","getBiwEITemplate");
    biwSubmenu.addItem("Use Sheet Metal Template","getBiwSMTemplate");
    menu.addSubMenu(biwSubmenu);

    // Create Chassis Submenu
    const chassisSubmenu = ui.createMenu("Chassis Trackers");
    chassisSubmenu.addItem("Use Chassis Template","getChassisTemplate");
    menu.addSubMenu(chassisSubmenu);

    // Create Electrical Submenu
    const elecSubmenu = ui.createMenu("Electrical Trackers");
    elecSubmenu.addItem("Use Hardware Template","getElecHardTemplate");
    elecSubmenu.addItem("Use Module Template","getElecModTemplate");
    elecSubmenu.addItem("Use Wiring Template","getElecWiringTemplate");
    menu.addSubMenu(elecSubmenu);

    // Create Engine Systems Submenu
    const engSubmenu = ui.createMenu("EngSys Trackers");
    engSubmenu.addItem("Use EngSys Template","getEngSysTemplate");
    menu.addSubMenu(engSubmenu);

    // Create Int/Ext Submenu
    const intExtSubmenu = ui.createMenu("Int/Ext Trackers");
    intExtSubmenu.addItem("Use Int/Ext Template","getIntExtTemplate");
    menu.addSubMenu(intExtSubmenu);

    // Create Powertrain Submenu
    const pwtSubmenu = ui.createMenu("Powertrain Trackers");
    pwtSubmenu.addItem("Use Powertrain Template","getPwtTemplate");
    menu.addSubMenu(pwtSubmenu);

    // Add Separator
    menu.addSeparator();
    
    // Add other items to menu
    menu.addItem("Red Item Report","getRedItemDetails");
    
    // Add Menu to UI
    menu.addToUi();
  }

// Sub Functions for Each BIW to get the right template
  // BIW Templates
  function getBiwEITemplate() {
    getTemplate(biwEndItemTempID, "Body")
  }
  function getBiwSMTemplate() {
    getTemplate(biwSheetMetalTempID, "Body")
  }

  // Chassis Templates
  function getChassisTemplate() {
    getTemplate(chassisTempID,"Chassis")
  }

  // Electrical Templates
  function getElecHardTemplate() {
    getTemplate(elecHardTempID,"Electrical")
  }
  function getElecModTemplate() {
    getTemplate(elecModTempID,"Electrical")
  }
  function getElecWiringTemplate() {
    getTemplate(elecWiringTempID,"Electrical")
  }

  // Engine System Templates
  function getEngSysTemplate() {
    getTemplate(engSysTempID,"Engine Systems")
  }

  // Int/Ext Templates
  function getIntExtTemplate() {
    getTemplate(intExtTempID,"Int/Ext")
  }

  // Powertrain Templates
  function getPwtTemplate() {
    getTemplate(pwtTempID,"Powertrain")
  }
