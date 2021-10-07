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

function onOpen(){
  PBEBMSLibrary.createMenu();  
}

// Sub Functions for Each BIW to get the right template
  // BIW Templates
  function getBiwEITemplate() {
    PBEBMSLibrary.getBiwEITemplate()
  }
  function getBiwSMTemplate() {
    PBEBMSLibrary.getBiwSMTemplate()
  }

  // Chassis Templates
  function getChassisTemplate() {
    PBEBMSLibrary.getChassisTemplate()
  }

  // Electrical Templates
  function getElecHardTemplate() {
    PBEBMSLibrary.getElecHardTemplate()
  }
  function getElecModTemplate() {
    PBEBMSLibrary.getElecModTemplate()
  }
  function getElecWiringTemplate() {
    PBEBMSLibrary.getElecWiringTemplate()
  }

  // Engine System Templates
  function getEngSysTemplate() {
    PBEBMSLibrary.getEngSysTemplate()
  }

  // Int/Ext Templates
  function getIntExtTemplate() {
    PBEBMSLibrary.getIntExtTemplate()
  }

  // Powertrain Templates
  function getPwtTemplate() {
    PBEBMSLibrary.getPwtTemplate()
    
  }

  // Red Item Report
  function getRedItemDetails() {
    PBEBMSLibrary.getRedItemDetails()
  }
