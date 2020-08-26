// ********* Start: Retrieve Purchase Order information from Master PPPPM REQ Google Sheet updated from SharePoint manually at this point

function getPOfromMasterPPPMREQFile() {
  // Master PPPM REQ File: ID - 1Aet1W9VJAnQnSNo6tfegHlpswTS65hx9vUTkqT0b1Hs/edit#gid=577545728
  
  var masterREQFile = SpreadsheetApp.openById("1Aet1W9VJAnQnSNo6tfegHlpswTS65hx9vUTkqT0b1Hs");
  var fileName = masterREQFile.getName();
  var dataRange = masterREQFile.getDataRange();
  var lr = dataRange.getLastRow()-1;
  var lc = dataRange.getLastColumn();
  var reqData = dataRange.getDisplayValues();
  var PODetails = [];
  var REQs = GetReqNo();
  
  for (var i=0; i<REQs.length; i++) {
    for (var j=0; j<lr; j++) {
      if (REQs[i] == reqData[j][0]) {
        PODetails.push([REQs[i], reqData[j][1], reqData[j][2]]);
      };
    };
  };
  
  return PODetails;
  
}


// ********* End: Retrieve Purchase Order information from Master PPPPM REQ Google Sheet updated from SharePoint manually at this point
