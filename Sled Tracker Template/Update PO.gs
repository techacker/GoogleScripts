// ********* Start: Update PO in Sled Tracker Sheet

function updateEntireSeriesPO(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkSledTemplate() function to get the required headerRow and column indexes.
  var headerRow = checkSledTemplate()[0];  
  var indices = checkSledTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
    
  // Get the index values of all necessary columns for Reqs.
  /*
  Logger.log(indices) - returns
  indices =  [0.0, 
              1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 
              9.0, 11.0, 
              13.0, 14.0, 15.0, 16.0, 
              18.0, 19.0, 20.0, 21.0, 
              23.0, 24.0, 25.0, 26.0, 
              28.0, 29.0, 30.0, 31.0, 
              33.0, 34.0, 35.0, 36.0, 
              38.0, 39.0, 40.0, 41.0,
              43.0, 44.0, 45.0, 46.0, 
              48.0, 49.0, 50.0, 51.0, 
              53.0, 54.0, 55.0, 56.0, 
              60.0]
  
  */
  
  var REQSeries1Col = indices[9]+1;  
  var PONum1Col = indices[10]+1;
  var POIssueDate1Col = indices[11]+1;
  
  var REQColumns = [REQSeries1Col, REQSeries1Col+5, REQSeries1Col+10, REQSeries1Col+15, REQSeries1Col+20, REQSeries1Col+25, REQSeries1Col+30, REQSeries1Col+35, REQSeries1Col+40];
  var POColumns = [PONum1Col, PONum1Col+5, PONum1Col+10, PONum1Col+15, PONum1Col+20, PONum1Col+25, PONum1Col+30, PONum1Col+35, PONum1Col+40];
  var POIssueDateColumns = [POIssueDate1Col, POIssueDate1Col+5, POIssueDate1Col+10, POIssueDate1Col+15, POIssueDate1Col+20, POIssueDate1Col+25, POIssueDate1Col+30, POIssueDate1Col+35, POIssueDate1Col+40];
    
  // Get PO details from Master File using PO Update Function
  var POs = getPOfromMasterPPPMREQFile();
  
  // Update POs for REQs that have POs issued for the entire series
  
  for (var i=0; i<REQColumns.length; i++) { 
    var REQNum = ss.getRange(headerRow+1, REQColumns[i], lr, 1).getDisplayValues();
    for (var j=0; j<REQNum.length; j++) {
      for (var k=0; k<POs.length; k++) {
        if (REQNum[j][0] !== "" && REQNum[j][0] === POs[k][0]) {
          ss.getRange(headerRow+j+1, POColumns[i], 1, 1).setValue(POs[k][1]);
          ss.getRange(headerRow+j+1, POIssueDateColumns[i], 1, 1).setValue(POs[k][2]);
        };
      };
    };
  };
  
  Browser.msgBox("All POs released to the suppliers are updated in the current sheet!");
 
}; 

// ********* End: Update PO in Sled Tracker Sheet



// ********* Start: Get Requisition Info

function GetReqNo(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  
  // Call checkSledTemplate() function to get the required headerRow and column indexes.
  
  var headerRow = checkSledTemplate()[0];  
  var indices = checkSledTemplate()[1];  
  var header = ss.getRange(headerRow, 1, 1, lc).getValues()[0];
  var range = ss.getRange(headerRow + 1, 1, lr-headerRow, lc).getValues()[0];
  
  // Get the index values of all necessary columns for Reqs.
 
  var REQSeries1Col = indices[9]+1;  
  var PONum1Col = indices[10]+1;
  
  var REQColumns = [REQSeries1Col, REQSeries1Col+5, REQSeries1Col+10, REQSeries1Col+15, REQSeries1Col+20, REQSeries1Col+25, REQSeries1Col+30, REQSeries1Col+35, REQSeries1Col+40];
  var POColumns = [PONum1Col, PONum1Col+5, PONum1Col+10, PONum1Col+15, PONum1Col+20, PONum1Col+25, PONum1Col+30, PONum1Col+35, PONum1Col+40];
  
  var REQs = [];
  
  for (var i=0; i<REQColumns.length; i++) { 
    var REQNum = ss.getRange(headerRow+1, REQColumns[i], lr, 1).getDisplayValues();
    for (var j=0; j<REQNum.length; j++) {
      if (REQNum[j][0].length >= 6 && REQs.indexOf(REQNum[j][0]) === -1.0) { // Only add unique REQs
        REQs.push(REQNum[j][0]);
      };
    };
  };
  
  return REQs;
  
}; 

// ********* End: Get Requisition Info
