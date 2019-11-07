//Setters and on setup functions

function setColumnProperties(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Log");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var properties = PropertiesService.getScriptProperties();
  properties.deleteAllProperties();
  properties.setProperty('Billable Column Number', '');
  properties.setProperty('Rejection Column Number', '');

  for(var i = 0; i < headers.length; i++){
    if(headers[i].slice(-19) == 'Ready for Invoicing'){
      //billable column
      var temp = properties.getProperty('Billable Column Number');
      properties.setProperty('Billable Column Number', temp+(i+1)+',');
    }
    else if(headers[i].slice(-22) == 'Invoice Rejection Date'){
      //rejection column
      var temp = properties.getProperty('Rejection Column Number');
      properties.setProperty('Rejection Column Number', temp+(i+1)+',');
    }
  }
}

function setChangeTable(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceLog = ss.getSheetByName("Invoice Log");
  var changetable = ss.getSheetByName("changetable");
  
  //if the change table doesn't exist yet, then add it
  if(changetable == null){
    ss.insertSheet("changetable");
    changetable = ss.getSheetByName("changetable");
    changetable.hideSheet();
  }
  
  changetable.clear();
  
  //add in FQNID, NFID, Work Order ID, and GDB Work Order ID
  var invoiceInfo = invoiceLog.getSheetValues(1, 4, invoiceLog.getLastRow(), 4);
  changetable.getRange(1, 1, invoiceInfo.length, 4).setValues(invoiceInfo);
  
  //get the rest of the sheets info
  var invoiceData = invoiceLog.getSheetValues(1, 1, invoiceLog.getLastRow(), invoiceLog.getLastColumn());
  var prop = PropertiesService.getScriptProperties();
  var valueArr = [];
  
  var billColumns = prop.getProperty("Billable Column Number").split(',')
  billColumns.pop();
  var rejectionColumns = prop.getProperty("Rejection Column Number").split(',')
  rejectionColumns.pop();
  Logger.log(billColumns);
  
  for(var row = 0; row < invoiceData.length; row++){
    valueArr.push([]);   
    for(var r = 0; r < rejectionColumns.length; r++){
      valueArr[row].push(invoiceData[row][rejectionColumns[r]],'','');
    }
    for(var b = 0; b < billColumns.length; b++){
      Logger.log(billColumns[b]);
      valueArr[row].push(invoiceData[row][billColumns[b]],'','');
    }
  }
  
  changetable.getRange(1, 5, valueArr.length, valueArr[0].length).setValues(valueArr);
}