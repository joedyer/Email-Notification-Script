//This file contains all functions that would change the structure of a table or of the properties. It should contain set up for 1) properties and 2) the changetable

function setColumnProperties(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Log");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var properties = PropertiesService.getScriptProperties();
  properties.deleteAllProperties();
  properties.setProperty('Info', '');
  properties.setProperty('Billable Column Number', '');
  properties.setProperty('Rejection Column Number', '');
  
  var lastCol = 0;

  for(var i = 0; i < headers.length; i++){
    if(headers[i] == 'FQNID'||headers[i] == 'Site NFID'||headers[i] == 'Internal Work Order ID'||headers[i] == 'Customer GDB Work Order ID'){
      
      var temp = properties.getProperty('Info');
      properties.setProperty('Info', temp+(i+1)+',');
      
    }
    else if(headers[i].slice(-19) == 'Ready for Invoicing'){
      //billable column
      
      var temp = properties.getProperty('Billable Column Number');
      properties.setProperty('Billable Column Number', temp+(i+1)+',');
      
      if(i > lastCol){ lastCol = i+1;}
      
    }
    else if(headers[i].slice(-22) == 'Invoice Rejection Date'){
      //rejection column
      
      var temp = properties.getProperty('Rejection Column Number');
      properties.setProperty('Rejection Column Number', temp+(i+1)+',');
      
      if(i > lastCol){ lastCol = i;}
      
    }
  }
  
  //trim the columns and set last column
  properties.setProperty('Last Column', lastCol+1);
  
  var temp = properties.getProperty('Info');
  properties.setProperty('Info', temp.slice(0,-1));
  
  temp = properties.getProperty('Billable Column Number');
  properties.setProperty('Billable Column Number', temp.slice(0,-1));
  
  temp = properties.getProperty('Rejection Column Number');
  properties.setProperty('Rejection Column Number', temp.slice(0,-1));
  
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
  
  var invoiceData = invoiceLog.getSheetValues(1, 1, invoiceLog.getLastRow(), invoiceLog.getLastColumn());
  var prop = PropertiesService.getScriptProperties();
  var valueArr = [];
  
  var infoColumns = prop.getProperty("Info").split(',');
  var billColumns = prop.getProperty("Billable Column Number").split(',');
  var rejectionColumns = prop.getProperty("Rejection Column Number").split(',');
  
  for(var row = 0; row < invoiceData.length; row++){
    valueArr.push([]); 
    for(var i = 0; i < infoColumns.length; i++){
      valueArr[row].push(invoiceData[row][infoColumns[i]-1]);
    }
    for(var r = 0; r < rejectionColumns.length; r++){
      if(row == 0){
        valueArr[row].push(invoiceData[row][rejectionColumns[r]-1],'User','Time');
      }else{
        valueArr[row].push(invoiceData[row][rejectionColumns[r]-1],'',''); 
      }
    }
    for(var b = 0; b < billColumns.length; b++){
     if(row == 0){
        valueArr[row].push(invoiceData[row][billColumns[b]-1],'User','Time');
      }else{
        valueArr[row].push(invoiceData[row][billColumns[b]-1],'',''); 
      }
    }
  }
  
  changetable.getRange(1, 1, valueArr.length, valueArr[0].length).setValues(valueArr);
}

function addRowToChangeTable(){
  
  var range = SpreadsheetApp.getActive().getActiveRange();
  
  if(range.getSheet().getName() == 'Invoice Log'){
    SpreadsheetApp.getActive().getSheetByName('changetable').insertRowBefore(range.getRow());
  }
  
}