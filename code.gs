// v7.2
// Updates: Now takes info as script property and uses it in changetable setup. Adding columns in info range now doesn't require manual code adjustment
// Known restrictions: Cannot handle addition of a milestone. Adding and deleting many rows is hit or miss. Current fix is a full reset at 3 AM everyday to ensure quality. 


//Triggered Functions
function edit(e){
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var user = e.user;
  var time = new Date();
  
 
  if(editFilter(sheet, range)){
    var position = getChangeTablePosition(range.getColumn(), range.getLastColumn());
    if(position != null){
      var invoiceValues = sheet.getRange(range.getRow(), position[1], range.getNumRows()).getValues();
      var changeLogRange = SpreadsheetApp.getActive().getSheetByName('changetable').getRange(range.getRow(), position[0], range.getNumRows(), 3);
      var changeValues = changeLogRange.getValues();
      var changeFlag = false;

      for(var i = 0; i < invoiceValues.length; i++){
        if(invoiceValues[i][0] != changeValues[i][0]){
          changeValues[i][0] = invoiceValues[i][0];
          changeValues[i][1] = user;
          changeValues[i][2] = time;
          changeFlag = true;
        }
      }
      
      if(changeFlag){
        changeLogRange.setValues(changeValues);
      }
    }
  } 
}

function editFilter(sheet, range){
  
  var prop = PropertiesService.getScriptProperties();
  
  if(sheet.getName() != "Invoice Log"){
    //must be an edit in invoice log
    return false;
  }
  else if(range.getColumn() <= 7 && range.getLastColumn() >= 4){
    //edit is to info and just needs copying over
    updateChangeTableInfo(range);
    return false;
  }
  else if(range.getColumn() > prop.getProperty('Last Column')){
    //edit is outside of range that would affect billable or rejection
    return false;
  }
  else if(range.getBackground() == '#f4cccc'){
    //background is changed light red, so its a cancelled FQNID
    return false;
  }
  else{
    return true;
  }
}

function clearChangeTable(){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("changetable");
  var sheetRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var sheetVals = sheetRange.getValues();
  var toMailArr = [];
  
  for(var col = 4; col < sheetVals[0].length; col +=3){
    //col starts at rejection at ms 1 then increments by 3 to get to next column
    //each column gets an array. If that array isn't empty by the end then headers are unshifted into it and it gets put in the toMailArr (to mail out array) 
    var colArr = [];
    for(var row = 3; row < sheetVals.length; row++){
      //row starts below headers and increments by 1
      if(sheetVals[row][col+1] != ''){
        //check for change with user columns
        colArr.push([sheetVals[row][0], sheetVals[row][1], sheetVals[row][2], sheetVals[row][3], sheetVals[row][col+1], sheetVals[row][col+2]]);
        sheetVals[row][col+1] = "";
        sheetVals[row][col+2] = "";
      }
    }
    
    if(colArr.length > 0){
      colArr.unshift(sheetVals[0][col]);             //unshift header
      toMailArr.push(colArr);                        //push column array into To mail out Array
    }
  }
  
  sheetRange.setValues(sheetVals);
  
  for(var i = 0; i < toMailArr.length; i++){
    sendEmail(toMailArr[i]);
  }
  
}

function sendEmail(arr){
  var name = '';
  var subject = '';
  var header = arr.shift();
  
  var ms = header.slice(0,header.indexOf(' '));
  header = header.slice(header.indexOf(' ')+1);
  
  if(header == 'Invoice Rejection Date'){
    subject = 'GRNB, '+arr.length+' FQNID(s) were rejected at '+ms;
    name = 'REJECTION';
  }else if(header == 'Ready for Invoicing'){
    subject = 'GRNB, '+arr.length+' FQNID(s) are now billable at '+ms;
    name = 'BILLABLE';
  }
  
  var tmpl = HtmlService.createTemplateFromFile('Email Template.html');
  tmpl.arr = arr;
  var body = tmpl.evaluate().getContent();
   
  MailApp.sendEmail({
    to: getRecipient(name, ms),
    name: name,
    subject: subject,
    htmlBody: body
  });
}

function getRecipient(name, milestone){
  
  var recipient = "DG-LCS-VZOF-GRNB-"+name;
  
  //for rejections, which MS
  if(name == "REJECTION"){
    if(milestone == "MS-1"){
      recipient +="-MS1";
    }else{
      recipient += "-MS2";
    }
  }
  
  recipient += "@lambertcable.com";

  return recipient;
}

function getChangeTablePosition(start, end){
  //returns an array with the position of column in the changetable as the first element and the position of the column in the invoice log as the second element. IE. [26, 59]
  
  var prop = PropertiesService.getScriptProperties();
  var billCols = prop.getProperty('Billable Column Number').split(',');
  var rejCols = prop.getProperty('Rejection Column Number').split(',');
  
  
  for(var i = 0; i < rejCols.length; i++){
    if(start <= rejCols[i] && end >= rejCols[i]){
      return [((i*3)+5), rejCols[i]];
      break;
    }
  }
  
  if(start <= billCols[0]){
    return [((i*3)+5), billCols[0]];
  }
  i++;
  
  for(var j = 1; j < billCols.length; i++, j++){
    if(start > billCols[j-1] && start <= billCols[j]){
      return [((i*3)+5), billCols[j]];
      break;
    }
  }
  
  return null;
}

function updateChangeTableInfo(range){
  SpreadsheetApp.getActive().getSheetByName('changetable').getRange(range.getRow(), range.getColumn()-3, range.getNumRows(), range.getNumColumns()).setValues(range.getValues());

}// v7.2
// Updates: Now takes info as script property and uses it in changetable setup. Adding columns in info range now doesn't require manual code adjustment
// Known restrictions: Cannot handle addition of a milestone. Adding and deleting many rows is hit or miss. Current fix is a full reset at 3 AM everyday to ensure quality. 


//Triggered Functions
function edit(e){
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var user = e.user;
  var time = new Date();
  
 
  if(editFilter(sheet, range)){
    var position = getChangeTablePosition(range.getColumn(), range.getLastColumn());
    if(position != null){
      var invoiceValues = sheet.getRange(range.getRow(), position[1], range.getNumRows()).getValues();
      var changeLogRange = SpreadsheetApp.getActive().getSheetByName('changetable').getRange(range.getRow(), position[0], range.getNumRows(), 3);
      var changeValues = changeLogRange.getValues();
      var changeFlag = false;

      for(var i = 0; i < invoiceValues.length; i++){
        if(invoiceValues[i][0] != changeValues[i][0]){
          changeValues[i][0] = invoiceValues[i][0];
          changeValues[i][1] = user;
          changeValues[i][2] = time;
          changeFlag = true;
        }
      }
      
      if(changeFlag){
        changeLogRange.setValues(changeValues);
      }
    }
  } 
}

function editFilter(sheet, range){
  
  var prop = PropertiesService.getScriptProperties();
  
  if(sheet.getName() != "Invoice Log"){
    //must be an edit in invoice log
    return false;
  }
  else if(range.getColumn() <= 7 && range.getLastColumn() >= 4){
    //edit is to info and just needs copying over
    updateChangeTableInfo(range);
    return false;
  }
  else if(range.getColumn() > prop.getProperty('Last Column')){
    //edit is outside of range that would affect billable or rejection
    return false;
  }
  else if(range.getBackground() == '#f4cccc'){
    //background is changed light red, so its a cancelled FQNID
    return false;
  }
  else{
    return true;
  }
}

function clearChangeTable(){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("changetable");
  var sheetRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var sheetVals = sheetRange.getValues();
  var toMailArr = [];
  
  for(var col = 4; col < sheetVals[0].length; col +=3){
    //col starts at rejection at ms 1 then increments by 3 to get to next column
    //each column gets an array. If that array isn't empty by the end then headers are unshifted into it and it gets put in the toMailArr (to mail out array) 
    var colArr = [];
    for(var row = 3; row < sheetVals.length; row++){
      //row starts below headers and increments by 1
      if(sheetVals[row][col+1] != ''){
        //check for change with user columns
        colArr.push([sheetVals[row][0], sheetVals[row][1], sheetVals[row][2], sheetVals[row][3], sheetVals[row][col+1], sheetVals[row][col+2]]);
        sheetVals[row][col+1] = "";
        sheetVals[row][col+2] = "";
      }
    }
    
    if(colArr.length > 0){
      colArr.unshift(sheetVals[0][col]);             //unshift header
      toMailArr.push(colArr);                        //push column array into To mail out Array
    }
  }
  
  sheetRange.setValues(sheetVals);
  
  for(var i = 0; i < toMailArr.length; i++){
    sendEmail(toMailArr[i]);
  }
  
}

function sendEmail(arr){
  var name = '';
  var subject = '';
  var header = arr.shift();
  
  var ms = header.slice(0,header.indexOf(' '));
  header = header.slice(header.indexOf(' ')+1);
  
  if(header == 'Invoice Rejection Date'){
    subject = 'GRNB, '+arr.length+' FQNID(s) were rejected at '+ms;
    name = 'REJECTION';
  }else if(header == 'Ready for Invoicing'){
    subject = 'GRNB, '+arr.length+' FQNID(s) are now billable at '+ms;
    name = 'BILLABLE';
  }
  
  var tmpl = HtmlService.createTemplateFromFile('Email Template.html');
  tmpl.arr = arr;
  var body = tmpl.evaluate().getContent();
   
  MailApp.sendEmail({
    to: getRecipient(name, ms),
    name: name,
    subject: subject,
    htmlBody: body
  });
}

function getRecipient(name, milestone){
  
  var recipient = "DG-LCS-VZOF-GRNB-"+name;
  
  //for rejections, which MS
  if(name == "REJECTION"){
    if(milestone == "MS-1"){
      recipient +="-MS1";
    }else{
      recipient += "-MS2";
    }
  }
  
  recipient += "@lambertcable.com";

  return recipient;
}

function getChangeTablePosition(start, end){
  //returns an array with the position of column in the changetable as the first element and the position of the column in the invoice log as the second element. IE. [26, 59]
  
  var prop = PropertiesService.getScriptProperties();
  var billCols = prop.getProperty('Billable Column Number').split(',');
  var rejCols = prop.getProperty('Rejection Column Number').split(',');
  
  
  for(var i = 0; i < rejCols.length; i++){
    if(start <= rejCols[i] && end >= rejCols[i]){
      return [((i*3)+5), rejCols[i]];
      break;
    }
  }
  
  if(start <= billCols[0]){
    return [((i*3)+5), billCols[0]];
  }
  i++;
  
  for(var j = 1; j < billCols.length; i++, j++){
    if(start > billCols[j-1] && start <= billCols[j]){
      return [((i*3)+5), billCols[j]];
      break;
    }
  }
  
  return null;
}

function updateChangeTableInfo(range){
  SpreadsheetApp.getActive().getSheetByName('changetable').getRange(range.getRow(), range.getColumn()-3, range.getNumRows(), range.getNumColumns()).setValues(range.getValues());

}