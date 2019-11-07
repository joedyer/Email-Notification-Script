// v7


//Triggered Functions
function edit(e){
  
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var user = e.user;
  var time = new Date();
  
  if(sheet.getName() == "Invoice Log"){
    var position = getChangeTablePosition(range.getColumn(), range.getLastColumn());
    if(position != null){
      
    
    }
  } 
}

function clearChangeTable(e){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("changetable");
  var sheetVals = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  
  for(var col = 4; col < sheetVals[0].length; col +=3){
    var arr = [];
    for(var row = 3; row < sheetVals.length; row++){
      if(sheetVals[row][col] != ''){
        if(sheetVals[row][col+1] != ''){
          arr.push([sheetVals[row][0], sheetVals[row][1], sheetVals[row][2], sheetVals[row][3], sheetVals[row][col+1], sheetVals[row][col+2]]);
        }
      }
    }
    
    if(arr.length > 0){
      var headArr = sheetVals[0][col].split(" ");
      var ms = headArr[0];
      
      if(headArr[1] == 'Ready'){
        sendEmail(arr, "ATL, "+arr.length+" FQNID(s) are now Billable at "+ms, ms, "BILLABLE");
      }else{
         sendEmail(arr, "ATL, "+arr.length+" FQNID(s) were rejected at "+ms, ms, "REJECTION");
      }
    }
  }
 
  setChangeTable();
}

function sendEmail(arr, subject, milestone, name){
  
  var tmpl = HtmlService.createTemplateFromFile('Email Template.html');
  tmpl.arr = arr;
  var body = tmpl.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: getRecipient(name, milestone),
    subject: subject,
    htmlbody: body
  });
}

function getRecipient(name, milestone){
  
  var recipient = "DG-LCS-VZOF-ATL-"+name;
  
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
  
  var prop = PropertiesService.getScriptProperties();
  var billCols = prop.getProperty('Billable Column Number').split(',');
  billCols.pop();
  var rejCols = prop.getProperty('Rejection Column Number').split(',');
  rejCols.pop();
  
  for(var i = 0; i < rejCols.length; i++){
    if(start <= rejCols[i] && end >= rejCols[i]){
      return ((i*3)+4);
    }
  }
  
  for(var x = 0; x < billCols.length; x++){
    if(start > billCols[x]){
      i++;
      if(x == billCols.length-1){
        return null;
      }
    }
  }
  return ((i*3)+4);
}

function test(){
  Logger.log(getChangeTablePosition(102, 105));
}