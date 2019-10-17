// v6.0.1 version updates: includes hidden changetable, will send emails based of time trigger (i.e. hourly), all setup (triggers, sheets, milestones, etc) contained within setUp()

function setUp(){
  
  var ss = SpreadsheetApp.getActive();
  
  //delete all previously set up triggers
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  
  //delete any previous properties and add market name
  PropertiesService.getScriptProperties().setProperties({Hub: ss.getName().split(" ")[0]}, true);
  
  //set milestone properties
  setColumnProperties();
  
  //add or update the changelog
  setChangeTable();
  
  //set new triggers
  ScriptApp.newTrigger("onAddColumns").forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger("clearChangeTable").timeBased().everyMinutes(10).create();
}

//Triggered Functions

function onAddColumns(e){
  //updates milestone columns whenever they change
  if(e.changeType == "INSERT_COLUMN" || e.changeType == "REMOVE_COLUMN"){
    setColumnProperties();
  }
}

function onEdit(e){
  
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var user = e.user;
  var time = new Date();
  
  if(sheet.getName() == "Invoice Log"){
  
    var properties = PropertiesService.getScriptProperties().getProperties();
    
    //which milestone are you working with
    var rangeCol = range.getColumn();
    var column = 100000;
    var name = "";
    for (var x in properties) {
      if(x.charAt(0) == "M"){
        var num = parseInt(properties[x]);
        if(num >= rangeCol && num < column){
          column = num;
          name = x;
        }
      }
    }
    
    if(name != ""){
      
      var rangeRow = range.getRow();
      var rangeNumRows = range.getNumRows();
      
      if(properties[("changetable "+name)]==undefined){
        addChangeTableColumn("changetable "+name); 
      }
      
      var changelogRange = e.source.getSheetByName("changetable").getRange(rangeRow, properties[("changetable "+name)], rangeNumRows, 3);
      var changelogRangeValues = changelogRange.getValues();
      var invoiceRangeValues = sheet.getRange(rangeRow, column, rangeNumRows).getValues();
      var arr = [];
      
      for(var i = 0; i < rangeNumRows; i++){
        if(invoiceRangeValues[i][0] != ""){
          if(changelogRangeValues[i][0] != ""){
            arr.push(changelogRangeValues[i]); 
          }else{
            arr.push([invoiceRangeValues[i][0],user,time]);
          }
        }else{
          arr.push(["","",""]);
        }
      }
      
      changelogRange.setValues(arr);
      
    }
  } 
}

function clearChangeTable(e){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("changetable");
  var sheetVals = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var properties = PropertiesService.getScriptProperties().getProperties();
  
  var billArr = [];
  var rejArr = [];
  
  for(var x in properties){
    var propArr = x.split(" ");
    if(propArr[0] == "changetable"){
      for(var i = 1; i < sheetVals.length; i++){
        var col = parseInt(properties[x]);
        if(sheetVals[i][col] != ""){
          //add info, the user, and the time
          var insertArr = sheetVals[i].slice(0,4);
          insertArr.push(propArr[1],sheetVals[i][col],sheetVals[i][col+1]);
          
          //insert array 
          if(propArr[2] == "billable"){
            billArr.push(insertArr);
          } else if(propArr[2] == "rejection"){
            rejArr.push(insertArr);
          }
          //clear any data as you move through the sheet
          sheetVals[i][col] = "";
          sheetVals[i][col+1] = "";
        }
      }
    }
  }
  
  //sheetvals now contains all the info it used too, but without the users and the time. BillArr and rejArr contain all the changes
  //reset the changetable
  sheet.getRange(1, 1, sheetVals.length, sheetVals[0].length).setValues(sheetVals);
  
  if(billArr.length > 0){
    sendEmail(properties["Hub"],billArr, "are now Billable","These fibers are now Billable:","BILLABLE");
  }
  if(rejArr.length > 0){
    sendEmail(properties["Hub"],rejArr, "were rejected","These fibers were rejected:","REJECTIONS");
  }
  
}

function sendEmail(hub, arr, subject, message, name){
  
  var tableHeaders = "<table style='border: 1px solid black;'><tr><th>FQNID</th><th>NFID</th><th>Internal WO</th><th>GDB WO</th><th>Milestone</th><th>User</th><th>Time</th></tr>";
  
  var subjectArr = [];
  var messageArr = [];
  var fiberCountArr = [];
  
  var fibers  = 0;
  var currMS = "";
  
  //construct an array of html email bodies
  for(var i = 0; i < arr.length; i++){
    
    if(currMS != arr[i][4]){
      currMS = arr[i][4]
      messageArr.push(message+"\n</p><p>At <b>"+currMS+"</b>:\n"+tableHeaders);
      //add that milestone to subject so we can build the entire string later
      subjectArr.push(currMS);
      if(fibers > 0){
        //this check will only be hit when the loop runs the first time
        fiberCountArr.push(fibers);
        fibers = 0;
        messageArr[messageArr.length-1] += "</table>";
      }
    }
    
    var tablerow = "<tr>";
    for(var j=0; j < arr[i].length; j++){
      tablerow += "<td style='padding: 15px;border: 1px solid black;'>"+arr[i][j]+"</td>";
    }
    tablerow+= "</tr>";
    messageArr[messageArr.length-1] += tablerow;
    fibers++;
    
  }
  //push the fiber count for the last milestone
  fiberCountArr.push(fibers);
  
  //fill in the rest of the subject line
  for(var i = 0; i < subjectArr.length; i++){
    
    subjectArr[i] = hub+", "+fiberCountArr[i]+" fibers "+subject+" at "+subjectArr[i];
    
  }
  
  //send the emails out
  for(var i = 0; i < subjectArr.length; i++){
    
    MailApp.sendEmail({
      to: "joseph.dyer@engineeringassociates.com",
      subject: subjectArr[i],
      htmlBody: messageArr[i],
      name:name
    });
    
  }
}

//Setters and on setup functions

function setColumnProperties(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Log");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var properties = PropertiesService.getScriptProperties();
  
  for(var i = 0; i < headers.length; i++){
    
    if(headers[i] == "FQNID"){
      properties.setProperty("FQNID", (i+1));
    }else if(headers[i] == "Site NFID"){
      properties.setProperty("NFID", (i+1));
    }else if(headers[i] == "Internal Work Order ID"){
      properties.setProperty("Internal WO", (i+1));
    }else if(headers[i] == "Customer GDB Work Order ID"){
      properties.setProperty("GDB WO", (i+1));
    }else{
      //for milestones
      var arr = headers[i].split(" ");
      if(arr[1]=="Ready" && arr[3]=="Invoicing"){
        properties.setProperty((arr[0]+" billable"), (i+1)); 
      }else if(arr[2] == "Rejection" && arr[3] == "Date"){
        properties.setProperty((arr[0]+" rejection"), (i+1)); 
      }
    }
  }
}

function setChangeTable(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceLog = ss.getSheetByName("Invoice Log");
  var invoiceLogValues = invoiceLog.getSheetValues(1, 1, invoiceLog.getLastRow(), invoiceLog.getLastColumn());
  var numRows = invoiceLogValues.length;
  
  var arr = [];
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var properties = scriptProperties.getProperties();
  
  var fqnid = properties["FQNID"];
  var GDBWO = properties["GDB WO"];
  
  //add info to changetable first
  for(var i = 0; i < numRows; i++){
    arr.push(invoiceLogValues[i].slice(fqnid-1, GDBWO));
  }
  
  for(var x in properties){
    if(x.charAt(0)=="M"){
      //for every property that's a milestone, billable or rejection
      for(var i =0; i < numRows; i++){
        var col = parseInt(properties[x])-1;
        if(i == 0){
          //add the header
          arr[0].push(x,"User","Time");
          scriptProperties.setProperty("changetable "+x, (arr[0].length-2));
        }else{
          //copy value and leave blanks for user and time
          arr[i].push(invoiceLogValues[i][col], "","");
        }
      }
    }
  } 
  
  var changetable = ss.getSheetByName("changetable");
  
  //if the change table doesn't exist yet, then add it
  if(changetable == null){
    ss.insertSheet("changetable");
    changetable = ss.getSheetByName("changetable");
    changetable.hideSheet();
  }
  
  //clearing any previous data
  changetable.clear();
  
  changetable.getRange(1, 1, arr.length, arr[0].length).setValues(arr);
}

function addChangeTableColumn(name){
  var changelog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("changetable");
  var column = changelog.getLastColumn()+1;
  PropertiesService.getScriptProperties().setProperty(name, column);
  changelog.getRange(1, column,1,3).setValues([[name,"User","Time"]]);
}
