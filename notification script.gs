//V 0.2.2 current version works for multiple cells edited

function sendEmail(subject, message, list) {
  
  //retrieves emails from Email Distro sheet. stored in string, separated by commas, starts on second row to avoid "Email Address" label at the top
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Distro");
  var emailAddress = sheet.getRange(2,list,sheet.getLastRow()-1).getValues();

  //sends email with retrieved email addresses, the subject and message as passed by parameters, and from a noReply email address
  MailApp.sendEmail(emailAddress, subject, message, {noReply:true});
}


function billable(sheet, range, time, user) {
  
  //creates Array[][] of values for active range and its corrolating information 
  var rangeValues = range.getValues();
  var info = sheet.getRange(range.getRow(), 4,range.getNumRows(),3).getValues();   //grabs info on row for that work order, starting with column C. Should be FQNID, NFID, then Internal WO ID 
  //getRanges( starting row of edited cells, the column that starts info (FQNID), the number of rows your grabbing, the number of columns you need

  //variables  
  var milestone;
  var fqnid;
  var nfid;
  var wo;
  
  //for every column in the range
  var col;
  for (col = 0; col < range.getNumColumns(); col  ++) {
    
    //check if column is a milestone
    milestone = sheet.getRange(1, range.getColumn()+col).getValue();
    if( (milestone == "MS-1 Ready for Invoicing") || (milestone == "MS-2.a Ready for Invoicing") || (milestone == "MS-2 Ready for Invoicing") || (milestone == "MS-3 Ready for Invoicing") || (milestone == "MS-4 Ready for Invoicing")){
      
      //abbreviate milestone. Cuts everything past first space
      milestone = milestone.substring(0,milestone.indexOf(" "));
     
      //for every row in that column
      var row;
      for (row = 0; row < range.getNumRows(); row++){
        
        //check if cell is BILLABLE
        if(rangeValues[row][col] == "BILLABLE"){
          
          fqnid = info[row][0];
          nfid = info [row][1];
          wo = info[row][2];
          
          //sendEmail(subject, message)
          sendEmail((wo+" at "+milestone+" is now BILLABLE, "+fqnid), (fqnid+" in "+wo+" at "+milestone+" was changed to BILLABLE "+"\n\nNFID: "+nfid+"\n\nTIME: "+time+"\nUSER: "+user),2);
        }
      }
    }
  }
}

//future functions go here

function onEdit(e) {
  
  //variables
  var sheet = e.source.getActiveSheet();
  var rangelist = sheet.getActiveRangeList().getRanges();
  var time = new Date().toString();
  var user = e.user.getEmail();
  
  //iterate through the rangelist
  var i;
  for(i = 0; i < rangelist.length; i++){
    
    //Check for Billable, if so, call method
    if(rangelist[i].getValue() == "BILLABLE"){
      billable(sheet,rangelist[i], time, user);
    }
    
    //future functions get called here. Formate: If( [conditions] ) { functionname(parameters); }
    
  }
  
}