//V 0.2.4 current version works for multiple cells edited 

function sendEmail(subject, message, list) {
  
  //retrieves emails from Email Distro sheet. stored in string, separated by commas, starts on second row to avoid "Email Address" label at the top
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Distro");
  var emailAddress = sheet.getRange(2,list,sheet.getLastRow()-1).getValues();

  //sends email with retrieved email addresses, the subject and message as passed by parameters, and from a noReply email address
  MailApp.sendEmail(emailAddress, subject, message, {noReply:true});
}


function billable(sheet, rangeValues, time, user, rangeRow, rangeCol) {
  
  //variables  
  var numRows = rangeValues.length;
  var numCols = rangeValues[0].length;
  var milestone;
  var fqnid;
  var nfid;
  var wo;
  
  
  //creates Array[][] of values for active range and its corrolating information 
  var info = sheet.getRange(rangeRow, 4,numRows,3).getValues();   //grabs info on row for that work order, starting with column C. Should be FQNID, NFID, then Internal WO ID 
  var headers = sheet.getRange(1, rangeCol, 1, numCols).getValues();
  //getRanges( starting row of edited cells, the column that starts info (FQNID), the number of rows your grabbing, the number of columns you need

  
  //for every column in the range
  var col;
  for (col = 0; col < numCols; col  ++) {
    
    //check if column is a milestone. This shouldn't be a problem for now but I can make this so it only calls the sheet once
    milestone = headers[0][col];
    if( (milestone == "MS-1 Ready for Invoicing") || (milestone == "MS-2.a Ready for Invoicing") || (milestone == "MS-2 Ready for Invoicing") || (milestone == "MS-3 Ready for Invoicing") || (milestone == "MS-4 Ready for Invoicing")){
      
      //abbreviate milestone. Cuts everything past first space
      milestone = milestone.substring(0,milestone.indexOf(" "));
     
      //for every row in that column
      var row;
      for (row = 0; row < numRows; row++){
        
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
  var range = sheet.getActiveRange();
  var rangeValues = range.getValues();
  var time = new Date().toString();
  var user = e.user.getEmail();
  
  
  //Any billable update will be in the invoice log sheet
  if(sheet.getName() == "Invoice Log"){
    
    //updates should be rectangular, if not, not update for billable. iterate only through the top row to find billable
    var x;
    for(x = 0 ; x < rangeValues[0].length; x++) {
      
      //Check x for Billable, if so, call method
      if(rangeValues[0][x] == "BILLABLE"){
        billable(sheet,rangeValues, time, user, range.getRow(),range.getColumn());
      }
    }
    
    //future functions get called here. Formate: If( [conditions] ) { functionname(parameters); }
    
  }
  
}