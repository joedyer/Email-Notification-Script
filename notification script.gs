//V 0.4.1 current version works for multiple cells edited and if change comes from formula

function sendEmail(subject, message, list) {
  
  //retrieves emails from Email Distro sheet. stored in string, separated by commas, starts on second row to avoid "Email Address" label at the top
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Distro");
  var emailAddress = sheet.getRange(2,list,sheet.getLastRow()-1).getValues();
  
  
  if(emailAddress[0][0] != ""){
    //sends email with retrieved email addresses, the subject and message as passed by parameters, and from a noReply email address
    MailApp.sendEmail(emailAddress, subject, message, {noReply:true});
  }else{
    //current script administrator: Joshua Cason
    MailApp.sendEmail("joshua.cason@lambertcable.com", "Email sent from a blank list", "Script attempted to send email from blank list at column "+list, {noReply:true});
  }
}

function setCurrentMilestone(curcol, milestones){
  if(curcol <= milestones.MS1[0]){
    return milestones.MS1;
  }else if(curcol <= milestones.MS2[0]){
    return milestones.MS2;
  }else if(curcol <= milestones.MS2a[0]){
    return milestones.MS2a;
  }else if(curcol <= milestones.MS3[0]){
    return milestones.MS3;
  }else if(curcol <= milestones.MS4[0]){
    return milestones.MS4;
  }
}

//future functions go here

function onEdit(e) {
  
  //milestone update function to run during edit???? pending runtime
  
  //variables
  var sheet = e.source.getActiveSheet();
  var range = sheet.getActiveRange();
  var rangeValues = range.getDisplayValues();
  var time = new Date().toString();
  var user = e.user.getEmail();
  
  //MilesStones Macro
  var milestones = {
    
    // [column of milestone, name of milestone, column of email list
    MS1: [31, "Milestone 1", 3], //AE
    MS2a: [47, "Milestone 2.a", 4], //AU
    MS2: [65, "Milestone 2", 5], //BM
    MS3: [94, "Milestone 3", 6], //CM
    MS4: [106, "Milestone 4", 7] //DB
    //for Column lookup: https://www.vishalon.net/blog/excel-column-letter-to-number-quick-reference
    
  }
  
  //Any billable update will be in the invoice log sheet
  if(sheet.getName() == "Invoice Log"){
    
    //info for the changed range
    var info = sheet.getRange(range.getRow(), 4,range.getNumRows(),3).getValues();
    
    //set current milestone
    var curMS = setCurrentMilestone(range.getColumn(), milestones);
   
    
    //for every row in the range
    for(var row = 0; row < range.getNumRows(); row++){
      
      //for every column in that range
      for(var col = 0; col < range.getNumColumns(); col++){
        
        //check if milestone is billable, if so, send email
        var cellValue =sheet.getRange(range.getRow(),curMS[0]).getValue();
        if(cellValue != ""){
          
          var fqnid = info[row][0];
          var nfid = info [row][1];
          var wo = info[row][2];
          
          //sendEmail(subject, message)
          sendEmail(("Atlanta "+wo+" at "+curMS[1]+" has been rejected, "+fqnid), (fqnid+" in "+wo+" at "+curMS[1]+" has been Rejected on "+cellValue+"\n\nNFID: "+nfid+"\n\nTIME: "+time+"\nUSER: "+user),curMS[2]);
          
          //break the column loop
          col = range.getNumColumns()+1;
        }
      }
    }
    
    
    //future functions get called here. Formate: If( [conditions] ) { functionname(parameters); }
    
  }
  
}