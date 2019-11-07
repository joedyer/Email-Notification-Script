//run to put in place a new iteration of the code
//this file will delete any 1) old properties and 2) triggers and run the functions to create and set up 3) the new properties and 4) the triggers and 5) the changetable
function setUp(){
  
  var ss = SpreadsheetApp.getActive();
  
  //clear any old triggers and properties
  PropertiesService.getScriptProperties().deleteAllProperties();
  deleteTriggers();
 
  //set milestones, create the changeTable
  setColumnProperties();
  setChangeTable();
  setTriggers(ss);
}