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
 
function deleteTriggers(){  
  //delete all previously set up triggers
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}
