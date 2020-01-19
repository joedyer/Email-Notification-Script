//functions to 1) set up new triggers which should be an event based trigger and a time based trigger

function setTriggers(){
  //set new triggers
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("changeEvent").forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('edit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('setUp').timeBased().atHour(3).everyDays(1).create();
  ScriptApp.newTrigger("clearChangeTable").timeBased().everyHours(1).create();

}

function changeEvent(e){

  //updates milestone columns
  if(e.changeType != 'OTHER' && range.getSheet().getName() == 'Invoice Log'){
    if(e.changeType == "INSERT_COLUMN" || e.changeType == "REMOVE_COLUMN"){
      setColumnProperties()
    }
    else if(e.changeType == 'INSERT_ROW'){
      
      var range = SpreadsheetApp.getActive().getActiveRange();
      SpreadsheetApp.getActive().getSheetByName('changetable').insertRowBefore(range.getRow());
    }
    
    else if(e.changeType == 'REMOVE_ROW'){
      
      var range = SpreadsheetApp.getActive().getActiveRange();
      SpreadsheetApp.getActive().getSheetByName('changetable').deleteRow(range.getRow());
    }
  }
}

function deleteTriggers(){  
  //delete all previously set up triggers
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}