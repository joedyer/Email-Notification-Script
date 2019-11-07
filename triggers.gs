function setTriggers(ss){
  //set new triggers
  ScriptApp.newTrigger("changeEvent").forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger("clearChangeTable").timeBased().everyHours(1).create();

}

function changeEvent(e){
  //Normal operation, a change is made to a value on the table
  if(e.changeType == 'EDIT'){
    edit(e);
  }
  //updates milestone columns
  else if(e.changeType == "INSERT_COLUMN" || e.changeType == "REMOVE_COLUMN"){
    setColumnProperties();
  }
  //update changetable

}