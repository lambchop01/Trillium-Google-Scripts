function onOpen() {
  var who = Session.getActiveUser()
  if (who == "trilliumlamb@gmail.com"){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invoicesheet = ss.getSheetByName('Invoice Generator');
    var tagsheet = ss.getSheetByName('Tag Numbers')
    var companimalsheet = ss.getSheetByName("Compromised Animal Form")
  
    invoicesheet.getRange('B20').setValue('Save Invoice?');
    invoicesheet.getRange('B19').setValue('Create New Invoice');
    tagsheet.getRange('B6').setValue('Save Tags?');
    companimalsheet.getRange("B4").setValue("Save Form?")
  }//end if
}
