function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invoicesheet = ss.getSheetByName('Invoice Generator');
  var tagsheet = ss.getSheetByName('Tag Numbers')
  var companimalsheet = ss.getSheetByName("Compromised Animal Form")
  
  invoicesheet.getRange('B20').setValue('Save Invoice?');
  invoicesheet.getRange('B19').setValue('Create New Invoice');
  tagsheet.getRange('B6').setValue('Save Tags?');
  companimalsheet.getRange("B4").setValue("Save Form?")
  
}

function iPhone() {
   var home = SpreadsheetApp.getActive().getSheetByName("Home")
   var culls = SpreadsheetApp.getActive().getSheetByName("Culls")
   var infosheet = SpreadsheetApp.getActive().getSheetByName("Producer Information")
   var forecast = SpreadsheetApp.getActive().getSheetByName("Forecast")
   var lrf = SpreadsheetApp.getActive().getSheetByName("Long Range Forecast")
   var invoice = SpreadsheetApp.getActive().getSheetByName("Invoice Generator")
   var companimal = SpreadsheetApp.getActive().getSheetByName("Compromised Animal Form")
   var shiprec = SpreadsheetApp.getActive().getSheetByName("Shipping Records")
   var NMPdata = SpreadsheetApp.getActive().getSheetByName("NMP Health Data")
   var prostats = SpreadsheetApp.getActive().getSheetByName("Production Stats")
   var tags = SpreadsheetApp.getActive().getSheetByName("Tag Numbers")
  
  
}