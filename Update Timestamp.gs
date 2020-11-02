function update() {
  var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Producer Information")
  var s = SpreadsheetApp.getActiveSheet();
  if( s.getName() == "Forecast" ) { //checks that we're on the correct sheet
   var r = s.getActiveCell();
   if( r.getColumn() == 3 ) { //checks the column
     var date = new Date();
     var user = Session.getActiveUser()
     infosheet.getRange("G52").setValue(user)
     infosheet.getRange("F52").setValue(date)
   }
   else if( r.getColumn() == 5 ) { //checks the column
     var date = new Date();
     var user = Session.getActiveUser()
     infosheet.getRange("G52").setValue(user)
     infosheet.getRange("F52").setValue(date)
   };
 };
  
  
}