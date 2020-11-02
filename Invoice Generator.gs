function runsaveinvoice() {
  var writecell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Generator').getRange("B20")
  var cell = writecell.getValue();
  if(cell == "Save and Email"){
       writecell.setValue('Saving...')
       Code.generateinvoice();
       writecell.setValue('Check your email');
  };
}

function runsavecompanform() {
  var writecell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Compromised Animal Form').getRange("B4")
  var cell = writecell.getValue();
  if(cell == "Save and Email"){
       writecell.setValue('Saving...')
       Code.companimal();
       writecell.setValue('Check your email');
    };
}