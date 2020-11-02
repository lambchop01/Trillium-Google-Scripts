function runSubmitTags() {
  var writecell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tag Numbers').getRange("B6")
  var cell = writecell.getValue();
  if(cell == "Submit Tags"){
       writecell.setValue('Saving...')
       Code.SubmitTags();
       writecell.setValue('All Done!');
  };
}
