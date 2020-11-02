function runSubmitTags() {
  var writecell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tag Numbers').getRange("B6")
  var cell = writecell.getValue();
  if(cell == "Submit Tags"){
       writecell.setValue('Saving...')
       SubmitTags();
       writecell.setValue('All Done!');
  };
  
}

function SubmitTags() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tag Numbers').getRange('K2').getValue();
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tag Numbers').getRange(range).getValues();
  var pasterange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tag Numbers').getRange('K3').getValue();
  var numtags = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tag Numbers').getRange('B8').getValue();
  var desid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('A43').getValue();
  SpreadsheetApp.openById(desid).getSheetByName('Tags').insertRowsAfter(1, numtags);
  SpreadsheetApp.openById(desid).getSheetByName('Tags').getRange(pasterange).setValues(data);
  
  var who = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('B4').getValue();
  var message = who+' just submitted tag numbers';
  var title = 'Tags Submitted'

  Pushover(who,message,title);
  
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tag Numbers').getRange('A10:A').clear();
  
  
}

function Pushover(who,message,title) {
/*
  var who = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('B4').getValue();
  var message = who+' just submitted tag numbers';
  var title = 'Tags Submitted'
*/  
//setup the Pushover API call
var baseUrl = "https://api.pushover.net/1/messages.json";


var parameters = {
  'token':'axdwgawh5kntppfyjquybc7pex2xas',
  'user':'uun2z3tqwnpnu5d7yx4pnfy4vv9to2',
  'title':title,
  'message':message};

  
var options = {
  'method':'POST',
  'payload':parameters
};

UrlFetchApp.fetch(baseUrl, options)
  
}