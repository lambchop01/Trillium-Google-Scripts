function sheetcheck(){
  var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information');
  var check = infosheet.getRange('L37').getValue();
  
  updatereminder()
  sendMail()
  
  if (check == 'ok'){}
  else if (check == '#ERROR!') {
    var check = infosheet.getRange('L37').getValue();
    if (check == 'ok'){}
    else {
    var who = infosheet.getRange('B4').getValue();
    var email = 'trilliumlamb@gmail.com';
    var title = who+" Sheet Error";
    var message = who+"'s sheet has an error, the next shipment to "+check+" the master list" 
    GmailApp.sendEmail(email, title, message);
    Pushover(who,title,message)
  }
  }
  else {
    var who = infosheet.getRange('B4').getValue();
    var email = 'trilliumlamb@gmail.com';
    var title = who+" Sheet Error";
    var message = who+"'s sheet has an error, the next shipment to "+check+" the master list" 
    GmailApp.sendEmail(email, title, message);
    Pushover(who,title,message)
  }
}

function sendMail(){
 var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information');
 var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information');
 var data = sh.getRange('G38:H40').getValues();
  //var htmltable =[];

var TABLEFORMAT = 'cellspacing="2" cellpadding="1" dir="ltr" border="0" style="font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:0.5px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:left;text-decoration:none;font-style:normal;'
var htmltable = '<table ' + TABLEFORMAT +' ">';

for (row = 0; row<data.length; row++){

htmltable += '<tr>';

for (col = 0 ;col<data[row].length; col++){
  if (data[row][col] === "" || 0) {htmltable += '<td>' + '' + '</td>';} 
  else {htmltable += '<td>' + data[row][col] + '</td>';}
}

     htmltable += '</tr>';
}

     htmltable += '</table>';
     Logger.log(data);
     Logger.log(htmltable);
  var email = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('B14').getValue()
  GmailApp.sendEmail(email, 'Shipping Reminder','' ,{htmlBody: 'Next shipping dates:'+htmltable})
}

function updatereminder(){
var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information');
var updated = infosheet.getRange("E52").getValue()
  if (updated == "Email") {
    var who = infosheet.getRange('B14').getValue();
    var subject = "Update Your Forecast"
    var body = "You have not updated your forecast in 2 months, please take a look and make sure it is updated. Overwrite any 0 in column 'C' if nothing needs changing to reset the date.  Thanks!"
    GmailApp.sendEmail(who, subject, body)
  }
}