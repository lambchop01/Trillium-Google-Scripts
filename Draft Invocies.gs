function runddraftemail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Send Producer Invoices");
  var run = sheet.getRange("B11").getValue()
  if (run == "yes"){
    sheet.getRange("B11").setValue("Running")
    Draftemail()
    sheet.getRange("B11").setValue("Created")
  }
}

function autorun() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Send Producer Invoices");
  var run = sheet.getRange("B10").setValue("NMP")
  Draftemail()
  //notify()
}

function Draftemail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Send Producer Invoices");
  var range = sheet.getRange("B1").getValue();
  var invoiceid = sheet.getRange(range).getValues();
  var Length = invoiceid.length;
  var invoice = [];
  
  for (var i = 0; i < Length; i++) {
       invoice[i] = DriveApp.getFileById(invoiceid[i]);
       //while(results.hasNext()) {
       //  invoice.push(results.next());
       //}
     };//endfor 
  
  var recipient = sheet.getRange("B5").getValue();
  var subject = sheet.getRange("B6").getValue();
  var body = sheet.getRange("B7").getValue();
  GmailApp.createDraft(recipient, subject, body, {
      attachments: invoice     
    });
}
