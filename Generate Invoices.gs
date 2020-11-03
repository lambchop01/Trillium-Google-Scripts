function rungenerateinvoice() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trillium Fee Invoices");
  var SSMP = sheet.getRange("P10").getValue()
  var LM = sheet.getRange("P57").getValue()
  var nmp = sheet.getRange("N95").getValue()
  var kawartha = sheet.getRange("P149").getValue()
  if (SSMP == "yes"){
    sheet.getRange("P10").setValue("Running")
    ssmpinvoice()
    sheet.getRange("P10").setValue("Created")
  }
  if (LM == "yes"){
    sheet.getRange("P57").setValue("Running")
    lminvoice()
    sheet.getRange("P57").setValue("Created")
  }
  if (nmp == "Yes"){
    sheet.getRange("N95").setValue("Running")
    nmpinvoice()
    sheet.getRange("N95").setValue("Created")
  }
  if (kawartha == "yes"){
    sheet.getRange("P149").setValue("Running")
    kawarthainvoice()
    sheet.getRange("P149").setValue("Created")
  }
}

function ssmpinvoice() {
  var Spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId()
  var datevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trillium Fee Invoices').getRange('J9').getValue()
  var date = Utilities.formatDate(datevalue, "GMT-5", "MMM dd yyyy")
    
  var url = "https://docs.google.com/spreadsheets/d/"+Spreadsheetid+"/export"+
                                                        "?format=pdf&"+
                                                        "size=7&"+
                                                        "fzr=true&"+
                                                        "portrait=true&"+
                                                        "fitw=true&"+
                                                        "gridlines=false&"+
                                                        "printtitle=false&"+
                                                        "sheetnames=false&"+
                                                        "pagenum=UNDEFINED&"+
                                                        "gid=1975415255&"+
                                                        "range=SSMP&"+
                                                        "ir=false&"+
                                                        "ic=false&";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var Invoice = UrlFetchApp.fetch(url, params).getBlob().setName("SSMP Trillium Fee Invoice "+date+".pdf");
  
  var dir = DriveApp.getFolderById("1QFmrtV92LfkCCMLRE3n3n8I_JbX1kKp4");
  dir.createFile(Invoice);
  
  var invnum = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H9").getValue();
  var newnum = invnum + 1;
  SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H9").setValue(newnum);
  var line = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("P3:T3").getValues();
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").insertRowBefore(8);
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").getRange("A8:E8").setValues(line);
  
  
  var email = "wzamanif@gmail.com,lambspork@gmail.com"
  var body = "Here is the Invoice for the last month.  Thanks" 
  var subject = "Trillium Fee Invoice"
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[Invoice]     
    });
  
}

function lminvoice() {
  var Spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId()
  var datevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trillium Fee Invoices').getRange('J55').getValue()
  var date = Utilities.formatDate(datevalue, "GMT-5", "MMM dd yyyy")
    
  var url = "https://docs.google.com/spreadsheets/d/"+Spreadsheetid+"/export"+
                                                        "?format=pdf&"+
                                                        "size=7&"+
                                                        "fzr=true&"+
                                                        "portrait=true&"+
                                                        "fitw=true&"+
                                                        "gridlines=false&"+
                                                        "printtitle=false&"+
                                                        "sheetnames=false&"+
                                                        "pagenum=UNDEFINED&"+
                                                        "gid=1975415255&"+
                                                        "range=LM&"+
                                                        "ir=false&"+
                                                        "ic=false&";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var Invoice = UrlFetchApp.fetch(url, params).getBlob().setName("L&M Trillium Fee "+date+".pdf");
  
  var dir = DriveApp.getFolderById("1QFmrtV92LfkCCMLRE3n3n8I_JbX1kKp4");
  dir.createFile(Invoice);
  
  var invnum = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H55").getValue();
  var newnum = invnum + 1;
  SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H55").setValue(newnum);
  var line = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("P49:T49").getValues();
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").insertRowBefore(8);
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").getRange("A8:E8").setValues(line);
  
  
  var email = "luanalauri5@hotmail.com,lambspork@gmail.com"
  var body = "Here is the Invoice for the last month.  Thanks" 
  var subject = "Trillium Fee Invoice"
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[Invoice]     
    });
  
}

function nmpinvoice() {
  var Spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId()
  var datevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trillium Fee Invoices').getRange('J101').getValue()
  var date = Utilities.formatDate(datevalue, "GMT-5", "MMM dd yyyy")
    
  var url = "https://docs.google.com/spreadsheets/d/"+Spreadsheetid+"/export"+
                                                        "?format=pdf&"+
                                                        "size=7&"+
                                                        "fzr=true&"+
                                                        "portrait=true&"+
                                                        "fitw=true&"+
                                                        "gridlines=false&"+
                                                        "printtitle=false&"+
                                                        "sheetnames=false&"+
                                                        "pagenum=UNDEFINED&"+
                                                        "gid=1975415255&"+
                                                        "range=NMP&"+
                                                        "ir=false&"+
                                                        "ic=false&";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var Invoice = UrlFetchApp.fetch(url, params).getBlob().setName("NMP Trillium Fee "+date+".pdf");
  
  var dir = DriveApp.getFolderById("1QFmrtV92LfkCCMLRE3n3n8I_JbX1kKp4");
  dir.createFile(Invoice);
  
  var invnum = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H101").getValue();
  var newnum = invnum + 1;
  SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H101").setValue(newnum);
  var line = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("P95:T95").getValues();
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").insertRowBefore(8);
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").getRange("A8:E8").setValues(line);
  
  
  var email = "diana@ontariolamb.ca,lambspork@gmail.com"
  var body = "Here is last weeks Trillium fee.  Thanks" 
  var subject = "Trillium Fee Invoice"
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[Invoice]     
    });
  
}

function kawarthainvoice() {
  var Spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId()
  var datevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trillium Fee Invoices').getRange('J147').getValue()
  var date = Utilities.formatDate(datevalue, "GMT-5", "MMM dd yyyy")
    
  var url = "https://docs.google.com/spreadsheets/d/"+Spreadsheetid+"/export"+
                                                        "?format=pdf&"+
                                                        "size=7&"+
                                                        "fzr=true&"+
                                                        "portrait=true&"+
                                                        "fitw=true&"+
                                                        "gridlines=false&"+
                                                        "printtitle=false&"+
                                                        "sheetnames=false&"+
                                                        "pagenum=UNDEFINED&"+
                                                        "gid=1975415255&"+
                                                        "range=Kawartha&"+
                                                        "ir=false&"+
                                                        "ic=false&";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var Invoice = UrlFetchApp.fetch(url, params).getBlob().setName("Kawartha Meats Trillium Fee "+date+".pdf");
  
  var dir = DriveApp.getFolderById("1QFmrtV92LfkCCMLRE3n3n8I_JbX1kKp4");
  dir.createFile(Invoice);
  
  var invnum = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H147").getValue();
  var newnum = invnum + 1;
  SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("H147").setValue(newnum);
  var line = SpreadsheetApp.getActive().getSheetByName("Trillium Fee Invoices").getRange("P141:T141").getValues();
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").insertRowBefore(8);
  SpreadsheetApp.getActive().getSheetByName("Trillium Fees").getRange("A8:E8").setValues(line);
  
  
  var email = "kawarthameatsinc@outlook.com,lambspork@gmail.com"
  var body = "Here is the Invoice for the last month.  Thanks" 
  var subject = "Trillium Fee Invoice"
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[Invoice]     
    });
  
}