function runinvoicegenerator() {
  var writecell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Generator').getRange("B20")
  var cell = writecell.getValue();
  if(cell == "Save and Email"){
       writecell.setValue('Saving...')
       generateinvoice();
       writecell.setValue('Check your email');
  };
}

function generateinvoice() {
  var invoicesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Generator');
  var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information');
  var farm = infosheet.getRange("B8").getValue()
  var edit = invoicesheet.getRange('B19').getValue();
  
  if (edit > 0){
    var Invoicenum = invoicesheet.getRange('B19').getValue();
    var invoicecount = invoicesheet.getRange('B4').getValue();
    invoicesheet.getRange('B4').setValue(Invoicenum+"a");
    var invid = infosheet.getRange('B56').getValue();
    var manid = infosheet.getRange('B57').getValue();
    DriveApp.getFileById(invid).setTrashed(true);
    DriveApp.getFileById(manid).setTrashed(true);
  }
  
  
  var Spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId()
  var datevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Generator').getRange('B3').getValue()
  var date = Utilities.formatDate(datevalue, "GMT-5", "MMM dd yyyy")
  var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information')
  var packer = invoicesheet.getRange('B5').getValue()
  var invoicesheetid = invoicesheet.getSheetId()
    
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
                                                        "gid="+invoicesheetid+"&"+
                                                        "range=Invoice&"+
                                                        "ir=false&"+
                                                        "ic=false&";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var Invoice = UrlFetchApp.fetch(url, params).getBlob().setName(farm+" "+packer+" Invoice "+date+".pdf");
  
  
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
                                                        "attachment=true&"+
                                                        "gid="+invoicesheetid+"&"+
                                                        "range=Manifest&"+
                                                        "ir=false&"+
                                                        "ic=false&";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var Manifest = UrlFetchApp.fetch(url, params).getBlob().setName(farm+" "+packer+" Manifest "+date+".pdf");
  
  
  // save to drive, specify location
  var driveid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('B41').getValue()
  var dir = DriveApp.getFolderById(driveid);
  var idinvoice = dir.createFile(Invoice).getId();
  invoicesheet.getRange('AG16').setValue(idinvoice);
  var dir2 = DriveApp.getFolderById("1AUTovOVP1M6La5UT_-Fwhi2jVO5E68Gg")
  var idinvoice2 = dir2.createFile(Invoice).getId()
  invoicesheet.getRange('AI16').setValue(idinvoice2)
  var idmanifest = dir.createFile(Manifest).getId();
  invoicesheet.getRange('AH16').setValue(idmanifest);
  DriveApp.getFolderById(driveid)
  
  //or send as email
  var email = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('B14').getValue()
  var body = "Here you go!" 
  var subject = packer+" Invoice and Manifest "+date
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[Invoice, Manifest]     
    });
  
  /*
  //attach to group email
  var gmaildraftid = infosheet.getRange("B59").getValue();
  var draftto = infosheet.getRange("B60").getValue();
  var draftsub = infosheet.getRange("B62").getValue()+" "+date;
  var draftbody = infosheet.getRange("B61").getValue();
  var idrow = infosheet.getRange("B63").getValue();
  if (gmaildraftid == ""){
    var gmaildraftid = GmailApp.createDraft(draftto, draftsub, draftbody, {
      htmlBody: body,
      attachments:[Invoice]     
    }).getId();
    SpreadsheetApp.openById("1G3Ygmv8FxKYtyr39_pP9i1Cv0kGaAO8aTgvx5X_RH28").getSheetByName("Code Variables").getRange(idrow).setValue(gmaildraftid);
  }//end if
  else {
    GmailApp.getDraft(gmaildraftid).update(draftto, draftsub, draftbody, {
      htmlBody: body,
      attachments:[Invoice]     
    });
  };//end else
  */
  
  // save invoice information to the shipping records sheet
  var copy = infosheet.getRange('B52').getValue();
  var records = invoicesheet.getRange(copy).getValues();
  var addrow = infosheet.getRange('B54').getValue();
  var paste = infosheet.getRange('B53').getValue();
  
  if (edit == "Create New Invoice"){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shipping Records').insertRowAfter(addrow);
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shipping Records').getRange(paste).setValues(records);
  invoicesheet.getRange('B19').setValue("Create New Invoice")
  
  //updates invoice number each time script is run
  if (edit == "Create New Invoice"){
  var Invoicenum = invoicesheet.getRange('B4').getValue();
  invoicesheet.getRange('B4').setValue(Invoicenum+1);
  }
  else {
    invoicesheet.getRange('B4').setValue(invoicecount)
  }
}

function runcompanimal() {
  var writecell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Compromised Animal Form').getRange("B4")
  var cell = writecell.getValue();
  if(cell == "Save and Email"){
       writecell.setValue('Saving...')
       companimal();
       writecell.setValue('Check your email');
    };
}

function companimal() {
  var companimalsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Compromised Animal Form');
  var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information');
  var farm = infosheet.getRange("B8").getValue()
  var invoicesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Generator');
  
  var Spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId()
  var datevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Generator').getRange('B3').getValue()
  var date = Utilities.formatDate(datevalue, "GMT-5", "MMM dd yyyy")
  var packer = invoicesheet.getRange('B5').getValue()
  var companimalsheetid = companimalsheet.getSheetId()
    
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
                                                        "gid="+companimalsheetid+"&"+
                                                        "range=Companimalform&"+
                                                        "ir=false&"+
                                                        "ic=false&";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var companimalform = UrlFetchApp.fetch(url, params).getBlob().setName(farm+" "+packer+" Compromised Animal Form "+date+".pdf");
  
  // save to drive, specify location
  var driveid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('B41').getValue()
  var dir = DriveApp.getFolderById(driveid);
  dir.createFile(companimalform)
  
  
  
  //or send as email
  var email = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Producer Information').getRange('B14').getValue()
  var body = "Here you go!" 
  var subject = packer+" Compromised Animal Form "+date
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[companimalform]     
  });
  
  
}