function runinvoicegenerator() {
  var writecell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Blank Pay Sheet').getRange("A1")
  var cell = writecell.getValue();
  if(cell == "Save and Send"){
       writecell.setValue('Saving...')
       generateinvoice();
       writecell.setValue('Check your email');
  };
}

function generateinvoice() {
  var invoicesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Blank Pay Sheet');
    
  
  var Spreadsheetid = SpreadsheetApp.getActiveSpreadsheet().getId()
  var datevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Blank Pay Sheet').getRange('A2').getValue()
  var date = Utilities.formatDate(datevalue, "GMT-5", "MMM dd yyyy")
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
  
  var Invoice = UrlFetchApp.fetch(url, params).getBlob().setName("Kev Pay Invoice "+date+".pdf");
  
  
  
  // save to drive, specify location
  var driveid = "1IgyNcwYE8KHISDiTiWqdGVpWuAcxsSpa"
  var dir = DriveApp.getFolderById(driveid);
  var idinvoice = dir.createFile(Invoice).getId();
    
  //or send as email
  
  var email = "board@trilliumlamb.ca";
  var body = "Here is the latest. Thanks"; 
  var subject = "Invoice "+date;
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[Invoice]     
    });
  
} 