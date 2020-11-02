function copySheets() {
  //Be very careful with this...  no easy way to switch them all back.
  
  //copies entire sheet
  var source = SpreadsheetApp.openById("1d0fqSKJ5Jtn0kT_kMW6cr3lCTCdo5Tyjc_Esy0Ffgbw"); // location of Sheet to copy
  var sheet = source.getSheetByName('Producer Information');    // sheet to copy
  var Sheetname = sheet.getName();  //name of sheet to copy
  
  
  //What to copy
  /*var sheet = "Invoice Generator";
  var range = "B106:F139";
  var copyv = source.getSheetByName(sheet).getRange(range).getValues();
  var copyf = source.getSheetByName(sheet).getRange(range).getFormulas();
  var copyborders = source.getSheetByName(sheet).getRange(range).getBorders();
  */
  
  //Id's of all forecast sheets excluding mine from Variable Backup sheet
  var destinationId = SpreadsheetApp.openById("1G3Ygmv8FxKYtyr39_pP9i1Cv0kGaAO8aTgvx5X_RH28").getSheetByName("Variables").getRange("M21:M41").getValues();
  var arrayLength = destinationId.length;
  
    //for ever sheet id do the following
    for (var i = 0; i < arrayLength; i++) {
         var destination = SpreadsheetApp.openById(destinationId[i]);
         
         //delete named sheet
         /*destination.getSheetByName("NMP Health Data").activate()
         destination.deleteActiveSheet()
         */
      
         //set destination sheet for changes
         var dessheet = destination.getSheetByName(Sheetname);
         
         //changes to copy
      var range = dessheet.getRange('E8');
        dessheet.protect().setDescription('premise').setUnprotectedRanges([range]);
         
         
         // copies named Sheet
         /*sheet.copyTo(destination).setName(Sheetname);
         var gid = destination.getSheetByName("NMP Health Data").getSheetId()
         dessheet.getRange("C16").setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/"&'+"'Producer Information'"+'!$A$41&"/edit?folder=1DoNZ66x6mEg_LCJ3V_MhwLpgv5mlZv9f#gid='+gid+'","'+Sheetname+'")')
         */
     }//endfor
  }//end function
