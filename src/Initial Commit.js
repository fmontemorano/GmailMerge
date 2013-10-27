
function PrivateCache(){
  
  //create and cache a spreadsheet  
  CacheService.getPrivateCache().remove('cachedObject');
  var Spreadsheet = SpreadsheetApp.create(Utilities.formatDate(new Date(), Session.getTimeZone(), "yyyy MM dd | HH:mm")).getId();
  CacheService.getPrivateCache().put('cachedObject', Spreadsheet); 
}

//////////////////////////////////////////////////////////////////////////////ℱℳ
function CreateSpreadsheet(){
  
  //locate spreadsheet
  PrivateCache();
  var cachedObject = CacheService.getPrivateCache().get('cachedObject');
  var Spreadsheet = SpreadsheetApp.openById(cachedObject);
  
  //resize the spreadsheet
  var lastColumn = Spreadsheet.getLastColumn();
  if(lastColumn>6){
    Spreadsheet.insertColumnsAfter(1,1);
    Spreadsheet.deleteColumns(1,lastColumn);
  }
  
  var lastRow = Spreadsheet.getLastRow();
  if(lastRow>100){
    Spreadsheet.insertRowsAfter(1,1);
    Spreadsheet.deleteRows(1,lastRow);
  }
  
  Spreadsheet.getActiveSheet().clear()
  
  //fill in spreadsheet
  Spreadsheet.getRange("A1").setValue("Name");
  Spreadsheet.getRange("B1").setValue("Address ");
  Spreadsheet.getRange("C1").setValue('=hyperlink("http://knopok.net/text-to-html.html","Body")');
  Spreadsheet.getRange("D1").setValue("Subject");
  Spreadsheet.getRange("E1").setValue("Attachment");
  Spreadsheet.getRange("F1").setValue("Status");
  
  // format spreadsheet
  Spreadsheet.getRange("A:F").setFontColor("Black");
  Spreadsheet.getRange("A:F").setFontFamily("Tahoma");
  Spreadsheet.getRange("A:F").setBackgroundColor("white");
  Spreadsheet.getRange("A:F").setFontSize(10);
  Spreadsheet.getRange("A1:F1").setBackgroundColor("#b2b2b2");
  Spreadsheet.getActiveSheet().setFrozenRows(1);
  Spreadsheet.getActiveSheet().getDataRange();
  var quota = MailApp.getRemainingDailyQuota(); 
  Spreadsheet.getRange("G1").setValue("you can still send " + quota + " messages today before reaching quota.");
  Spreadsheet.getRange("H1").setValue("[group:]");
  Spreadsheet.setColumnWidth(2,100);
}

//////////////////////////////////////////////////////////////////////////////ℱℳ
function LoadContacts() {
  
  //locate spreadsheet
  var cachedObject = CacheService.getPrivateCache().get('cachedObject');
  var Spreadsheet = SpreadsheetApp.openById(cachedObject);
  
  //load contacts
  var groupName = Spreadsheet.getRange("H1").getValue();
  if (groupName != "cancel") {
    var myGroup = ContactsApp.getContactGroup(groupName);
    
    if (myGroup) {
      var myContacts = myGroup.getContacts();
      for (i=0; i < myContacts.length; i++) {   
        var myContact = [[myContacts[i].getFullName(), myContacts[i].getPrimaryEmail(), ""]];
        Spreadsheet.getActiveSheet().getRange(i+2, 1,1,3).setValues(myContact);
      }
      
      if ( myContacts.length == 0) 
        //Browser.msgBox("Google Contacts Error", "Sorry, Gmail could not find any contacts in the specified Group.", Browser.Buttons.OK);
        Spreadsheet.getActiveSheet().getRange("H1").setValue("Sorry, Gmail could not find any contacts in the specified Group.");
    }
    else
      //Browser.msgBox("Google Contacts Error", "Sorry, the specified Google Contacts Group does not exist.", Browser.Buttons.OK);
      Spreadsheet.getActiveSheet().getRange("H1").setValue("Sorry, the specified Google Contacts Group does not exist.");
  } 
}

//////////////////////////////////////////////////////////////////////////////ℱℳ
function SendMessages() {
  
  //locate spreadsheet
  var cachedObject = CacheService.getPrivateCache().get('cachedObject');
  var Spreadsheet = SpreadsheetApp.openById(cachedObject);
  var myContact = Spreadsheet.getDataRange().getValues();
  var me = Session.getUser().getEmail();
  var mme = Session.getUser().getEmail();
  
  //prepare and send messages
  for (i=1; i < myContact.length; i++) {
    var col = myContact[i];
    var file      = col[4];
    var attachmentID  = file?DocsList.getFileById(file):"";
    if (col[5] != "sent") {
      var emailMsg = "Hello, " + col[0] + ", <br /><br />" + col[2];
      var advancedArgs = {htmlBody:emailMsg, name:me, replyTo:mme};
      
      if (file)
        advancedArgs["attachments"] = attachmentID;
      
      GmailApp.sendEmail(col[1], col[3], col[2], advancedArgs);
      Spreadsheet.getActiveSheet().getRange(i+1,6).setValue("sent"); 
      SpreadsheetApp.flush();
    }
  }
}


