// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Extract Emails')
      .addItem('Extract Emails...', 'extractEmails')
      .addToUi();
}
 
// extract emails from label in Gmail
function extractEmails() {
  
  // get the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var label = sheet.getRange(1,2).getValue();
  
  // get all email threads that match label from Sheet
  var threads = GmailApp.search ("label:" + label);
  
  // get all the messages for the current batch of threads
  var messages = GmailApp.getMessagesForThreads (threads);
  
  var emailArray = [];
  
  // get array of email addresses
  var addedEmails = {};

  messages.forEach(function(message) {
    message.forEach(function(d) {
      var emailsArray = d.getFrom().match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
      var email = emailsArray[0];
      if (addedEmails[email] === undefined && 
          email.indexOf("f6s.com") === -1 && 
            email.indexOf("info") === -1  && 
            email.indexOf("invite") === -1  && 
            email.indexOf("noreplies") === -1  && 
            email.indexOf("order") === -1  && 
            email.indexOf("service") === -1  &&               
            email.indexOf("new") === -1  &&        
            email.indexOf("notif") === -1  &&        
            email.indexOf("admin") === -1  &&        
            email.indexOf("contact") === -1  &&        
            email.indexOf("events") === -1  &&        
            email.indexOf("mailer") === -1  &&        
            email.indexOf("updates") === -1  &&        
              email.indexOf("reply") === -1 &&
              email.indexOf("gov") === -1) {
          addedEmails[email] = 1;
          emailArray.push([email]);
        }
    });
  });
  
  sheet.getRange(4,1,emailArray.length,1).setValues(emailArray).sort(1);


  
  // clear any old data
  
  // paste in new names and emails and sort by email address A - Z
 
}
