function sendEmail() {
    var sv = SpreadsheetApp.getActiveSpreadsheet()
    var form_data=sv.getSheetByName('registered');
    var emails= sv.getSheetByName('send-email');
    var subject = emails.getRange(2,1).getValue();;
    var n=form_data.getLastRow();
    for (var i = 2; i < n+1 ; i++ ) {
  
      // extract registrant info
      var emailAddress = form_data.getRange(i,2).getValue();
      var name = form_data.getRange(i,3).getValue();
      var status = form_data.getRange(i,6).getValue();
  
      // Already set a registration
      if (status != "") {
        continue;
      }
  
      // create id
      var id = 'SVUW2-' + i.toString();
  
      // prepare message
      var templ = HtmlService.createTemplateFromFile('registration-email');
      templ.first_name = name;
      templ.regid = id;
      var message = templ.evaluate().getContent();
      
      // send email
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message
    });
  
      // set values
      form_data.getRange(i,5).setValue(id);
      form_data.getRange(i,6).setValue('SENT');
    }
  }