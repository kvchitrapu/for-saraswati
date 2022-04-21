function copyDoc(srcId, targetId, name, id) {
  
  var sourceDoc = DocumentApp.openById(srcId);
  var targetDoc = DocumentApp.openById(targetId);
  var totalElements = sourceDoc.getNumChildren();

  for( var j = 0; j < totalElements; ++j ) {
    var body = targetDoc.getBody()
    var element = sourceDoc.getChild(j).copy();
    var type = element.getType();
    if( type == DocumentApp.ElementType.PARAGRAPH ){
      body.appendParagraph(element.replaceText('<name>', name));
    }
    else if( type == DocumentApp.ElementType.TABLE){
        body.appendTable(element.replaceText('<regno>', id));
      }
    else if( type == DocumentApp.ElementType.LIST_ITEM){
        body.appendListItem(element);
      }
    }
}

function sendEmail() {
  var sv = SpreadsheetApp.getActiveSpreadsheet()
  var form_data=sv.getSheetByName('registered');
  var emails= sv.getSheetByName('send-email');
  var subject = emails.getRange(2,1).getValue();;
  // var message = emails.getRange(2,2).getValue();
  
  // let doc = DocumentApp.openById('1D46QeOKx8sDSijs_eekjXS8Dj9QiFTRabDAVCrP-UIo')
  
  // message = doc.getBody().getText();
  var n=form_data.getLastRow();
  for (var i = 2; i < n+1 ; i++ ) {

    

    var emailAddress = form_data.getRange(i,2).getValue();
    var name = form_data.getRange(i,3).getValue();
    var id = 'SVUW2-' + i.toString();
    var user_doc_name = 'RegEmail-' + id
    let user_doc = DocumentApp.create(user_doc_name);
    user_doc.saveAndClose();
    copyDoc(doc.getId(), user_doc.getId(), name, id);
    // user_body.replaceText('<name>', name);
  

    form_data.getRange(i,5).setValue(id);
    form_data.getRange(i,6).setValue('SENT');
    // message=message.replace("<name>",name);
    MailApp.sendEmail(emailAddress, subject, user_doc.getUrl());
  }
}