function sendEmail() {
  var sv = SpreadsheetApp.getActiveSpreadsheet()

  // get workshop details
  var batch_sheet_name = 'batch3'
  var form_data=sv.getSheetByName(batch_sheet_name);
  var subject = form_data.getRange(1,10).getValue();
  var batch_no = form_data.getRange(2,10).getValue();
  var starting_seq_id = form_data.getRange(10,10).getValue();
  var index_count = form_data.getRange(11,10).getValue();


  // time math
  var event_duration_hrs = form_data.getRange(6,10).getValue();
  var event_duration_ms =  event_duration_hrs * 60 * 60 * 1000; // in milliseconds
  
  // date and time
  var event_date = form_data.getRange(3,10).getValue();
  var event_start_date = new Date(event_date);
  var event_end_date = new Date(event_start_date.getTime() + event_duration_ms);
  
  // set event time details for html
  var event = {
    date: Utilities.formatDate(event_start_date, "Europe/London", "EEEEEEEE, MMMMM dd, yyyy"),
    timelink: form_data.getRange(5,10).getValue(),
    start_time_uk: Utilities.formatDate(event_start_date, "Europe/London", "hh:mm a"),
    end_time_uk: Utilities.formatDate(event_end_date, "Europe/London", "hh:mm a"),
    loc_uk: "UK",

    start_time_in: Utilities.formatDate(event_start_date, 'Asia/Calcutta', 'hh:mm a'),
    end_time_in: Utilities.formatDate(event_end_date, 'Asia/Calcutta', 'hh:mm a'),
    loc_in: "India",

    start_time_pst: Utilities.formatDate(event_start_date, 'America/Los_Angeles', 'hh:mm a'),
    end_time_pst: Utilities.formatDate(event_end_date, 'America/Los_Angeles', 'hh:mm a'),
    loc_pst: "PST",
  
    zoom_link: form_data.getRange(7,10).getValue(),
    zoom_meeting_id: form_data.getRange(8,10).getValue(),
    zoom_passcode: form_data.getRange(9,10).getValue(),
  };
  
  index_start = 2;
  for (var index = index_start; index < index_start + index_count; index++ ) {

    // extract registrant info
    var emailAddress = form_data.getRange(index, 2).getValue();
    var name = form_data.getRange(index, 3).getValue();
    var status = form_data.getRange(index, 6).getValue();

    // Already registered
    if (status != "") {
      console.log('Skip ' + name);
      continue;
    }

    // create id
    seq = starting_seq_id + index - index_start;
    var id = 'SVUW' + batch_no + '-' + seq.toString();

    // prepare message for this user
    var templ = HtmlService.createTemplateFromFile('registration-email');
    templ.first_name = name;
    templ.regid = id;
    templ.event = event;
    var message = templ.evaluate().getContent();
    
    // send email with customized message
    MailApp.sendEmail({
      to: emailAddress,
      subject: subject,
      htmlBody: message
    });

    console.log('Sent to ' + name);

    // mail sent. Bookkeeping now.
    form_data.getRange(index,5).setValue(id);
    form_data.getRange(index,6).setValue('SENT');
  }
} 