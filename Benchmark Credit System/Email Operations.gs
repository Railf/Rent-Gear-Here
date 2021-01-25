function
SendEmail(recipient)
{
  var guest = {
    booking_number:      recipient[0],
    unit_code:           recipient[1],
    property_name:       recipient[2],
    first_name:          recipient[3],
    last_name:           recipient[4],
    email_address:       recipient[5],
    coupon_code:         recipient[6],
    coupon_amount:       recipient[7],
    tier:                recipient[8],
    first_night:         recipient[9],
    last_night:          recipient[10],
    nights:              recipient[11],
    reservation_type:    recipient[12],
    beach_access_notice: recipient[13],
    month_out:           recipient[14],
    two_weeks_out:       recipient[15]
  };
    
  var subject = "Rent Gear Here - Your rental credit awaits!";
  var template = HtmlService.createTemplateFromFile('Template.html');
  template.guest = guest;
  var message = template.evaluate().getContent();
  
  MailApp.sendEmail({
    to: guest.email_address,
    subject: subject,
    htmlBody: message
  });
}


function
GetEmailList()
{
  var today        = new Date();
  var guestSummary = GetGuestSummary();
  var list         = [];

  var firstDay;
  var monthOut;
  var twoWeeksOut;

  for (var i = 0; i < guestSummary.length; ++i)
  {
    firstDay    = guestSummary[i][9];
    monthOut    = guestSummary[i][14];
    twoWeeksOut = guestSummary[i][15];

    if (IsSameDate(today,firstDay) || IsSameDate(today,monthOut) || IsSameDate(today,twoWeeksOut)) list.push(guestSummary[i]);
  }

  return list;
}


function
EmailIndividual()
{
  var guestSummary = GetGuestSummary();
  var ui           = SpreadsheetApp.getUi();
  var response     = ui.prompt('Enter Guest Booking Number:');
  var guest;

  for (var i = 0; i < guestSummary.length; ++i)
  {
    if (guestSummary[i][0] == response.getResponseText())
    {
      guest = guestSummary[i];
      break;
    }
  }

  SendEmail(guest);
}


function
SendEmailsToTodaysList()
{
  var emailList = GetEmailList();

  for (var i = 0; i < emailList.length; ++i)
  {
    SendEmail(emailList[i]);
  }

  SpreadsheetApp.getUi().alert('' + emailList.length + ' emails sent.');
}


function EmailQuota() { SpreadsheetApp.getUi().alert(MailApp.getRemainingDailyQuota()); }
