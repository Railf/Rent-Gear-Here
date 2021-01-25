function
onOpen()
{
  SpreadsheetApp.getUi().createMenu('Benchmark')
  .addItem('Import Guests', 'ImportGuests')
  .addSeparator()
  .addItem('Email Individual Guest','EmailIndividual')
  .addItem('Send Today\'s Emails','SendEmailsToTodaysList')
  .addSeparator()
  .addItem('Get Remaining Email Quota', 'EmailQuota')
  .addToUi();
}


function
ImportGuests()
{
  var currentBookings = GetCurrentBookings();
  
  var ss   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA IMPORT');
  var data = ss.getRange(2, 1, ss.getLastRow()-1, ss.getLastColumn()-1).getValues();
  
  var newGuestBookings = [];
  for (var i = 0; i < data.length; ++i)
  {
    if (IsNewValidBooking(data[i][1],currentBookings,data[i][26],data[i][24],data[i][29]))
    {
      var isNewGuest = true;
      
      for (var j = 0; j < currentBookings.length; ++j)
      {
        if (data[i][1] == currentBookings[j])
        {
          isNewGuest = false;
          break;
        }
      }
      
      if (isNewGuest)
      {
        var tier        = GetTier(data[i][26]);
        var coupon      = GetCoupon(tier, data[i][24]);
        var monthOut    = GetMonthOutDate(data[i][21]);
        var twoWeeksOut = GetTwoWeeksOutDate(data[i][21]);
        var beachAccess = GetBeachAccessRecommendation(data[i][26]);

        var booking = [];
        booking.push(data[i][1]);  // bookingNumber
        booking.push(data[i][26]); // unitCode
        booking.push(data[i][25]); // propertyName
        booking.push(data[i][3]);  // firstName
        booking.push(data[i][4]);  // lastName
        booking.push(data[i][13]); // emailAddress
        booking.push(coupon[0]);   // couponCode        -- GetCoupon(tier, nights)[0]
        booking.push(coupon[1]);   // couponValue       -- GetCoupon(tier, nights)[1]
        booking.push(tier);        // tier              -- GetTier(unitCode)
        booking.push(data[i][21]); // firstNight
        booking.push(data[i][22]); // lastNight
        booking.push(data[i][24]); // nights
        booking.push(data[i][29]); // reservationType
        booking.push(beachAccess); // beachAccessNotice -- GetBeachAccessRecommendation(unitCode) -- TODO
        booking.push(monthOut);    // monthOut          -- GetMonthOutDate(firstNight)
        booking.push(twoWeeksOut); // twoWeeksOut       -- GetTwoWeeksOutDate(firstNight)
        
        newGuestBookings.push(booking);
      }
    }
  }
  
  AddGuestsToRecord(newGuestBookings);
}


function
GetCurrentBookings()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GUEST SUMMARY');
  var data  = sheet.getRange('A2:A').getValues();
  
  return data;
}


function
GetGuestSummary()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GUEST SUMMARY');
  var data  = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  
  return data;
}


function
IsNewValidBooking(bookingNumber,currentBookings, unitCode, nights, type)
{
  var isBooking = IsRightInLeft(bookingNumber,'BKG');
  var isOnCurrentGuestList = IsRightInLeft(currentBookings,bookingNumber);
  var isOnTierProgram = (IsRightInLeft(unitCode,'*') || IsRightInLeft(unitCode,'^') || IsRightInLeft(unitCode,'+') || IsRightInLeft(unitCode,'-'));
  var isMoreThanTwoNights = nights > 2;
  var isRenterOrOwnerReferral = (IsRightInLeft(type,'Renter') || IsRightInLeft(type,'Owner Referral'));

  return (isBooking && !isOnCurrentGuestList && isOnTierProgram && isMoreThanTwoNights && isRenterOrOwnerReferral);
}


function
AddGuestsToRecord(bookings)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GUEST SUMMARY');
  sheet.activate();
  sheet.getRange(sheet.getLastRow()+1, 1, bookings.length, sheet.getLastColumn()).setValues(bookings);
}


function
GetTier(unitCode)
{
  var tier = 0;
  
  if      (IsRightInLeft(unitCode,'*')) tier = 1;
  else if (IsRightInLeft(unitCode,'^')) tier = 2;
  else if (IsRightInLeft(unitCode,'+')) tier = 3;
  else if (IsRightInLeft(unitCode,'-')) tier = 4;
  
  return tier;
}


function
GetCoupon(tier, nights)
{
  var sheet;
  var data;

  var coupon           = [];
  var optionFromNights = 1;

  if (nights > 14) optionFromNights = 14;
  else             optionFromNights = nights;

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COUPON CODES');
  data  = sheet.getRange(optionFromNights-1,tier).getValue();
  coupon.push(data);

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COUPON VALUES');
  data  = sheet.getRange(optionFromNights-1,tier).getValue();
  coupon.push(data);

  return coupon;
}


function
GetBeachAccessRecommendation(unitCode) // TODO
{
  var beachOptions = GetBeachAccessOptions();
  var property = [];
  var options = [];
  var recommendation = "";

  for (var i = 0; i < beachOptions.length; ++i)
  {
    if (IsRightInLeft(unitCode,beachOptions[i][0]))
    {
      property = beachOptions[i];
      break;
    }
  }

  for (var i = 0; i < 3; ++i)
  {
    if (property[6+i] != '')
    {
      options.push(property[6+i]);
    }
  }

  recommendation = GetRecommendationWording(options);

  return recommendation;
}


function
GetBeachAccessOptions()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BEACH ACCESS OPTIONS');
  var data  = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  
  return data;
}


function
GetRecommendationWording(options)
{
  var recommendation = "";

  if (options == null || options.length == 0)
  {
    recommendation += "We recommend the Backpack Chair Rental item instead of Beach Chair Service. There is not a beach access near enough to the property you are staying in, for us to confidently recommend our Beach Chair Service. The good news is that with the Backpack Chair Rental, you can utilize any public beach space, instead of being restricted to a vendor zone that applies to the Beach Chair Service!";
  }

  else if (options.length == 1)
  {
    if (options[0] != "Backpack Chairs")
    {
      recommendation += "We recommend " + options[0] + " for Beach Chair Service.";
    }
    else if (options[0] == "Backpack Chairs")
    {
      recommendation += "We recommend the Backpack Chair Rental item instead of Beach Chair Service. There is not a beach access near enough to the property you are staying in, for us to confidently recommend our Beach Chair Service. The good news is that with the Backpack Chair Rental, you can utilize any public beach space, instead of being restricted to a vendor zone that applies to the Beach Chair Service!";
    }
  }

  else if (options.length == 2)
  {
    if (options[0] != "Backpack Chairs")
    {
      recommendation += "We recommend " + options[0] + " for Beach Chair Service.";
    }
    else if (options[0] == "Backpack Chairs")
    {
      recommendation += "We recommend the Backpack Chair Rental item instead of Beach Chair Service. There is not a beach access near enough to the property you are staying in, for us to confidently recommend our Beach Chair Service. The good news is that with the Backpack Chair Rental, you can utilize any public beach space, instead of being restricted to a vendor zone that applies to the Beach Chair Service!";
    }

    recommendation += " Alternatively, ";

    if (options[1] != "Backpack Chairs")
    {
      recommendation += "we recommend " + options[1] + " for Beach Chair Service.";
    }
    else if (options[1] == "Backpack Chairs")
    {
      recommendation += "we recommend the Backpack Chair Rental item instead of Beach Chair Service. The good news is that with the Backpack Chair Rental, you can utilize any public beach space, instead of being restricted to a vendor zone that applies to the Beach Chair Service!";
    }
  }

  else if (options.length == 3)
  {
    if (options[0] != "Backpack Chairs")
    {
      recommendation += "We recommend " + options[0] + " for Beach Chair Service.";
    }
    else if (options[0] == "Backpack Chairs")
    {
      recommendation += "We recommend the Backpack Chair Rental item instead of Beach Chair Service. There is not a beach access near enough to the property you are staying in, for us to confidently recommend our Beach Chair Service. The good news is that with the Backpack Chair Rental, you can utilize any public beach space, instead of being restricted to a vendor zone that applies to the Beach Chair Service!";
    }

    recommendation += " Alternatively, ";

    if (options[1] != "Backpack Chairs")
    {
      recommendation += "we recommend " + options[1] + " for Beach Chair Service.";
    }
    else if (options[1] == "Backpack Chairs")
    {
      recommendation += "we recommend the Backpack Chair Rental item instead of Beach Chair Service. The good news is that with the Backpack Chair Rental, you can utilize any public beach space, instead of being restricted to a vendor zone that applies to the Beach Chair Service!";
    }

    recommendation += " Finally, ";

    if (options[2] != "Backpack Chairs")
    {
      recommendation += "we recommend " + options[2] + " for Beach Chair Service.";
    }
    else if (options[2] == "Backpack Chairs")
    {
      recommendation += "we recommend the Backpack Chair Rental item instead of Beach Chair Service. The good news is that with the Backpack Chair Rental, you can utilize any public beach space, instead of being restricted to a vendor zone that applies to the Beach Chair Service!";
    }
  }

  return recommendation;
}


function
GetMonthOutDate(firstNight)
{
  return new Date(firstNight.getFullYear(),firstNight.getMonth()-1,firstNight.getDate());
}


function
GetTwoWeeksOutDate(firstNight)
{
  return new Date(firstNight.getFullYear(),firstNight.getMonth(),firstNight.getDate()-14);
}


function
IsSameDate(date1, date2)
{
  var isSameDay   = (date1.getDate()     == date2.getDate());
  var isSameMonth = (date1.getMonth()    == date2.getMonth());
  var isSameYear  = (date1.getFullYear() == date2.getFullYear());

  return (isSameDay && isSameMonth && isSameYear);
}
