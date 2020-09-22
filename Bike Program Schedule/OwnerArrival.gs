/*===========================================================
  OwnerArrival.gs
  -------------------
  
  Fetch the "OWNER ARRIVAL" status,
  for the property that exists on the same row,
  if an RGH partner has an owner arrival for said property;
  otherwise, return nothing.
  
  
  LEGEND:
  
    fetchOwnerArrival()
      => return "OWNER ARRIVAL,"
      if results of external query include an owner keyword
    
    isValidSheet(spreadsheet)
      => return true,
      if sheet name is one of the valid sheet names
    
    isPropertyInResults(property, results)
      => return true,
      if a single property exists in the results array
    
    isOwnerArrival(status)
      => returns true,
      if the status is an owner-arrival keyword
    
    fetchResults(spreadsheet)
      => return results from QUERY
    
    fetchProperties(spreadsheet)
      => return properties (homes) from QUERY
    
    fetchStatuses(spreadsheet)
      => return statuses from QUERY
  
  
  
  NOTES:
  
  - Work based around Google Sheets's =QUERY() function.
  
  - A QUERY pulls departures and arrivals,
  based on a given date, from RGH partner occupancy reports.
  
  - A separate QUERY function looks at the results,
  and pulls the property data for properties that exist,
  on RGH's Bike Program Schedule.
  
  - fetchOwnerArrivals looks at both QUERYs,
  and works to highlight owner arrivals.
  
  - When an RGH bike tech sees that a bike check is an
  "OWNER ARRIVAL," they put more effort into their check.
  
  - Owner-arrival keywords are based on how RGH's partners
  each express an owner arrival in their occupancy reports.
  
  
  
  KNOWN ISSUE(S):
  
  1) spreadsheet.getRange() is called three times.
  This is not the most efficient approach; and
  could be called once, and arrays could do the work.
  Currently using "fetch" instead of "get" in function names,
  to convey the slowness of the operations.
  
  
  
  USE EXAMPLE:
  
  =fetchOwnerArrival()
=============================================================
Ralph McCraken, III */



function fetchOwnerArrival()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (isValidSheet(spreadsheet))
  {
    var results    = fetchResults(spreadsheet);
    var properties = fetchProperties(spreadsheet);
    var statuses   = fetchStatuses(spreadsheet);
//    var nextIn     = fetchNextArrivals(spreadsheet);
    
    if (spreadsheet.getName() == "CHECKS_GARTS")
    {
      for (var i = 0; i < statuses.length; ++i)
      {
        if(isPropertyInResults(properties[i], results) && isOwnerDepart(statuses[i]))
          return "OWNER DEPART";
        if(isPropertyInResults(properties[i], results) && isGuestOfOwnerDepart(statuses[i]))
          return "GUEST OF OWNER DEPART";
      }
    }
    
    for (var i = 0; i < statuses.length; ++i)
    {
      if(isPropertyInResults(properties[i], results) && isOwnerArrival(statuses[i]))
        return "OWNER ARRIVAL";
//      else if(isPropertyInResults(properties[i], results) && !isOwnerArrival(statuses[i]))
//        return nextIn[i];
    }
    
  }
}



function
isValidSheet(spreadsheet)
{
  var sheet = spreadsheet.getName();
  
  var validSheets = [
    "CHECKS_RGH",
    "CHECKS_BAMBOO",
    "CHECKS_GARTS"
  ];
  
  return (validSheets.indexOf(sheet) > -1);
}



function
isPropertyInResults(property, results)
{
  return (results.toString().indexOf(property) > -1);
}



function
isOwnerArrival(status)
{
  return (
    status == "Owner"          ||
    status == "Guest of Owner" ||
    status == "Owner Referral" ||
    status == "Complimentary"  ||
    status == "Group"          ||
    status == "Long Term"      ||
    status == "OWN1"           ||
    status == "OWN2"           ||
    status == "OWN3"           ||
    status == "OWN4"           ||
    status == "OWN5"           ||
    status == "OWN6"           ||
    status == "OWN-MOD"        ||
    status == "OWN-BKD"
  );
}



function
isOwnerDepart(status)
{
  return (
    status == "Owner Depart"             ||
    status == "Owner Departure"
  );
}

function
isGuestOfOwnerDepart(status)
{
  return (
    status == "Guest of Owner Depart"    ||
    status == "Guest of Owner Departure"
  );
}



function
fetchResults(spreadsheet)
{
  var row = spreadsheet.getActiveCell().getRow();
  
  if      (spreadsheet.getName() == "CHECKS_RGH"   ) return (spreadsheet.getRange(row, 15).getValues());
  else if (spreadsheet.getName() == "CHECKS_BAMBOO") return (spreadsheet.getRange(row, 13).getValues());
  else if (spreadsheet.getName() == "CHECKS_GARTS" ) return (spreadsheet.getRange(row, 18).getValues());
}



function
fetchProperties(spreadsheet)
{
  if      (spreadsheet.getName() == "CHECKS_RGH"   ) return (spreadsheet.getRange("L2:L").getValues());
  else if (spreadsheet.getName() == "CHECKS_BAMBOO") return (spreadsheet.getRange("K2:K").getValues());
  else if (spreadsheet.getName() == "CHECKS_GARTS" ) return (spreadsheet.getRange("O2:O").getValues());
}



function
fetchStatuses(spreadsheet)
{
  if      (spreadsheet.getName() == "CHECKS_RGH"   ) return (spreadsheet.getRange("M2:M").getValues());
  else if (spreadsheet.getName() == "CHECKS_BAMBOO") return (spreadsheet.getRange("L2:L").getValues());
  else if (spreadsheet.getName() == "CHECKS_GARTS" ) return (spreadsheet.getRange("P2:P").getValues());
}



function
fetchNextArrivals(spreadsheet)
{
  if      (spreadsheet.getName() == "CHECKS_RGH"   ) return (spreadsheet.getRange("N2:N").getValues());
  if      (spreadsheet.getName() == "CHECKS_GARTS" ) return (spreadsheet.getRange("Q2:Q").getValues());
}
