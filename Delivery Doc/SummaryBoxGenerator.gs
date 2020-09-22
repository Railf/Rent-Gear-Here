/*===========================================================
  SummaryBoxGenerator.gs
  ----------------------
  
  Generates a summary box of all items needed,
  for a particular route.
  
  
  LEGEND:
  
    BOX(date, route)
      => returns summary box data
      
    BuildSummary(items, stops, isDelivery)
      => builds the item summary
         see notes for algoritm
    
    BuildSheetSummary(items, stops, isDelivery)
      => build the sheets summary
         see notes
    
    getDateArray(date)
      => returns array object [month, day]
    
    getDeliveryStatuses(date, route)
      => returns array of booleans;
         true if stop's delivery date matches date argument,
         false otherwise
    
    getItemsFromStops(route)
      => returns array of arrays;
         each stop's items are stored in its own array;
         each array is pushed into one, encompassing array
    
    fetchItems()
      => returns array of item names found in the header row
    
    isSameDate(date1, date2)
      => returns true, if two date array objects match;
         returns false otherwise
  
  
  
  NOTES:
  
  - A row in the Delivery Doc is a "stop" made by the driver.
  
  - Builds a summary based on:
      items in the given tab,
      stops in the given route, and
      delivery status of each line in the route.
    
  - Algorithm (for each item):
      If not delivery, add to pickups.
      If delivery and no pickups, add to needed.
      If delivery and pickups, subtract from pickups.
        => If pickups ever go negative, each would-be negative is added to needed.
        
  - Sheet Algorithm (for each item):
      If the item contains the sheet-item keywords and is a delivery, add to needed.
  
  
  
  USE EXAMPLE:
  
  =BOX(B2, 2:35)
=============================================================
Ralph McCracken, III */


function
BOX(date, route)
{
  var items      = fetchItems();
  var stops      = getItemsFromStops(route);
  var isDelivery = getDeliveryStatuses(date, route);
  var summary    = BuildSummary(items, stops, isDelivery);
  var sheets     = BuildSheetSummary(items, stops, isDelivery);
  
  if (sheets != "")
    return (summary + '\n' + '\n' + sheets);
  
  return summary;
}



function
BuildSummary(items, stops, isDelivery)
{
  var summary = "";
  var needed  = [];
  var need    =  0;
  var count   =  0;
  var amount  =  0;
    
  for (var i = 0; i < stops[0].length; ++i)
  {
    for (var j = 0; j < stops.length; ++j)
    {
      amount = +stops[j][i];
      
      if (!isDelivery[j])
      {
        count += amount;
      }
      else if (isDelivery[j] && count <= 0)
      {
        need  += amount;
      }
      else if (isDelivery[j] && count >  0)
      {
        count -= amount;
        
        if (count < 0)
        {
          need += (-1)*(count);
          count = 0;
        }
      }
      
      amount = 0;
    }
    
    needed.push(need);
    need  = 0;
    count = 0;
  }
  
  var first = true;
  for (var i = 0; i < items.length; ++i)
  {
    if (needed[i] == 0) continue;
    
    if (first)
    {
      summary += "(" + needed[i] + ") " + items[i];
      first    = false;
    }
    else
    {
      summary += '\n' + "(" + needed[i] + ") " + items[i];
    }
  }
  
  return summary;
}



function
BuildSheetSummary(items, stops, isDelivery)
{
  var summary = "";
  var needed  = [];
  var need    =  0;
  var count   =  0;
  var amount  =  0;
    
  for (var i = 0; i < stops[0].length; ++i)
  {
    for (var j = 0; j < stops.length; ++j)
    {
      amount = +stops[j][i];
      
      if (isDelivery[j])
        need += amount;

      amount = 0;
    }
    
    needed.push(need);
    need  = 0;
    count = 0;
  }
  
  var crib  = false;
  var pnp   = false;
  var twin  = false;
  
  var cribs = 0;
  var pnps  = 0;
  var twins = 0;
  
  for (var i = 0; i < items.length; ++i)
  {
    if (needed[i] == 0) continue;
    
    if ( isLeftContainingRight(items[i],"Full")    ) { crib = true; cribs += needed[i]; }
    if ( isLeftContainingRight(items[i],"Compact") ) { crib = true; cribs += needed[i]; }
    if ( isLeftContainingRight(items[i],"Pac")     ) { pnp  = true; pnps  += needed[i]; }
    if ( isLeftContainingRight(items[i],"Bjorn")   ) { pnp  = true; pnps  += needed[i]; }
    if ( isLeftContainingRight(items[i],"Roll")    ) { twin = true; twins += needed[i]; }
  }
  

  if (!crib && !pnp && !twin)     // 000
  {
    summary = "";
  }
  else if (!crib && !pnp && twin) // 001
  {
    summary += "(" + twins + ") "  + "Twin Sheet";
  }
  else if (!crib && pnp && !twin) // 010
  {
    summary += "(" + pnps   + ") " + "PNP Sheet";
  }
  else if (!crib && pnp && twin)  // 011
  {
    summary += "(" + pnps   + ") " + "PNP Sheet";
    summary += '\n';
    summary += "(" + twins + ") "  + "Twin Sheet";
  }
  else if (crib && !pnp && !twin) // 100
  {
    summary += "(" + cribs + ") "  + "Crib Sheet";
  }
  else if (crib && !pnp && twin)  // 101
  {
    summary += "(" + cribs + ") "  + "Crib Sheet";
    summary += '\n';
    summary += "(" + twins + ") "  + "Twin Sheet";
  }
  else if (crib && pnp && !twin)  // 110
  {
    summary += "(" + cribs + ") "  + "Crib Sheet";
    summary += '\n';
    summary += "(" + pnps   + ") " + "PNP Sheet";
  }
  else if (crib && pnp && twin)   // 111
  {
    summary += "(" + cribs + ") "  + "Crib Sheet";
    summary += '\n';
    summary += "(" + pnps   + ") " + "PNP Sheet";
    summary += '\n';
    summary += "(" + twins + ") "  + "Twin Sheet";
  }

  
  return summary;
}



function
getDateArray(date)
{
  return [date.getMonth() + 1, date.getDate()];
}



function
getDeliveryStatuses(date, route)
{ 
  var userDate   = getDateArray(date);
  var rows       = route.length;
  var isDelivery = [];
  
  for (var i = 0; i < rows; ++i)
  {
    if (route[i][1] == "")
    {
      isDelivery[i] = false;
      continue;
    }
    
    if (isSameDate(userDate,getDateArray(route[i][1])))
      isDelivery[i] = true;
    else
      isDelivery[i] = false;
  }
  
  return isDelivery;
}



function
getItemsFromStops(route)
{
  var stops   = [];
  
  var rows    = route.length;
  var columns = route[0].length;
  var stop    = [];
  var start   = Global().c_item_start;
  
  for (var j = 0; j < rows; ++j)
  {
    for (var i = 0; i < (columns - start); ++i)
      stop.push(route[j][start+i]);
    
    stops.push(stop);
    stop = [];
  }
  
  return stops;
}



function
fetchItems()
{ 
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var width = ss.getLastColumn();
  var names = ss.getActiveSheet().getRange(1,1,1,width).getValues();
  var start = Global().c_item_start;
  var items = [];
  
  for (var i = 0; i < width - start; ++i)
    items.push(names[0][start+i]);
  
  return items;
}



function
isSameDate(date1, date2)
{
  if (date1[0] != date2[0]) return false;
  if (date2[1] != date2[1]) return false;
  
  return true;
}
