/*===========================================================
  Geosort.gs
  -------------------
  
  Geographically sort addresses provided by the user,
  using Google's Maps.newDirectionFinder().
  
  
  LEGEND:
  
    GEOSORTE
      => returns optimal East-to-West route order
    
    GEOSORTW
      => returns optimal West-to-East route order
  
  
  
  NOTES:
  
  - Rent Gear Here's area of operation is:
    Destin, FL to Panama City Beach, FL.
    
  - Fort Walton Bridge is used as the Eastern boundary.
  
  - Panama City Bridge is used as the Western boundary.
  
  
  
  KNOWN ISSUE(S):
  
  1) There is a limit of 25 waypoints.
  
  
  
  USE EXAMPLE:
  
  =GEOSORTE(L2:L37)
=============================================================
Ralph McCraken, III */



function
GEOSORTE(addresses)
{ 
  if (addresses.length > 25)
    return "GEOSORT has a 25-address limit. Please narrow the selection to 25 addresses.";
  
  const  map = Maps.newDirectionFinder();
  const west = "42 FL-30, Fort Walton Beach, FL 32548";
  const east = "17 US-98, Panama City Beach, FL 32407";
  
  map.setOrigin(east);
  map.setDestination(west);
  map.setOptimizeWaypoints(true);
  map.setMode(Maps.DirectionFinder.Mode.DRIVING);
  
  for (var i = 0; i < addresses.length; ++i)
    map.addWaypoint(addresses[i] + ", FL, 32459");

  return map.getDirections().routes[0].waypoint_order;
}



function
GEOSORTW(addresses)
{ 
  if (addresses.length > 25)
    return "GEOSORT has a 25-address limit. Please narrow the selection to 25 addresses.";
  
  const  map = Maps.newDirectionFinder();
  const west = "42 FL-30, Fort Walton Beach, FL 32548";
  const east = "17 US-98, Panama City Beach, FL 32407";
  
  map.setOrigin(west);
  map.setDestination(east);
  map.setOptimizeWaypoints(true);
  map.setMode(Maps.DirectionFinder.Mode.DRIVING);
  
  for (var i = 0; i < addresses.length; ++i)
    map.addWaypoint(addresses[i] + ", FL, 32459");

  return map.getDirections().routes[0].waypoint_order;
}
