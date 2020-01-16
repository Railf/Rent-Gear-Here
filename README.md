# Rent Gear Here - Solutions with Code
Code written while working for [Rent Gear Here](https://www.rentgearhere.com) (RGH),
as **Systems & Data Manager** (May 2019 - present).



## Delivery Doc
Code written for RGH's "Delivery Doc."

### Global Variables
Global variables are not naturally supported, in Google Script code.
This is a work-around.
```
var nameColumn = Global().c_first_name;
```

### Scheduler

### Geosort
Geographically sort addresses provided by the user,
using Google's Maps.newDirectionFinder(),
working East-to-West or West-to-East.
```
=GEOSORTE(L2:L37)
```
```
=GEOSORTW(L2:L37)
```

### Summary Box Generator
Generates a summary box of all items needed, for a particular route.
```
=BOX(A2,2:37)
```

### Helper Functions
Functions that can be utilized in more than one project,
in an effort to not recreate the wheel.
```
if(isLeftContainingRight(large,small)) { Logger.log("Left contains right."); }
```

## Bike Program Schedule
Code written for RGH's "Bike Program Schedule."



## Bike & Beach Master
Code written for RGH's "Bike & Beach Master."



## Baby Master
Code written for RGH's "Baby Master."
