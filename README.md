# Rent Gear Here - Solutions with Code
Code written while at working for [Rent Gear Here](https://www.rentgearhere.com) (RGH),
as **Systems & Data Manager** (May 2019 - present).



## Delivery Doc
Code written for RGH's Delivery Doc.

### Global Variables
Global variables are not naturally supported, in Google Script code.
This is a work-around.
```
var sheetID = Global().s_bp_schedule;
```
```
var tabName = Global().t_workorders;
```
```
var nameColumn = Global().c_first_name;
```
```
var dataColumns = Global().n_data_columns;
```

### Scheduler

### Geosort
Geographically sort addresses provided by the user,
using Google's Maps.newDirectionFinder(),
work East-to-West (GEOSORTE) or West-to-East (GEOSORTW).
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
if(isLeftContainingRight(large,small)) {}
```

## Bike Program Schedule
Code written for RGH's Bike Program Schedule.



## Bike & Beach Master
Code written for RGH's Bike & Beach Master.



## Baby Master
Code written for RGH's Baby Master.
