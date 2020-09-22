/*===========================================================
  GlobalVariables.gs
  ------------------
  
  Global variables are not naturally supported,
  in Google Script code. This is a work-around.
  
  
  LEGEND:
  
    s_ => returns sheet id
    t_ => returns tab name
    c_ => returns column number
    n_ => returns number
  
  
  USE EXAMPLE:
  
    var nameColumn = Global().c_first_name;
=============================================================
Ralph McCraken, III */



function Global()
{
  var variables = {
    
    s_bp_schedule:   "1JZ-JDk4_VuX6AmlifvjnSVA1b43BA65PcRL2-RpVfbY",
    s_wos_register:  "1-tlME8HP42SwvM1xwvW2HEs1ZeS7gMdfMxgSkdhbewk",
    
    t_tba:           "TBA BIKE",
    t_workorders:    "WORK ORDERS",
    t_360_wo:        "360 WO",
    
    c_drop_sign:     1,
    c_drop_date:     2,
    c_driver_notes:  3,
    c_pick_sign:     4,
    c_pick_date:     5,
    c_company:       6,
    c_first_name:    7,
    c_last_name:     8,
    c_phone:         9,
    c_email:        10,
    c_area:         11,
    c_address:      12,
    c_attempt:      13,
    c_notes:        14,
    c_access:       15,
    c_summary:      16,
    c_item_start:   17,
    
    n_data_columns: 15
  };
  
  return variables;
}
