/*===========================================================
  Scheduler.gs
  ------------
  
  Schedule 360BLUE work orders,
  and keep track of their completion status.
  
  
  LEGEND:
  
    onOpen()
      => Add a menu item that performs the "Get Work Orders" operation.
      
    onEditTrigger(e)
      => Detect stop-completed case.
    
    CheckValidConditions(e)
      => Check whether certain conditions are met, prior to proceeding further.
    
    PurpleStop(e)
    
    isDelivery(e)
    
    isPickup(e)
    
    noStopIssues(e)
    
    SendConfirmation(e)
    
    getWorkOrders()
    
    getDate()
    
    getConvertedDate(date)
    
    getPropertyDetails(property_id)
    
    getEmailAddress(from)
    
    getWorkOrderNumber(heading)
    
    isNewWorkOrder(workOrderNumber, record)
    
    RecordWorkOrders(recordDetails)
    
    getWorkOrderRecord()
    
    getWorkOrderStatus(workorder)
    
    setWorkOrderStatus(workorder, address, email, status)
    
    
  
  USE EXAMPLE:
  
    Navigate to menu item, "360 Blue," and click "Get Work Orders."
=============================================================
Ralph McCracken, III */


/*========================================================================================================================================================================
  onOpen()
  
  DESCRIPTION:
    ADD A CUSTOM MENU TO THE TOP OF THE SPREADSHEET.
    ADD A MENU ITEM THAT PERFORMS A 'Get Work Orders' OPERATION.
========================================================================================================================================================================*/

function
onOpen()
{
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("360 Blue")
  .addItem('Get Work Orders', 'getWorkOrders')
  .addToUi();
}





/*========================================================================================================================================================================
  onEditTrigger(e)
  
  DESCRIPTION:
    DETECT STOP COMPLETED CASE.
========================================================================================================================================================================*/

function
onEditTrigger(e)
{
  CheckValidConditions(e);
}





/*========================================================================================================================================================================
  CheckValidConditions(e)
  
  DESCRIPTION:
    CHECK WHETHER CERTAIN CONDITIONS ARE MET, PRIOR TO FURTHER OPERATIONS.
========================================================================================================================================================================*/

function
CheckValidConditions(e)
{
  if (e.range.getColumn() == Global().c_attempt && e.range.getValue() == "GTG")
  {
    var title;
    var ui;
    var response;
    var data = e.source.getActiveSheet().getRange(e.range.getRow(), Global().c_first_name, 1, 2).getValues();

    
    if (isLeftContainingRight(data[0][0],"360BLUE WORKORDER")) title = data[0][0];
    else                                          title = data[0][0] + ' ' + data[0][1];
    
    
    ui       = SpreadsheetApp.getUi();
    response = ui.alert(title, "Confirm stop is completed, and send confirmation email if applicable?\n\nPlease ensure stop is signed off with no issues present.\nThis action cannot be undone.", ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES)
    {
      SendConfirmation(e);
      PurpleStop(e);
    }
    else
    {
      if (e.oldValue == undefined) e.range.setValue("");
      else                         e.range.setValue(e.oldValue);
    }
  }
}





/*========================================================================================================================================================================
  PurpleStop(e)
  
  DESCRIPTION:
    TURN THE COLOR OF THE STOP-SIGN-OFF AREA PURPLE--DESIGNATING A COMPLETED STOP.
========================================================================================================================================================================*/

function
PurpleStop(e)
{
  if (isDelivery(e))    sheet.getRange(row, Global().c_drop_sign, 1, 2).setBackground("#B4A7D6");
  else if (isPickup(e)) sheet.getRange(row, Global().c_pick_sign, 1, 2).setBackground("#B4A7D6");
}





/*========================================================================================================================================================================
  isDelivery(e)
  
  DESCRIPTION:
    RETURN TRUE,  IF STOP IS SIGNED OFF ON THE DATE CONTAINED IN THE DROP DATE COLUMN.
    RETURN FALSE, IF STOP IS NOT SIGNED OFF ON THE DATE CONTAINED IN THE DROP DATE COLUMN.
========================================================================================================================================================================*/

function
isDelivery(e)
{
  var sheet = e.source.getActiveSheet();
  var drop  = getConvertedDate(sheet.getRange(e.range.getRow(), Global().c_drop_date).getValue());
  var date  = getDate();
  
  return (date == drop);
}





/*========================================================================================================================================================================
  isPickup(e)
  
  DESCRIPTION:
    RETURN TRUE,  IF STOP IS SIGNED OFF ON THE DATE CONTAINED IN THE PICK DATE COLUMN.
    RETURN FALSE, IF STOP IS NOT SIGNED OFF ON THE DATE CONTAINED IN THE PICK DATE COLUMN.
========================================================================================================================================================================*/

function
isPickup(e)
{
  var sheet = e.source.getActiveSheet();
  var pick  = getConvertedDate(sheet.getRange(e.range.getRow(), Global().c_pick_date).getValue());
  var date  = getDate();
  
  return (date == pick);
}



/*========================================================================================================================================================================
  noStopIssues(e)
  
  DESCRIPTION:
    CHECK WHETHER THE STOP HAS ANY ISSUE STARS (*).
    
  NOTES:
    CURRENTLY NOT IN USE.
========================================================================================================================================================================*/

function
noStopIssues(e)
{
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getActiveSheet();
  var range  = sheet.getRange(e.range.getRow(), 1, 1, (Global().n_data_columns));
  var values = range.getValues()[0];
  
  return !(
    isLeftContainingRight(values[Global().c_drop_off     - 1],"*") ||
    isLeftContainingRight(values[Global().c_driver_notes - 1],"*") ||
    isLeftContainingRight(values[Global().c_pick_up      - 1],"*") ||
    isLeftContainingRight(values[Global().c_notes        - 1],"*")
  );
}





/*========================================================================================================================================================================
  SendConfirmation(e)
  
  DESCRIPTION:
    SEND EMAIL CONFIRMATION--360 WORK ORDER, DELIVERY, OR PICKUP. 
========================================================================================================================================================================*/

function
SendConfirmation(e)
{
  var send      = false;
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheet     = ss.getActiveSheet();
  var stop      = sheet.getRange(e.range.getRow(), 1, 1, (Global().n_data_columns)).getValues()[0];
  var heading   = stop[Global().c_first_name - 1];
  var email     = stop[Global().c_email      - 1];
  var address   = stop[Global().c_address    - 1];
  var workorder;
  var status;
  var subject;
  var body;
  
  if (email == null || email == undefined) return;
  
  if(isLeftContainingRight(heading,"360BLUE WORKORDER"))
  {
    workorder = getWorkOrderNumber(heading);
    status    = getWorkOrderStatus(workorder);
    
    if      (status == "SCHEDULED") send = true;
    else if (status == -1)          send = false;
    else return;
    
    if (send)
    {
      subject = heading + " | COMPLETED";
      body    = "This work order has been signed off as complete.";
    }
    
    setWorkOrderStatus(workorder, address, email, "COMPLETED");
  }
  else if(
    isLeftContainingRight(heading,"DELIVERY")       ||
    isLeftContainingRight(heading,"TEMPORARY")      ||
    isLeftContainingRight(heading,"SWAP")           ||
    isLeftContainingRight(heading,"REPLACEMENT")    ||
    isLeftContainingRight(heading,"BIKE FIX")       ||
    isLeftContainingRight(heading,"PICKUP")         ||
    isLeftContainingRight(heading,"ITEM MOVE")      ||
    isLeftContainingRight(heading,"ABANDONED")      ||
    isLeftContainingRight(heading,"BIKE CHECK")     ||
    isLeftContainingRight(heading,"GART CHECK")     ||
    isLeftContainingRight(heading,"QUANTITY CHECK")
    )
  {
    // Standard, RGH Operations
  }
  else
  {
    // Order Delivered / Picked Up
  }
  
  if(send) GmailApp.sendEmail(email, subject, body);
  else     return;
}





/*========================================================================================================================================================================
  getWorkOrders()
  
  DESCRIPTION:
    PEER INTO GMAIL, LOOKING FOR 360BLUE WORK ORDERS.
    PARSE EACH OF THE WORK ORDERS TO ADHERE TO RGH'S ROUTE SYNTAX.
    ADD THE PARSED-ROUTE STOP TO THE BOTTOM OF TBA.
  
  NOTES:
  - Email subject format:
      360BLUE WORKORDER | {!WorkOrder.WorkOrderNumber} | {!WorkOrder.Type__c} | {!WorkOrder.Property__c}
    
  - Email body format:
      {!WorkOrder.Description}
========================================================================================================================================================================*/

function
getWorkOrders()
{
  // =============
  // = VARIABLES =
  // =============
  
  // CONSTANTS
  
  const SUBJECT = "360BLUE WORKORDER"; // SUBJECT CONTENT TO BASE OUR PARSE
  const BLANK   = "";
  
  // SHEET OBJECT
  
  var   sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Global().t_360_wo);
  
  // EMAILS OBJECT
  
  var   email   = GmailApp.search("subject: " + SUBJECT);   // LOOKING FOR: "360BLUE WORKORDER | {!WorkOrder.WorkOrderNumber} | {!WorkOrder.Type__c} | {!WorkOrder.Property__c}"
  
  // STRINGS
  
  var   subject;                                            // HOLDS THE ENTIRE SUBJECT
  var   components;                                         // HOLDS COMPONENETS DERIVED FROM THE SUBJECT
  var   title;       // 360BLUE WORKORDER                   // FROM SUBJECT: components, index 0
  var   workorder;   // {!WorkOrder.WorkOrderNumber}        // FROM SUBJECT: components, index 1
  var   property_id; // {!WorkOrder.Property__c}            // FROM SUBJECT: components, index 3
  var   agent_email; // {!WorkOrder.CreatedBy.Email}        // FROM SENDER:  getFrom()
  var   description  // {!WorkOrder.Description}            // FROM BODY:    getPlainBody()
  var   date            = getDate();                        // RETRIEVE STRING OF DATE IN MM/DD FORMAT
  
  // ARRAYS
  
  var   record           = getWorkOrderRecord();
  var   recordDetails    = [];
  var   stops            = [];
  
  // PLACEHOLDERS
  
  var   property;                                           // OBJECT WILL CONTAIN: .notes, .name, .area, .address, .type, .count
  
  // BOOLEAN
  
  var   newWorkOrders = false;
  
  
  // ==================
  // = FOR EACH EMAIL =
  // ==================
  
  for (var i = 0; i < email.length; ++i)
  { 
    subject     = email[i].getFirstMessageSubject();                    // RETRIEVE SUBJECT FROM EMAIL
    components  = subject.split(" | ");                                 // BREAK THE SUBJECT INTO ITS COMPONENTS
    
    title       = components[0];                                        // "360BLUE WORKORDER"
    workorder   = components[1];                                        // {!WorkOrder.WorkOrderNumber}
    property_id = components[3];                                        // {!WorkOrder.Property__c}
    
    
    // =====================
    // = VERIFY WORK ORDER =
    // =====================
    
    if      (title == null || workorder == null || property_id == null) {             continue; }
    else if (!isNewWorkOrder(workorder, record))                        {             continue; }
    else                                                                { newWorkOrders = true; }
    
    
    // =====================
    // = GET EMAIL CONTENT =
    // =====================
    
    agent_email = getEmailAddress(email[i].getMessages()[0].getFrom()); // GET EMAIL ADDRESS FROM THE SENDER -- SALESFORCE USES AGENT EMAIL AS SENDER
    
    description = email[i].getMessages()[0].getPlainBody();             // {!WorkOrder.Description}
    description = description.slice(0,-1);                              // REMOVES THE NEWLINE THAT IS ATTACHED TO THE END OF EVERY DESCRIPTION
    
    property    = getPropertyDetails(property_id);                      // GET PROPERTY DETAILS FROM THE PROPERTY CODE PROVIDED
    
    
    // =====================
    // = RECORD WORK ORDER =
    // =====================
    
    recordDetails.push(
      [
        workorder,
        property_id,
        property.address,
        agent_email,
        "SCHEDULED"
      ]
    );
    
    
    // ===============
    // = RECORD STOP =
    // ===============
  
    stops.push(
      [
        BLANK,
        BLANK,
        BLANK,
        BLANK,
        date,
        "360 Blue",
        title + "#" + workorder,
        "360 Blue",
        property.name,
        agent_email,
        property.area,
        property.address,
        BLANK,
        description + '\n' + "--" + '\n' + "HOUSE BIKES: (" + property.count + ") " + property.type,
        property.notes
      ]
    );
    
    
    // =====================
    // = SEND CONFIRMATION =
    // =====================
    
    email[i].replyAll("Thank you,\n\nWork Order #" + workorder + " has been scheduled!\nA separate email will be dispatched upon completion.");
  }
  
  // ==============
  // = DATA ENTRY =
  // ==============
  
  if (newWorkOrders)
  {
    sheet.getRange(sheet.getLastRow() + 1, 1, stops.length, stops[0].length).setValues(stops); // INSERT COLUMN DATA, AFTER THE LAST LINE OF DATA IN THE SPREADSHEET
    RecordWorkOrders(recordDetails);
  }
}





/*========================================================================================================================================================================
  getDate()
  
  DESCRIPTION:
    RETURN STRING OF THE DATE IN MM/DD FORMAT.
    
  NOTES:
    - getMonth() RETURNS THE MONTH NUMBER - 1, SO WE NEED TO ADD 1 TO THE RESULT.
    
    - getDate()  RETURNS THE DATE OF THE MONTH AS-IS, SO NO MANIPULATION IS NEEDED.
========================================================================================================================================================================*/

function
getDate()
{
  var date = new Date();
  
  return ("" + (date.getMonth() + 1) + '/' + date.getDate() + "");
}





/*
========================================================================================================================================================================
  getConvertedDate(date)
  
  DESCRIPTION:
    RETURN STRING OF THE DATE IN MM/DD FORMAT, FOR AN ALREADY-EXISTING DATE OBJECT.
    
  NOTES:
    - getMonth() RETURNS THE MONTH NUMBER - 1, SO WE NEED TO ADD 1 TO THE RESULT.
    
    - getDate()  RETURNS THE DATE OF THE MONTH AS-IS, SO NO MANIPULATION IS NEEDED.
========================================================================================================================================================================
*/

function
getConvertedDate(date)
{ 
  if (date == null || date == undefined || date == "") return;
  
  return ("" + (date.getMonth() + 1) + '/' + date.getDate() + "");
}





/*========================================================================================================================================================================
  getPropertyDetails(property_id)
  
  DESCRIPTION:
    PEER INTO THE BIKE PROGRAM SCHEDULE.
    QUERY THE DATA TO RETRIEVE DETAILS PERTAINING TO THE PROPERTY PROVIDED AS AN ARGUMENT (property_id).
    RETURN THE DETAILS IN THE FORM OF AN OBJECT.
  
  NOTES:
  - BIKE PROGRAM SCHEDULE SPREADSHEET ID:
      1JZ-JDk4_VuX6AmlifvjnSVA1b43BA65PcRL2-RpVfbY
  
  - DETAILS PERTAINING TO PROPERTY:
      PROPERTY NOTES, PROPERTY NAME, PROPERTY AREA, PROPERTY ADDRESS, BIKE TYPE, AND BIKE COUNT
========================================================================================================================================================================*/

function
getPropertyDetails(property_id)
{
  var spreadsheet_id = Global().s_bp_schedule;
  var query          = "=QUERY(SCHEDULE_RGH!A:K, \"SELECT D, F, G, H, J, K WHERE B = \'" + property_id + "\' LIMIT 1\", 0)";
  var schedule       = SpreadsheetApp.openById(spreadsheet_id);
  var temp           = schedule.insertSheet();
  
  temp.getRange(1,1).setFormula(query);
  
  var results = temp.getDataRange().getValues();

  schedule.deleteSheet(temp);
  
    
  return {
    "notes"  : results[0][0],
    "name"   : results[0][1],
    "area"   : results[0][2],
    "address": results[0][3],
    "type"   : results[0][4],
    "count"  : results[0][5]
  }
}





/*========================================================================================================================================================================
  getEmailAddress(from)
  
  DESCRIPTION:
    RETURN THE EMAIL ADDRESS FROM GMAIL'S getFrom() RESPONSE.
  
  NOTES:
  - FROM RESPONSE FORMAT:
      "FirstName LastName" <email@address.com>
  
  - PARSE REGULAR EXPRESSION SYNTAX SOURCE:
      https://ctrlq.org/code/19985-parse-email-address
      
  - REGULAR EXPRESSION BREAKDOWN:
      /  => OPEN               => INDICATES THE START OF A REGULAR EXPRESSION.
      [^ => NEGATED SET OPEN   => START: MATCH ANY CHARACTER THAT IS NOT IN THE SET.
      @  => CHARACTER          => MATCHES A "@" CHARACTER (CHAR CODE 64).
      <  => CHARACTER          => MATCHES A "<" CHARACTER (CHAR CODE 60).
      \s => WHITESPACE         => MATCHES ANY WHITESPACE CHARACTER (SPACES, TABS, LINE BREAKS).
      ]  => NEGATED SET CLOSES => END:   MATCH ANY CHARACTER THAT IS NOT IN THE SET.
      +  => QUANTIFIER         => MATCH 1 OR MORE OF THE PRECEDING TOKEN.
      @  => CHARACTER          => MATCHES A "@" CHARACTER (CHAR CODE 64).
      [^ => NEGATED SET OPEN   => START: MATCH ANY CHARACTER THAT IS NOT IN THE SET.
      @  => CHARACTER          => MATCHES A "@" CHARACTER (CHAR CODE 64).
      <  => CHARACTER          => MATCHES A "<" CHARACTER (CHAR CODE 60).
      \s => WHITESPACE         => MATCHES ANY WHITESPACE CHARACTER (SPACES, TABS, LINE BREAKS).
      >  => CHARACTER          => MATCHES A ">" CHARACTER (CHAR CODE 62).
      ]  => NEGATED SET CLOSES => END:   MATCH ANY CHARACTER THAT IS NOT IN THE SET.
      +  => QUANTIFIER         => MATCH 1 OR MORE OF THE PRECEDING TOKEN.
      /  => CLOSE              => INDICATES THE END OF A REGULAR EXPRESSION AND THE START OF EXPRESSION FLAGS.
========================================================================================================================================================================*/

function
getEmailAddress(from)
{
  return from.match(/[^@<\s]+@[^@\s>]+/)[0];
}





/*========================================================================================================================================================================
  getWorkOrderNumber(prompt)
  
  DESCRIPTION:
    RETURN THE WORK ORDER NUMBER FROM THE WORK ORDER STOP.
  
  NOTES:
  - FROM RESPONSE FORMAT:
      360BLUE WORKORDER  #WorkOrderNumber
  
  - PARSE REGULAR EXPRESSION SYNTAX SOURCE:
      https://stackoverflow.com/questions/4058923/how-can-i-use-regex-to-get-all-the-characters-after-a-specific-character-e-g-c
      
  - REGULAR EXPRESSION BREAKDOWN:
      /  => OPEN               => INDICATES THE START OF A REGULAR EXPRESSION.
      [^ => NEGATED SET OPEN   => START: MATCH ANY CHARACTER THAT IS NOT IN THE SET.
      @  => CHARACTER          => MATCHES A "#" CHARACTER (CHAR CODE 35).
      ]  => NEGATED SET CLOSES => END:   MATCH ANY CHARACTER THAT IS NOT IN THE SET.
      *  => QUANTIFIER         => MATCH 0 OR MORE OF THE PRECEDING TOKEN.
      $  => END                => MATCHES THE END OF THE STRING, OR THE END OF A LINE IF THE MULTILINE FLAG (m) IS ENABLED.
      /  => CLOSE              => INDICATES THE END OF A REGULAR EXPRESSION AND THE START OF EXPRESSION FLAGS.
========================================================================================================================================================================*/

function
getWorkOrderNumber(heading)
{
  return heading.match(/[^#]*$/)[0];
}





/*========================================================================================================================================================================
  isNewWorkOrder(workOrderNumber, record)
  
  DESCRIPTION:
    VERIFY IF A WORK ORDER HAS BEEN RECORDED PREVIOUSLY OR NOT.
    RETURN TRUE, IF THE WORK ORDER HAS NOT BEEN RECORDED;
    RETURN FALSE, IF THE WORK ORDER HAS BEEN RECORDED.
========================================================================================================================================================================*/

function
isNewWorkOrder(workOrderNumber, record)
{
  var result = true;
  
  for (var i = 0; i < record.length; ++i)
  {
    if (record[i][0] == workOrderNumber)
    {
      result = false;
      break;
    }
  }
  
  return result;
}





/*========================================================================================================================================================================
  recordWorkOrder(workOrderNumber,property_id, property_address)
  
  DESCRIPTION:
    RECORD THE WORK ORDER, PROPERTY ID, PROPERTY ADDRESS, AGENT EMAIL, AND WORK ORDER STATUS.
    THIS IS DONE IN AN ARCHIVE WORKSHEET, "360Blue Work Orders," IN THE "WORK ORDERS" TAB.
========================================================================================================================================================================*/

function
RecordWorkOrders(recordDetails)
{
  var ss    = SpreadsheetApp.openById(Global().s_wos_register);
  var sheet = ss.getSheetByName(Global().t_workorders);
  
  sheet.getRange(sheet.getLastRow() + 1, 1, recordDetails.length, recordDetails[0].length).setValues(recordDetails); // INSERT COLUMN DATA, AFTER THE LAST LINE OF DATA IN THE SPREADSHEET
}





/*========================================================================================================================================================================
  getWorkOrderRecord()
  
  DESCRIPTION:
    PULL THE WORK ORDER RECORD, PREVIOUSLY SUBMITTED BY RecordWorkOrders.
========================================================================================================================================================================*/

function
getWorkOrderRecord()
{
  var ss    = SpreadsheetApp.openById(Global().s_wos_register);
  var sheet = ss.getSheetByName(Global().t_workorders);
  
  return (sheet.getDataRange().getValues());
}





/*========================================================================================================================================================================
  getWorkOrderStatus(workorder)
  
  DESCRIPTION:
    PULL THE WORK ORDER STATUS, FROM A PARTICULAR WORK ORDER.
========================================================================================================================================================================*/

function
getWorkOrderStatus(workorder)
{
  var ss    = SpreadsheetApp.openById(Global().s_wos_register);
  var sheet = ss.getSheetByName(Global().t_workorders);
  var data  = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; ++i)
  {
    if (data[i][0] == workorder)
    {
      return data[i][4];
    }
  }
  
  return -1;
}





/*========================================================================================================================================================================
  setWorkOrderStatus(workorder, address, email, status)
  
  DESCRIPTION:
    SET THE WORK ORDER STATUS, FOR A PARTICULAR WORK ORDER.
========================================================================================================================================================================*/

function
setWorkOrderStatus(workorder, address, email, status)
{ 
  var ss    = SpreadsheetApp.openById(Global().s_wos_register);
  var sheet = ss.getSheetByName(Global().t_workorders);
  var data  = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; ++i)
  {
    if (data[i][0] == workorder)
    {
      sheet.getRange(i+1,5).setValue(status);
      return;
    }
  }
  
  sheet.getRange(sheet.getLastRow()+1,1,1,5).setValues([[workorder, "", address, email, status]]);
}
