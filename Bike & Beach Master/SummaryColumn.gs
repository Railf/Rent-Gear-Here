/*===========================================================
  SummaryColumn.gs
  ----------------
  
  Utilizes the item data for each stop,
  to create a summary of the items needed.
  
  
  LEGEND:
  
    CreateSummary()
      => returns order summary for each row in the sheet
         invoked by the Summary tab, created with inOpen()
    
    onOpen()
      => creates "Summary" tab in menu bar of window
    
    onEdit(e)
      => creates the summary for the row when an edit is made
  
  
  USE EXAMPLE:
  
    When an edit is made to an order (a row),
    the summary is updated with each edit.
    
    Alternatively, you can manually update each row,
    by navigating to the Summary menu, and clicking,
    "Generate Summaries."
=============================================================
Ralph McCracken, III */



function CreateSummary() {
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = Spreadsheet.getName();
  
  if (name != "Bike/Beach Master" && name != "In Progress DELIVERY" && name != "In Progress PICK UP")
    return;
  
  var Rows        = Spreadsheet.getLastRow();                    // Get number of non-empty rows.
  var Columns     = Spreadsheet.getLastColumn();                 // Get number of non-empty columns.
  
  var HeaderRow   = 1;                                           // Static:   Row number for header information.                => ItemName
  var SummColumn  = 16;                                          // Static:   Column of location to store summary.              => EditCell
  
  var CountRow    = 2;                                           // Variable: Row number for item quantity (count) information. => EditCell, ItemCount
  var ItemColumn  = 17;                                          // Variable: Column number for item header (name).             => ItemName, ItemCount
  
  var EditCell    = Spreadsheet.getRange(CountRow, SummColumn);  // Cell final Summary will go into.
  
  var ItemName    = Spreadsheet.getRange(HeaderRow, ItemColumn); // Item header in doc, for the item name.
  var ItemCount   = Spreadsheet.getRange(CountRow, ItemColumn);  // Quantity data in doc.
  var Summary     = '';                                          // String that will be added to until summar is complete, then stored in EditCell.
  
  var Newline     = false;
  
  for (var j = 2; j <= Rows; ++j)
  {
    EditCell = Spreadsheet.getRange(j, SummColumn);
    Summary  = '';
    Newline  = false;
    
    // Create Summary of the current row.
    for (var i = 17; i <= Columns; ++i)
    {
      ItemName  = Spreadsheet.getRange(HeaderRow, i);
      ItemCount = Spreadsheet.getRange(j, i);
    
      if(ItemCount.isBlank())
        continue;

      if (Newline == false)
      {
        Summary += '(' + ItemCount.getValue() + ") " + ItemName.getValue();
        Newline = true;
      }
      else
        Summary += '\n' + '(' + ItemCount.getValue() + ") " + ItemName.getValue();
      
      continue;
    }
    
    if (Summary != '')
    {
      EditCell.setValue(Summary);
      EditCell.setBackground("yellow");
    }
    else
    {
      EditCell.setValue(Summary);
      EditCell.setBackground("white");
    }
  }
}


function
onOpen()
{
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Summary")
  .addItem('Generate Summaries', 'CreateSummary')
  .addToUi();
}



function onEdit(e)
{
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = Spreadsheet.getName();
  
  if (name != "Bike/Beach Master" && name != "In Progress DELIVERY" && name != "In Progress PICK UP")
    return;
  
  var range = e.range;
  var row = range.getRow();
  var column = range.getColumn();
  
  if (row == 1)
    return;
  
  var Columns     = Spreadsheet.getLastColumn();                 // Get number of non-empty columns.
  var HeaderRow = 1;
  var SumColumn = 16;
  
  Logger.log(row + ',' + column);
  
  var Newline = false;
  
  var ItemName;
  var ItemCount;
  
  var Summary = '';
  var EditCell = Spreadsheet.getRange(row, SumColumn);
  
  // Create Summary of the current row.
  for (var i = 17; i <= Columns; ++i)
  {
    ItemName  = Spreadsheet.getRange(HeaderRow, i);
    ItemCount = Spreadsheet.getRange(row, i);
    
      if(ItemCount.isBlank())
        continue;

    if (Newline == false)
    {
      Summary += '(' + ItemCount.getValue() + ") " + ItemName.getValue();
      Newline = true;
    }
    else
      Summary += '\n' + '(' + ItemCount.getValue() + ") " + ItemName.getValue();
      
    continue;
  }
    
  if (Summary != '')
  {
    EditCell.setValue(Summary);
    EditCell.setBackground("yellow");
  }
  else
  {
    EditCell.setValue(Summary);
    EditCell.setBackground("white");
  }
}
