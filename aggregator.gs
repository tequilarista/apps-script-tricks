//
// These functions are useful if you have a doc with many tabs within a spreadsheet 
// where data is structured the same in each tab and you want a simple way to 
// aggregate them (that avoids having to wrangle a bunch of macros)
//

// Get a list of all the tabs in a sheets doc

function getAllSheetNames() {
  var dataDump = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) 
    dataDump.push( sheets[i].getName() )
  Logger.log(dataDump);
  return dataDump;
}

//
// Grabs all the data from a particular sheet tab, and return it in an array. 
//
// This function assumes these scripts will get invoked from a sheet tab named "Aggregate", 
// so that name is excluded, rename as appropriate.

function getSheetData(SheetName) {
  if (SheetName == "Aggregate") { // exclude (at a minimum) the tab you're
    return
  } else if (sheet.isSheetHidden()) { 
    Logger.log(sheet.getSheetName() + ": HIDDEN")
    return 
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SheetName);
  Logger.log(sheet.getSheetName())
  rows = sheet.getDataRange().getValues();
  if (rows.length == 0) {return} // don't bother with empty sheets
  rows.shift(); // exclude the first row (column headers) of each sheet, usually the header
  
  //  uncomment this section to insert the sheet tab name at the beginning of each row 
  //  so that it can be queried against:
   
  /* 
  rows.forEach( 
  //  function(value) {
  //    value.unshift(SheetName) 
    }); 
  */
  Logger.log(rows)
  return rows;
}

// This is the function to actually invoke in an empty sheet via: =getAllSheetData()
// For each sheet tab in a doc, grab and concatenate all the data into a single sheet

function getAllSheetData() {
  var aggDump = new Array();
  var slist = getAllSheetNames();
  for (var i=0; i<slist.length; i++) 
    aggDump = aggDump.concat(getSheetData(slist[i]));
  Logger.log(aggDump)
  return aggDump;
}
