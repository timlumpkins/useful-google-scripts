//These are the sheets that are being written to 
var antiSheet = SpreadsheetApp.getActive().getSheetByName("Anti-Reproductive Health Bills");
var supportedRawSheet = SpreadsheetApp.getActive().getSheetByName("Supported Bills");
var broadRawSheet = SpreadsheetApp.getActive().getSheetByName("Broader Legislation Opposed");

//These are the Internal Bill Summary's lists
  var inSumAnti = SpreadsheetApp.openByUrl('your-sheets-url').getSheets()[0];
  var inSumSupported = SpreadsheetApp.openByUrl('your-sheets-url').getSheets()[1];
  var inSumBroad = SpreadsheetApp.openByUrl('your-sheets-url').getSheets()[2];

// The onOpen function is executed automatically every time a Spreadsheet is loaded
 function onOpen() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var menuEntries = [];
   //menuEntries.push({name: "Refresh the anti list", functionName: "updateAnti"});
   //menuEntries.push(null); // line separator
   menuEntries.push({name: "Refresh all lists", functionName: "updateAll"});

   ss.addMenu("PPMO Update", menuEntries);
 }

function setColumnSizes(){
  var curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colWidths = [75, 125, 85, 175, 475, 85];
  for (var c = 1; c < 7; c++){
//   Logger.log("Column " + c + " is " + curSheet.getColumnWidth(c));
    if (curSheet.getColumnWidth(c) !== colWidths[c-1]){
      curSheet.setColumnWidth(c,colWidths[c-1]);
    }
  }  
}

function fontFormatting(curSheet){
  //var curSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //Setting the ranges for the columns
  var bigText = curSheet.getRange("A:E");
  var smallText = curSheet.getRange("F:F");
  
  //Setting the font style and sizes
  
  curSheet.getRange("A:F").setFontFamily("Arial");
  bigText.setFontSize(11);
  smallText.setFontSize(10);
  
  //Setting Header size, weight and background
  curSheet.getRange("A1:F1").setFontSize(12);
  curSheet.getRange("A1:F1").setFontWeight("bold");
  curSheet.getRange("A1:F1").setBackground("#ec008c");
  curSheet.getRange("A1:F1").setFontColor("white");
    curSheet.setRowHeight(1, 50);
  
  //Setting cell alignment and wrapping
  curSheet.getRange("A:D").setHorizontalAlignment("center");
  curSheet.getRange("E2:E100").setHorizontalAlignment("left");
  curSheet.getRange("F:F").setHorizontalAlignment("center");
  curSheet.getRange("E1").setHorizontalAlignment("center");
  curSheet.getRange("A:F").setWrap(true);
  curSheet.getRange("A:F").setVerticalAlignment("middle"); 
  
  //Sort everything!
  //curSheet.sort(1);
  
  //Freeze the first row
  curSheet.setFrozenRows(1);
  
  //sets column widths
  setColumnSizes();
}

function updateAll(){
  pullBills(inSumAnti, antiSheet); 
  pullBills(inSumSupported, supportedRawSheet);
  pullBills(inSumBroad, broadRawSheet); 
  
  fontFormatting(antiSheet);
  fontFormatting(supportedRawSheet);
  fontFormatting(broadRawSheet);
}

function pullBills(origSheet, newSheet){
  newSheet.getRange("A:F").clearContent();
  var rownumber = origSheet.getLastRow();
  var extdata =  origSheet.getRange(2, 3, rownumber, 1).getValues();
  var extlinks = origSheet.getRange(2, 1, rownumber, 2).getFormulas();
  var col3values = origSheet.getRange(2, 4, rownumber, 5).getValues();
  var col3formulas = origSheet.getRange(2, 4, rownumber, 5).getFormulas();
  var extdatacolors = origSheet.getRange(2, 1, rownumber, 6).getBackgrounds();
  var extdatanew = [];
  var extlinksnew = [];
  var col3final = [];
  
  for(var row in extdata){
    if (extdatacolors[row][0] == "#ffffff" || extdatacolors[row][0] == "#ead1dc" ){
      extdatanew.push(extdata[row]);      
      extlinksnew.push(extlinks[row]);
      
      if (col3formulas[row][0] == ""){
        col3final.push(col3values[row]);
        Logger.log(col3values[row]);
      }else{      
        col3final.push(col3formulas[row]);
      }
    }
  }
  
  //Set new final row number
  rownumber = col3final.length;
  
  newSheet.getRange(2, 3, rownumber, 1).setValues(extdatanew);
  newSheet.getRange(2, 1, rownumber, 2).setFormulas(extlinksnew);
  newSheet.getRange(2, 4, rownumber, 5).setValues(col3final);
  newSheet.getRange("A1:F1").setValues(origSheet.getRange("A1:F1").getValues());
}

function updateAnti(){
  pullBills(inSumAnti, antiSheet);   
  fontFormatting(antiSheet);
}
