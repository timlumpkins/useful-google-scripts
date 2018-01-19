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
   menuEntries.push({name: "Refresh the anti list", functionName: "updateAnti"});
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
  var smallText = curSheet.getRange("F:H");
  
  //Setting the font style and sizes
  
  curSheet.getRange("A:H").setFontFamily("Arial");
  curSheet.getRange("A2:H").setFontColor("").setFontWeight("normal").setBackground("white");
  bigText.setFontSize(11);
  smallText.setFontSize(10);
  
  //Setting Header size, weight and background
  curSheet.getRange("A1:H1").setFontSize(12).setFontWeight("bold");
  //curSheet.getRange("A1:F1").setFontWeight("bold");
  curSheet.getRange("A1:H1").setBackground("#ec008c");
  curSheet.getRange("A1:H1").setFontColor("white");
  curSheet.setRowHeight(1, 50);
  
  //Setting cell alignment and wrapping
  curSheet.getRange("A:D").setHorizontalAlignment("center");
  curSheet.getRange("E2:E100").setHorizontalAlignment("left");
  curSheet.getRange("F:H").setHorizontalAlignment("center");
  curSheet.getRange("E1").setHorizontalAlignment("center");
  curSheet.getRange("A:H").setWrap(true);
  curSheet.getRange("A:H").setVerticalAlignment("middle"); 
  
  //Sort everything!
  curSheet.sort(1);
  
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
  newSheet.getRange("A:H").clearContent();
  
  //Grabs all of the data from the sheet
  var fullRange = origSheet.getDataRange();
  var formulas = fullRange.getFormulas();
  var filteredRange = ["Bill Number", "Sponsor", "District", "Title", "Summary", "Last Action", "Co-sponsors", "Raw Bill #"];

  for(var row=2; row<fullRange.getLastRow(); row++){
    if (fullRange.getCell(row,1).getBackground() == "#ffffff" || fullRange.getCell(row,1).getBackground() == "#d9d9d9" ){ //Checks for white or grey background
      for(var col=0; col<8; col++){
        var formula = formulas[row-1][col];
        var value = fullRange.getValues()[row-1][col];
        
        if (formula && (col==0 || col==1 || col==3)){
          filteredRange.push(formula);
          }else{
          filteredRange.push(value);
          }
        }
    }
  }
  var finalBills = TwoDimensional(filteredRange,8);
  var rownumber = finalBills.length;

  newSheet.getRange(1, 1, rownumber, 8).setValues(finalBills);
}

function updateAnti(){
  pullBills(inSumAnti, antiSheet);   
  fontFormatting(antiSheet);
}

function TwoDimensional(arr, size){
  var res = []; 
  for(var i=0;i < arr.length;i = i+size)
    res.push(arr.slice(i,i+size));
  return res;
}
