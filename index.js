function orderSummary() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sum = 0;
  var productType = new Array();
  var at = sheets.length;
  
  for (var i = 2; i < sheets.length ; i++ ) {
    var sheet = sheets[i];
    var lastRow = 20+getFirstEmptyRow(sheet);
    var productRange = sheet.getRange('A21:A'+lastRow);
    var qtyRange = sheet.getRange('E21:E'+ lastRow);
    var numRows = productRange.getNumRows();
    var numCols = productRange.getNumColumns();
    
    for (var ri = 1; ri <= numRows; ri++) {
      for (var rj = 1; rj <= numCols; rj++) {
        var product = productRange.getCell(ri,rj).getValue().toUpperCase();        
        var qty = qtyRange.getCell(ri, rj).getValue();
        
        if (product in productType){
          productType[product] += qty; 
        }
        else{
          productType[product] = qty; 
         
        }   
      }
    } 
  }
  
  
  var invSheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var invRow = 4;
  for (var key in productType) {
    invSheet.getRange("C"+ invRow).setValue(key);
    invSheet.getRange("G"+ invRow).setValue(productType[key]);
    invRow ++;
    //console.log("key " + key + " has value " + productType[key]);
  }
}


function getFirstEmptyRow(sheet) {
  var spr = sheet;
  var column = spr.getRange('A21:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct][0] != "" ) {
    ct++;
  }
  return (ct);
}



function createOrder()
{

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('template').copyTo(ss);
  var orderNum = ss.getSheetByName('template').getRange("G3").getValue() + 1;
  ss.getSheetByName('template').getRange("G3").setValue(orderNum);
  
  //SpreadsheetApp.flush(); // Utilities.sleep(2000);
  sheet.setName(orderNum);
  sheet.getRange("G3").setValue(orderNum);
  var dToday = new Date();
  sheet.getRange("G2").setValue(dToday);
  
  /* Make the new sheet active */
  ss.setActiveSheet(sheet);
  
  
}
