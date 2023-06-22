function onOpen() {
  var menu = [{name: 'Create graphy', functionName: 'doIt'}];
  SpreadsheetApp.getActive().addMenu('show', menu);
}
function doIt(){
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(800)
    .setHeight(600);
    
  var spreadsheet = SpreadsheetApp.getActive();
  
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html,'Org Chart');
}
function dataZoom() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
//  var range = sheet.getDataRange();
  var values = sheet.getRange("A:F").getValues()
  var allData =[];
  for(var i = 0; i < values.length; i++){
    allData[i-1] = [];
    var num = 0;
    for(var j = 0; j < values[0].length; j++){
      allData[i - 1][num] = values[i][j];
      num = num + 1;
    }
  }
  return allData;
}


