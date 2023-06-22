function onOpen() {
  var menu = [{name: 'Create graphy', functionName: 'doIt'}];
  SpreadsheetApp.getActive().addMenu('show', menu);
}
function doIt(){
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html,'Org Chart');
}
function dataZoom() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
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

function macro1() {
  var spreadsheet = SpreadsheetApp.getActive();
  var newSheet = spreadsheet.getSheetByName('chart');
  if(newSheet != null){
    spreadsheet.deleteSheet(newSheet);
  }
  spreadsheet.insertSheet();
  spreadsheet.getActiveSheet().setName('chart');
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .setChartType(Charts.ChartType.ORG)
  .addRange(spreadsheet.getRange('Sheet2!A2:A'))
  .addRange(spreadsheet.getRange('Sheet2!E2:E'))
  .addRange(spreadsheet.getRange('Sheet2!B2:B'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('curveType', 'none')
  .setOption('color', '#4285f4')
  .setOption('selectionColor', '#f8f1a8')
  .setOption('size', 'medium')
  .setOption('domainAxis.direction', 1)
  .setOption('height', 458)
  .setOption('width', 741)
  .setPosition(2, 3, 111, 1)
  .build();
  sheet.insertChart(chart);
//  sheet = spreadsheet.moveChartToObjectSheet(chart);
};
