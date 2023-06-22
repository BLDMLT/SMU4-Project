//A FUNCTION TO CREATE A GRAPH IN THE GOOGLE SHEET
function createGraphYenYou()
{
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var activeSheet = ss.getSheetByName("Maptek");
  
  var data = activeSheet.getRange("A2:B30");
  
  var chart = activeSheet.newChart();
  var orgChart = chart.setChartType(Charts.ChartType.ORG)
  .addRange(data)
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setPosition(4, 4, 0, 0)
  .setOption('width', 2200)
  .build();
  
  activeSheet.insertChart(orgChart);

}

//A FUNCTION TO REMOVE THE ORGANISATIONAL CHART IN THE GOOGLE SHEET
function removeGraphYenYou()
{
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var activeSheet = ss.getSheetByName("Maptek");
  
  var chart = activeSheet.getCharts();
  chart = chart[chart.length - 1]; //What does this do?
  activeSheet.removeChart(chart);
}

//SHOW A POPUP TO ADD NEW EMPLOYEE
function showModelYenYou()
{
  var html = HtmlService.createHtmlOutputFromFile("addEmployee").setWidth(800).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Add New Employee");
}

//SHOW A POPUP TO SEARCH EMPLOYEE
function showSearchYenYou()
{
  var html = HtmlService.createHtmlOutputFromFile("searchFunction").setWidth(800).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Search Employee");
}

//THIS IS AN ALTERNATIVE, ADD EMPLOYEE FROM ANOTHER WEB
function doGet()
{
  return HtmlService.createHtmlOutputFromFile("addEmployee");
}

//A FUNCTION TO ADD EMPLOYEE INTO THE GOOGLE SHEET
function addEmployeeYenYou(employeeInfo)
{
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var activeSheet = ss.getSheetByName("Maptek");
 
  activeSheet.appendRow([employeeInfo.firstName + " " + employeeInfo.lastName, new Date(), employeeInfo.age, employeeInfo.position]);
}

function getPosition(nameID)
{
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var activeSheet = ss.getSheetByName("Maptek");
  var data = activeSheet.getRange(38, 1, activeSheet.getLastRow()-37, 4).getValues();
  
  var nameList = data.map(function(r){return r[0];});
  var positionList = data.map(function(r){return r[3];});
  
  var positionName = nameList.indexOf(nameID);

  if(positionName > -1)
  {
    return positionList[positionName];
  }
  else
  {
    return "Not Found!";
  }
}