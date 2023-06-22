function add() {

  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('sheet2'), true);
//  var hiderange = spreadsheet.getRange('B:C');
//  spreadsheet.showColumns(2,2)
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('graph'), true);
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .setChartType(Charts.ChartType.ORG)
  .addRange(spreadsheet.getRange('Sheet2!B2:E'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('width', 776)
  .setPosition(2, 2, 36, 3)
  .build();
  sheet.insertChart(chart);
//  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet1'), true);
//  spreadsheet.hideColumn(range);
//  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('graph'), true);
};

function delete1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('graph'), true);
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet2'), true);
};
function drawChart() {
      var spreadsheet = SpreadsheetApp.getActive();
      //  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('sheet1'), true);
      //  var hiderange = spreadsheet.getRange('B:C');
      //  spreadsheet.showColumns(2,2)
        spreadsheet.setActiveSheet(spreadsheet.getSheetByName('graph'), true);
        var sheet = spreadsheet.getActiveSheet();
        var chart = new OrgChart(document.getElementById('chart_div'));
          var data = new DataTable();
        data.addRange(spreadsheet.getRange('Sheet1!B2:E'))
        setOption('bubble.stroke', '#000000')
        .setOption('width', 776)
        .setPosition(2, 2, 36, 3)
        var data = Charts.newDataTable()
//        data.addColumn('string', 'Name');
//        data.addColumn('string', 'Manager');
//        data.addColumn('string', 'ToolTip');
//
//        // For each orgchart box, provide the name, manager, and tooltip to show.
//        data.addRows([
//          [{'v':'Mike', 'f':'Mike<div style="color:red; font-style:italic">President</div>'},
//           '', 'The President'],
//          [{'v':'Jim', 'f':'Jim<div style="color:red; font-style:italic">Vice President</div>'},
//           'Mike', 'VP'],
//          ['Alice', 'Mike', ''],
//          ['Bob', 'Jim', 'Bob Sponge'],
//          ['Carol', 'Bob', '']
//        ]);

        // Create the chart.
        var chart = new OrgChart(document.getElementById('chart_div'));
        // Draw the chart, setting the allowHtml option to true for the tooltips.
        sheet.insertChart(chart);
//        chart.draw(data, {'allowHtml':true});
      }


function Untitledmacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2:B9').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('graph'), true);
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  
  .setChartType(Charts.ChartType.ORG)
  .addRange(spreadsheet.getRange('Sheet1!A2:E1000'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setPosition(3, 2, 79, 8)
  .build();
  sheet.insertChart(chart);
};