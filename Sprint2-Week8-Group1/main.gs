// Create menu
function onOpen() {
  var menu = [{name: 'Create graphy', functionName: 'doIt'}];
  SpreadsheetApp.getActive().addMenu('show', menu);
}

// Create Html page
function doIt(){



  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(800)
    .setHeight(600);
 
  var spreadsheet = SpreadsheetApp.getActive();
  
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html,'Org Chart');
      
}

// Get Data from spreadsheet following the order：
// 0-Team	1-First Name 2-Last Name 3-Full Name 4-Position 5-Superior after 6-tooltips
function dataZoom() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var blank = 0;
  // Only 26 column can be identifed
  var values = sheet.getRange("A:Z").getValues()
  var allData =[];
  for(var i = 0; i < values.length; i++){
    allData[i] = [];
    var num = 6;
    for(var j = 0; j < values[0].length; j++){
      if(values[0][j] == 'Team'){
        allData[i][0] = values[i][j];
      }
      else if(values[0][j] == 'First Name'){
        allData[i][1] = values[i][j];
      }
      else if(values[0][j] == 'Last Name'){
        allData[i][2] = values[i][j];
      }
      else if(values[0][j] == 'Full Name'){
        allData[i][3] = values[i][j];
      }
      else if(values[0][j] == 'Position'){
        allData[i][4] = values[i][j];
      }
      else if(values[0][j] == 'Superior'){
        allData[i][5] = values[i][j];
      }
      else{
        allData[i][num] = values[i][j];
        num ++;
      }
    }
  }
  return allData;
}
