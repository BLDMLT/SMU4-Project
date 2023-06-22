
function onEdit(e){
  var sheetName = e.source.getActiveSheet();
  var sheets = ['Maptek'];
  if(sheetName.getName() == "Maptek"){
    createOperationsMenu();
    updateTemplate();
    Logger.log("OnEdit Running");
  }
}

//Custom parameters for search operations 
function getCustomParameters(){
  var customPara = [
//    [Label            :     WorkingColumn],
//    ["DateOfJoiningAfter", "DateOfJoining"],
//    ["DateOfJoiningBefore" ,"DateOfJoining"],
//    ["ToDateAfter", "ToDate"],
//    ["ToDateBefore","ToDate"],
//    ["SaleryGreaterThan","Salary"],
//    ["SaleryLowerThan","Salary"],
//    ["RatingGreaterThan","Rating"],
//    ["RatingLowerThan","Rating"],
//    ["DateOfBirthAfter","DateOfBirth"],
//    ["DateOfBirthBefore","DateOfBirth"],
  ];
   
    return customPara;
}

// Function returns the working Column for a given CustomParameter x 
function findCustomParaWorkVariable(x){
  var customPara = getCustomParameters();
    for(let v in customPara){
      if(customPara[v][0] === x){
        return customPara[v][1];
      }
  }
  return "NONE";
}
    
//return True if x is a Custom Parameter, else returns False
function isACustomParam(x){
  var customPara = getCustomParameters();
  for(let v in customPara){
    if(customPara[v][0] === x){
      return true;
    }
  }
  return false;
}

//function is run by clicking search button
function searchDynamic(){
  getInformation();
}
    
function getInformation(){
    
  Logger.log("GetInformation")
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var information = ss.getSheetByName("Maptek");
  var lastColumn = information.getLastColumn();
  var lastRow = information.getLastRow();
  
  var searchPara = getSearchPara();
  var info = information.getRange(2,1,lastRow,lastColumn).getValues();
  var searchResult = info;
  Logger.log("Search Starting");
  for(let v of searchPara){
    var key = v[0];
    var value = v[1];
    Logger.log("Key->",key,"Value",value);
    var column_by_index = find_column_index(v[0]);

    if(value !==""){
      searchResult = searchResult.filter(function (item){
       
        if(isACustomParam( key ) === true){
          var workVariable = findCustomParaWorkVariable(key);
          column_by_index = find_column_index(workVariable);
          if(key.indexOf("After") >=0 || key.indexOf("GreaterThan") >=0 ){
            if(item[column_by_index] >= value ){
              return true;
            }else{
              return false;
            };
            
          }else{
            if(item[column_by_index] < value ){
              return true;
            }else{
              return false;
            };
          }
        }//custom parameter ends
        else if(typeof(value) === 'number' ){
          if(item[column_by_index] === value ){
            return true;
          }
          else{
            return false;
          };
        }
        else if(key.indexOf("Date") >=0){
          Logger.log("Date->",typeof(value));
          if(item[column_by_index].toString().indexOf(value.toString()) >=0 ){
            return true;
          }
          else{
            return false;
          };
        }
        else{
          var itemLowerCase = item[column_by_index].toString().toLowerCase();
          var valueLowerCase = value.toLowerCase();
          if(key == "Sex"){
            if(itemLowerCase.indexOf(valueLowerCase) ===0 ){
              return true;
            }
            else{
              return false;
            };
          }else{
            if(itemLowerCase.indexOf(valueLowerCase) >=0 ){
              return true;
            }else{
              return false;
            };
          }
        }

      

        
      });//filter ends
      
    }  //  if(v[1].toString() !== "") ends
    
  }//for ends

  if(searchResult.length != 0){
    writeToSheet(searchResult);
  }else{
    Logger.log("No Result Found")

  }
}

//function to write search results on Results Sheet
function writeToSheet(data){
  
  Logger.log("Writing To Result Sheet")
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var results = ss.getSheetByName("SearchResults");
  results.clearContents();
  var columnTitles = getInfoColumnTitles();
  columnTitles = [columnTitles];
  
  var lastColumn = data[0].length;
  var lastRow = data.length;
  results.getRange(1, 1, 1, lastColumn).setValues(columnTitles);
  results.getRange(2, 1, lastRow, lastColumn).setValues(data);
}

//function returns the column index of item
function find_column_index(item){
  var column_titles = getColumnTitles();
  for(x in column_titles){
    if(column_titles[x]===item){
      return parseInt(x);
    }
  }
}

//function to create search operations menu on operation sheet
function createOperationsMenu(){
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var information = ss.getSheetByName("Maptek");
  var infoLastColumn = information.getLastColumn();
  var infoLastRow = information.getLastRow();
  var operationsDynamic = ss.getSheetByName("SearchOperation");
  var columnTitles = getColumnTitles();
  var lastRow = columnTitles.length;
  operationsDynamic.getRange(1, 1,lastRow+10,2).clear();
  operationsDynamic.clear();
  for(let i=0;i<lastRow;i++){
    operationsDynamic.getRange(i+1, 1).setValue(columnTitles[i]);
    if(columnTitles[i]=="Position" || columnTitles[i]=="Sex"  || columnTitles[i]=="Team" || columnTitles[i]=="Superior"){
      var validCell = operationsDynamic.getRange(i+1, 2);
      validCell.clearDataValidations();
      var validRange = information.getRange(2, i+1, infoLastRow);
      var customRule = app.newDataValidation().requireValueInRange(validRange).build();
      validCell.setDataValidation(customRule);
    }
    
  }
  
  var myrange = operationsDynamic.getRange(1, 1,lastRow, 2);
  myrange.setBorder(true, true, true,true,true,true);
  Logger.log("Created Menu For operations.");
}

//function to get all columns, column label from information sheet and custom parameters
function getColumnTitles(){
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var information = ss.getSheetByName("Maptek");
  var lastColumn = information.getLastColumn();
  var lastRow = information.getLastRow();
  var column_title = information.getRange(1,1,1,lastColumn).getValues();
  for(let v in column_title[0]){
    column_title[0][v] = column_title[0][v].replace(/ /g,"");
  }
  column_title2 = column_title[0];
  var customPara = getCustomParameters();
  for(let x in customPara){
    column_title2.push(customPara[x]);
  }
  return column_title2;
}

//function to get labels only from information sheet
function getInfoColumnTitles(){
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var information = ss.getSheetByName("Maptek");
  var lastColumn = information.getLastColumn();
  var lastRow = information.getLastRow();
  var column_title = information.getRange(1,1,1,lastColumn).getValues();
  for(let v in column_title[0]){
    column_title[0][v] = column_title[0][v].replace(/ /g,"");
  }
  column_title2 = column_title[0];
  return column_title2;
}

//function to get search values from operations menu
function getSearchPara(){
  
  var searchPara = new Map();
  var app = SpreadsheetApp;
  var operationsDynamic = app.getActiveSpreadsheet().getSheetByName("SearchOperation"); 
  var columnTitles = getColumnTitles();
  var lastRow = columnTitles.length;  
  info = operationsDynamic.getRange(1, 1, lastRow,2).getValues();
  for(let x in info){
    searchPara.set(info[x][0],info[x][1]);
  }
  return searchPara;
}

//email function, sends email only to search results
function sendEmail(){
  
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var results = ss.getSheetByName("SearchResults");
  var lastRow = results.getLastRow();
  var lastColumn = results.getLastColumn();
  var emailIndex = find_column_index("Email");
  var resultsInfo = results.getRange(2, 1, lastRow-1, lastColumn).getValues();
  var recipientId = results.getRange(2,emailIndex+1, lastRow-1, 1).getValues();
  var app2 = SpreadsheetApp;
  var ss2 = app.getActiveSpreadsheet();
  var template = ss.getSheetByName("EmailTemplate");
  var mailSubject = template.getRange("B1").getValue();
  var keys = getTemplateKeys();
  
  var mailBody = template.getRange("B2").getValue();
  
  var columnTitles = getInfoColumnTitles();
  var templateKeys = getTemplateKeys();
  var id =  find_column_index("Email");
  
  for(let x in resultsInfo ){
    var mailCopy = mailBody;
    var editedMail = updateMailBody(mailCopy,resultsInfo[x],columnTitles);
    Logger.log("Sending Mail to ",resultsInfo[x][emailIndex]);
    MailApp.sendEmail(resultsInfo[x][emailIndex], mailSubject, editedMail);
    Logger.log("Sent Mail",editedMail);
  }
  
}

//function to replace {keys} with values in mail function
function updateMailBody(mail,x,keys){
  
  for(let v in keys){
    var index = find_column_index(keys[v]);
    var str = "{" + keys[v].toString() + "}";
    mail = mail.replace(str,x[index]);
  }
  return mail;
}

//function to generate keys on template sheet, keys are columnsof information sheet
function updateTemplate(){
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet()
  var template = ss.getSheetByName("EmailTemplate");
  var keys = getColumnTitles();
  var customPara = getCustomParameters();
  var rows = keys.length - customPara.length;
  template.getRange(4, 1,rows+4,2).clearContent();
  for(var i = 1;i<=rows;i++){
    template.getRange(3+i,1).setValue("Key");
    template.getRange(3+i,2).setValue("{"+keys[i-1]+"}");
  }
}

//function to read all the keys from template sheet
function getTemplateKeys(){
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet()
  var template = ss.getSheetByName("EmailTemplate");
  var columnTitle = getColumnTitles();
  var rows = columnTitle.length;
  var templateKeys = template.getRange(4, 2,rows,2).getValues();
  return templateKeys;
}


