function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('SMU')
       .addItem('Org Chart', 'doIt')
      .addSeparator()
      .addItem('SMU Report', 'printall')
      .addSubMenu(ui.createMenu('Report List')
          .addItem('People-Building Report', 'print1')
          .addItem('Team-Building Report', 'print2')
          .addItem('People-Team Report', 'print3')
          .addItem('Ratio Report', 'print4'))
      .addToUi();
  
  
}

// Create Html page
function doIt(){
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(1200)
  .setHeight(900);
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
  .showModalDialog(html,'Org Chart');
}

// Get Data from spreadsheet following the order：
// 0-Team	1-First Name 2-Last Name 3-Full Name 4-Position 5-Superior after 6-tooltips
function dataZoom() {
  var ss = SpreadsheetApp.getActive();
var  sheet = ss.getSheetByName('Maptek');
//  var sheet = ss.getActiveSheet();
  var blank = 0;
  // Only 26 column can be identifed
  var values = sheet.getRange("A:Z").getDisplayValues();
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

// Report Genetation
// Report global variable
  var team_admin = 0;
  var team_leader = 0;
  var team_tech = 0;
  var team_sales = 0;
  var team_finance = 0;
  var team_point = 0;
  var team_sentry = 0;
  var team_eureka = 0;
  var team_vulcan = 0;
  var team_blast = 0;
  var team_evo = 0;  
  var team_material = 0;
  var team_data = 0;
  var team_work = 0;
  var team_uix = 0;
  var team_hardware = 0;
  var team_integration = 0;
  var team_dev = 0;
  
  var team_admin_name = '';
  var team_leader_name = '';
  var team_tech_name = '';
  var team_sales_name = '';
  var team_finance_name = '';
  var team_point_name = '';
  var team_sentry_name = '';
  var team_eureka_name = '';
  var team_vulcan_name = '';
  var team_blast_name = '';
  var team_evo_name = '';  
  var team_material_name = '';
  var team_data_name = '';
  var team_work_name = '';
  var team_uix_name = '';
  var team_hardware_name = '';
  var team_integration_name = '';
  var team_dev_name = '';
  
  var people_number_A = 0;
  var people_number_B = 0;
  var people_number_C = 0;
  var people_number_D = 0;
  var people_number_E = 0;
  
  var people_name_A = '';
  var people_name_B = '';
  var people_name_C = '';
  var people_name_D = '';
  var people_name_E = '';
  
  var team_number_A = 0;
  var team_number_B = 0;
  var team_number_C = 0;
  var team_number_D = 0;
  var team_number_E = 0;
  
  var team_name_A = '';
  var team_name_B = '';
  var team_name_C = '';
  var team_name_D = '';
  var team_name_E = '';
  
  var male_number = 0;
  var female_number = 0;
  
  var name = '';

  var style1 = {};// style example 1
  style1[DocumentApp.Attribute.FONT_SIZE] = 10;
  style1[DocumentApp.Attribute.FONT_FAMILY] = DocumentApp.FontFamily.CALIBRI;
  style1[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000";
  var style2 = {};// style example 2
  style2[DocumentApp.Attribute.FONT_SIZE] = 16;
  style2[DocumentApp.Attribute.FONT_FAMILY] =DocumentApp.FontFamily.ARIAL_NARROW;
  style2[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000";
  var style3 = {};// style example 3
  style3[DocumentApp.Attribute.FONT_SIZE] = 12;
  style3[DocumentApp.Attribute.FONT_FAMILY] =DocumentApp.FontFamily.ARIAL;
  style3[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000";

function doReport(){
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
        if(allData[i][0] == 'Admin'){
          team_admin++;
        } 
        if(allData[i][0] == 'Leader'){
          team_leader++;
        }
        if(allData[i][0] == 'Technical services'){
          team_tech++;
        }
        if(allData[i][0] == 'Sales/marketing'){
          team_sales++;
        }
        if(allData[i][0] == 'Admin/finance'){
          team_finance++;
        }
        if(allData[i][0] == 'PointStudio'){
          team_point++;
        }
        if(allData[i][0] == 'Sentry'){
          team_sentry++;
        }
        if(allData[i][0] == 'Eureka'){
          team_eureka++;
        }
        if(allData[i][0] == 'Vulcan'){
          team_vulcan++;
        }
        if(allData[i][0] == 'BlastLogic'){
          team_blast++;
        }
        if(allData[i][0] == 'Evolution'){
          team_evo++;
        }
        if(allData[i][0] == 'MaterialMRT'){
          team_material++;
        }
        if(allData[i][0] == 'Data & SDK'){
          team_data++;
        }
        if(allData[i][0] == 'Workbench'){
          team_work++;
        }
        if(allData[i][0] == 'UIX'){
          team_uix++;
        }
        if(allData[i][0] == 'Hardware'){
          team_hardware++;
        }
        if(allData[i][0] == 'Hardware integration'){
          team_integration++;
        }
        if(allData[i][0] == 'Dev ops'){
          team_dev++;
        }
      }
      else if(values[0][j] == 'First Name'){
        allData[i][1] = values[i][j];
      }
      else if(values[0][j] == 'Last Name'){
        allData[i][2] = values[i][j];
      }
      else if(values[0][j] == 'Full Name'){
        allData[i][3] = values[i][j];
        if(allData[i][0] == 'Admin'){
          team_admin_name = allData[i][3]+' ,  '+team_admin_name;
        }
        if(allData[i][0] == 'Leader'){
          team_leader_name = allData[i][3]+' ,  '+team_leader_name;
        }
        if(allData[i][0] == 'Technical services'){
          team_tech_name = allData[i][3]+' ,  '+team_tech_name;
        }
        if(allData[i][0] == 'Sales/marketing'){
          team_sales_name = allData[i][3]+' ,  '+team_sales_name;
        }
        if(allData[i][0] == 'Admin/finance'){
          team_finance_name = allData[i][3]+' ,  '+team_finance_name;
        }
        if(allData[i][0] == 'PointStudio'){
          team_point_name = allData[i][3]+' ,  '+team_point_name;
        }
        if(allData[i][0] == 'Sentry'){
          team_sentry_name = allData[i][3]+' ,  '+team_sentry_name;
        }
        if(allData[i][0] == 'Eureka'){
          team_eureka_name = allData[i][3]+' ,  '+team_eureka_name;
        }
        if(allData[i][0] == 'Vulcan'){
          team_vulcan_name = allData[i][3]+' ,  '+team_vulcan_name;
        }
        if(allData[i][0] == 'BlastLogic'){
          team_blast_name = allData[i][3]+' ,  '+team_blast_name;
        }
        if(allData[i][0] == 'Evolution'){
          team_evo_name = allData[i][3]+' ,  '+team_evo_name;
        }
        if(allData[i][0] == 'MaterialMRT'){
          team_material_name = allData[i][3]+' ,  '+team_material_name;
        }
        if(allData[i][0] == 'Data & SDK'){
          team_data_name = allData[i][3]+' ,  '+team_data_name;
        }
        if(allData[i][0] == 'Workbench'){
          team_work_name = allData[i][3]+' ,  '+team_work_name;
        }
        if(allData[i][0] == 'UIX'){
          team_uix_name = allData[i][3]+' ,  '+team_uix_name;
        }
        if(allData[i][0] == 'Hardware'){
          team_hardware_name = allData[i][3]+' ,  '+team_hardware_name;
        }
        if(allData[i][0] == 'Hardware integration'){
          team_integration_name = allData[i][3]+' ,  '+team_integration_name;
        }
        if(allData[i][0] == 'Dev ops'){
          team_dev_name = allData[i][3]+' ,  '+team_dev_name;
        }
      }
      else if(values[0][j] == 'Position'){
        allData[i][4] = values[i][j];
      }
      else if(values[0][j] == 'Superior'){
        allData[i][5] = values[i][j];
      }
      else if(values[0][j] == 'Building'){
        allData[i][7] = values[i][j];
        
        if(allData[i][7] == 'Unit A'){
          people_number_A++;
          people_name_A = allData[i][3]+' ,  '+people_name_A;
          if(name != allData[i][0]){
            team_name_A = allData[i][0]+' ,  '+team_name_A;
            team_number_A ++;
          }
          name = allData[i][0];
        }
        if(allData[i][7] == 'Unit B'){
          people_number_B++;
          people_name_B = allData[i][3]+' ,  '+people_name_B;
          if(name != allData[i][0]){
            team_name_B = allData[i][0]+' ,  '+team_name_B;
            team_number_B ++;
          }
          name = allData[i][0];
        }
        if(allData[i][7] == 'Unit C'){
          people_number_C++;
          people_name_C = allData[i][3]+' ,  '+people_name_C;
          if(name != allData[i][0]){
            team_name_C = allData[i][0]+' ,  '+team_name_C;
            team_number_C ++;
          }
          name = allData[i][0];
        }
        if(allData[i][7] == 'Unit D'){
          people_number_D++;
          people_name_D = allData[i][3]+' ,  '+people_name_D;
          if(name != allData[i][0]){
            team_name_D = allData[i][0]+' ,  '+team_name_D;
            team_number_D ++;
          }
          name = allData[i][0];
        }
        if(allData[i][7] == 'Unit E'){
          people_number_E++;
          people_name_E = allData[i][3]+' ,  '+people_name_E;
          if(name != allData[i][0]){
            team_name_E = allData[i][0]+' ,  '+team_name_E;
            team_number_E ++;
          }
          name = allData[i][0];
        }
      }else if(values[0][j] == 'Sex'){
        allData[i][8] = values[i][j];
        if(allData[i][8] == 'Male'){
          male_number++;
        }
        if(allData[i][8] == 'Female'){
          female_number++;
        }
        
      }
      else{
        allData[i][num] = values[i][j];
        num ++;
      }
    }
  }

  // DELETE LAST COMMA
  people_name_A = people_name_A.substring(0,people_name_A.length-3);
  people_name_B = people_name_B.substring(0,people_name_B.length-3);
  people_name_C = people_name_C.substring(0,people_name_C.length-3);
  people_name_D = people_name_D.substring(0,people_name_D.length-3);
  people_name_E = people_name_E.substring(0,people_name_E.length-3);
  
  team_name_A = team_name_A.substring(0,team_name_A.length-3);
  team_name_B = team_name_B.substring(0,team_name_B.length-3);
  team_name_C = team_name_C.substring(0,team_name_C.length-3);
  team_name_D = team_name_D.substring(0,team_name_D.length-3);
  team_name_E = team_name_E.substring(0,team_name_E.length-3);
  
  team_admin_name = team_admin_name.substring(0,team_admin_name.length-3);
  team_leader_name = team_leader_name.substring(0,team_leader_name.length-3);
  team_tech_name = team_tech_name.substring(0,team_tech_name.length-3);
  team_sales_name = team_sales_name.substring(0,team_sales_name.length-3);
  team_finance_name = team_finance_name.substring(0,team_finance_name.length-3);
  team_point_name = team_point_name.substring(0,team_point_name.length-3);
  team_sentry_name = team_sentry_name.substring(0,team_sentry_name.length-3);
  team_eureka_name = team_eureka_name.substring(0,team_eureka_name.length-3);
  team_vulcan_name = team_vulcan_name.substring(0,team_vulcan_name.length-3);
  team_blast_name = team_blast_name.substring(0,team_blast_name.length-3);
  team_evo_name = team_evo_name.substring(0,team_evo_name.length-3);
  team_material_name = team_material_name.substring(0,team_material_name.length-3);
  team_data_name = team_data_name.substring(0,team_data_name.length-3);
  team_work_name = team_work_name.substring(0,team_work_name.length-3);
  team_uix_name = team_uix_name.substring(0,team_uix_name.length-3);
  team_hardware_name = team_hardware_name.substring(0,team_hardware_name.length-3);
  team_integration_name = team_integration_name.substring(0,team_integration_name.length-3);
  team_dev_name = team_dev_name.substring(0,team_dev_name.length-3);
  
//  Logger.log('The number of people in team Admin: ' + team_admin);  
//  Logger.log('The number of people in team Admin: ' + team_leader);
//  Logger.log('The number of people in team Admin: ' + team_tech);  
//  Logger.log('The number of people in team Admin: ' + team_sales);
//  Logger.log('The number of people in team Admin: ' + team_finance);  
//  Logger.log('The number of people in team Admin: ' + team_point);
//  Logger.log('The number of people in team Admin: ' + team_sentry);  
//  Logger.log('The number of people in team Admin: ' + team_eureka);
//  Logger.log('The number of people in team Admin: ' + team_vulcan);  
//  Logger.log('The number of people in team Admin: ' + team_blast);
//  Logger.log('The number of people in team Admin: ' + team_evo);  
//  Logger.log('The number of people in team Admin: ' + team_material);
//  Logger.log('The number of people in team Admin: ' + team_data);  
//  Logger.log('The number of people in team Admin: ' + team_work);
//  Logger.log('The number of people in team Admin: ' + team_uix);  
//  Logger.log('The number of people in team Admin: ' + team_hardware);
//  Logger.log('The number of people in team Admin: ' + team_integration);  
//  Logger.log('The number of people in team team_number_A: ' + team_number_A);
//  Logger.log('The number of people in team team_name_A: ' + team_name_A);  
//  
  

  
  
}


function printall(){
  doReport();
  var doc = DocumentApp.create('SMU Report');
  
  doc.getBody().appendParagraph('SMU Report').setAttributes(style2);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('The number of people in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: '+people_number_A).setAttributes(style3);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: '+people_number_B);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: '+people_number_C);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: '+people_number_D);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: '+people_number_E);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('List of names of staff in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_A).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_B).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_C).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_D).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_E).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('The number of people in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: '+team_number_A).setAttributes(style3);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: '+team_number_B);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: '+team_number_C);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: '+team_number_D);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: '+team_number_E);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('List of names of staff in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_A).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_B).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_C).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_D).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_E).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  
  
  doc.getBody().appendParagraph('The number of people in a team:').setAttributes(style2);
  doc.getBody().appendParagraph('Admin: '+team_admin).setAttributes(style3);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Leader: '+team_leader);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Technical services: '+team_tech);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sales/marketing: '+team_sales);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Admin/finance: '+team_finance);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('PointStudio: '+team_point);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sentry: '+team_sentry);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Eureka: '+team_eureka);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Vulcan: '+team_vulcan);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('BlastLogic: '+team_blast);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Evolution: '+team_evo);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('MaterialMRT: '+team_material);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Data & SDK: '+team_data);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Workbench: '+team_work);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('UIX: '+team_uix);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware: '+team_hardware);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware integration: '+team_integration);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Dev ops: '+team_dev);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('List of names of staff in a team:').setAttributes(style2);
  
  doc.getBody().appendParagraph('Admin: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_admin_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Leader: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_leader_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Technical services: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_tech_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sales/marketing: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_sales_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Admin/finance: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_finance_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('PointStudio: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_point_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sentry: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_sentry_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Eureka: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_eureka_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Vulcan: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_vulcan_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('BlastLogic').setAttributes(style3)
  doc.getBody().appendParagraph(team_blast_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Evolution: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_evo_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('MaterialMRT: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_material_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Data & SDK: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_data_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Workbench: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_work_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('UIX: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_uix_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_hardware_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware integration: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_integration_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Dev ops: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_dev_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  
  var sum =0 ;
  sum = male_number+female_number;
  doc.getBody().appendParagraph('Male and female staff ratio:').setAttributes(style2);
  doc.getBody().appendParagraph('Male Ratio: '+((male_number*100)/sum)+'%').setAttributes(style3);
  doc.getBody().appendParagraph('Female Ratio: '+((female_number*100)/sum)+'%');
  doc.getBody().appendParagraph(' ');
  
  Browser.msgBox('Successful','Please check your Google Drive.', Browser.Buttons.OK);
}

function print1(){ 
  doReport();
  var doc = DocumentApp.create('SMU People-Building Report');

  doc.getBody().appendParagraph('SMU People-Building Report').setAttributes(style2);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('The number of people in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: '+people_number_A).setAttributes(style3);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: '+people_number_B);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: '+people_number_C);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: '+people_number_D);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: '+people_number_E);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('List of names of staff in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_A).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_B).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_C).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_D).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: ').setAttributes(style3)
  doc.getBody().appendParagraph(people_name_E).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  
  Browser.msgBox('Successful','Please check your Google Drive.', Browser.Buttons.OK);
}

function print2(){
  doReport();
  var doc = DocumentApp.create('SMU Team-Building Report');
  
  doc.getBody().appendParagraph('SMU Team-Building Report').setAttributes(style2);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('The number of people in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: '+team_number_A).setAttributes(style3);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: '+team_number_B);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: '+team_number_C);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: '+team_number_D);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: '+team_number_E);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('List of names of staff in a building:').setAttributes(style2);
  doc.getBody().appendParagraph('Unit A: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_A).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit B: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_B).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit C: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_C).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit D: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_D).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Unit E: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_name_E).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  
  Browser.msgBox('Successful','Please check your Google Drive.', Browser.Buttons.OK);
}

function print3(){
  doReport();
  var doc = DocumentApp.create('SMU People-Team Report');

  doc.getBody().appendParagraph('SMU People-Team Report').setAttributes(style2);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('The number of people in a team:').setAttributes(style2);
  doc.getBody().appendParagraph('Admin: '+team_admin).setAttributes(style3);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Leader: '+team_leader);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Technical services: '+team_tech);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sales/marketing: '+team_sales);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Admin/finance: '+team_finance);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('PointStudio: '+team_point);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sentry: '+team_sentry);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Eureka: '+team_eureka);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Vulcan: '+team_vulcan);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('BlastLogic: '+team_blast);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Evolution: '+team_evo);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('MaterialMRT: '+team_material);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Data & SDK: '+team_data);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Workbench: '+team_work);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('UIX: '+team_uix);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware: '+team_hardware);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware integration: '+team_integration);
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Dev ops: '+team_dev);
  doc.getBody().appendParagraph(' ');
  
  doc.getBody().appendParagraph('List of names of staff in a team:').setAttributes(style2);
  
  doc.getBody().appendParagraph('Admin: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_admin_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Leader: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_leader_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Technical services: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_tech_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sales/marketing: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_sales_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Admin/finance: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_finance_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('PointStudio: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_point_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Sentry: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_sentry_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Eureka: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_eureka_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Vulcan: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_vulcan_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('BlastLogic').setAttributes(style3)
  doc.getBody().appendParagraph(team_blast_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Evolution: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_evo_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('MaterialMRT: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_material_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Data & SDK: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_data_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Workbench: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_work_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('UIX: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_uix_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_hardware_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Hardware integration: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_integration_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  doc.getBody().appendParagraph('Dev ops: ').setAttributes(style3)
  doc.getBody().appendParagraph(team_dev_name).setAttributes(style1);;
  doc.getBody().appendParagraph(' ');
  
  Browser.msgBox('Successful','Please check your Google Drive.', Browser.Buttons.OK);
}

function print4(){
  doReport();
  var doc = DocumentApp.create('SMU Ratio Report');
  
  doc.getBody().appendParagraph('SMU Ratio Report').setAttributes(style2);
  doc.getBody().appendParagraph(' ');

  var sum =0 ;
  sum = male_number+female_number;
  doc.getBody().appendParagraph('Male and female staff ratio:').setAttributes(style2);
  doc.getBody().appendParagraph('Male Ratio: '+((male_number*100)/sum)+'%'+'Female Ratio: '+((female_number*100)/sum)+'%').setAttributes(style3);
  doc.getBody().appendParagraph(' ');

  Browser.msgBox('Successful','Please check your Google Drive.', Browser.Buttons.OK);
}
