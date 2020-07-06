function onOpen(){
  myFunction();
}

function myFunction() {
    var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('讀取google calendar日誌小工具 »')
  .addItem('1.表列當月工作細項','main')
  .addSeparator().addToUi()
}
function calread(){
  //讀自己日曆
  var cal = CalendarApp.getDefaultCalendar();
  //讀現在開啟的資料表
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("查詢日期輸入");
  var output_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("output");
  //分類用
  var PT_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PT");
  var VA_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VA");
  var Others_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Others");
  var PT_row=1;
  var VA_row=1;
  var Others_row=1;
  //讀取資料表中的A1與B1
  var StartDate=sheet.getSheetValues(2,1,1,1);
  var EndDate=sheet.getSheetValues(2,2,1,1);
  var events = cal.getEvents(new Date(StartDate),new Date(EndDate));
  var timezone= cal.getTimeZone();
  
  //清空資料表
  resetsheet("output")
  resetsheet("PT")
  resetsheet("VA")
  resetsheet("Others")
  
  for (var i=0;i<events.length;i++) {
    //http://www.google.com/google-d-s/scripts/class_calendarevent.html
    //因API取出endtime會多算一天，到隔天00:00，所以故意GMT-1，使結束時間正確
    var details=[[i+1,events[i].getTitle(), events[i].getDescription(), events[i].getStartTime(), Utilities.formatDate(new Date(events[i].getEndTime()), "GMT-1", "YYYY/MM/dd")]];
    var row=i+1;
    var range=output_sheet.getRange(row+2,1,details.length,details[0].length);
    //總表輸出於output sheet
    range.setValues(details);
    //分類用
    var PT_range=PT_sheet.getRange(PT_row+2,1,details.length,details[0].length);
    var VA_range=VA_sheet.getRange(VA_row+2,1,details.length,details[0].length);
    var Others_range=Others_sheet.getRange(Others_row+2,1,details.length,details[0].length);
    var company= details[0][1].split('_');
    //Logger.log('company %s' ,company[0]);
    if (details[0][1].match('PT')){
      Logger.log('PT %s PT_row= %s',details[0][1],PT_row);
      PT_range.setValues(details);
      PT_row=PT_row+1;
      Logger.log('Time zone is  %s',timezone);
    }else if (details[0][1].match('VA')){
      Logger.log('VA %s VA_row= %s',details[0][1],VA_row);
      VA_range.setValues(details);
      VA_row=VA_row+1;
    }else{
      Logger.log('Others %s Others_row= %s',details[0][1],Others_row);
      Others_range.setValues(details);
      Others_row=Others_row+1;
    }
  }
  return PT_row,VA_row,Others_row;
}
/**
function classification(string){
  var details_string=string;
  if (details_string.match('PT')){
    //Logger.log('PT %s PT_row= %s',details_string,PT_row);
    PT_range.setValues(details);
    PT_row=PT_row+1;
  }else if (details_string.match('VA')){
    //Logger.log('VA %s VA_row= %s',details_string,VA_row);
    VA_range.setValues(details);
    VA_row=VA_row+1;
  }else{
    //Logger.log('Others %s Others_row= %s',details_string,Others_row);
    Others_range.setValues(details);
    Others_row=Others_row+1;
  }
}
**/

function resetsheet(sheetname){  //清空資料表
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var range=sheet.getRange(1,1,2,2);
  var defaultValues =[['',''],['項次','工作']];
  
  //不同sheet給不同初始值
  switch(sheetname){
    case "PT":
      defaultValues =[['',''],['','滲透測試(PT)']];
      break;
    case "VA":
      defaultValues =[['',''],['','主機弱掃(VA)']];
      break;
    case "Others":
      defaultValues =[['',''],['','其他']];
      break;
    default:
      //do not thing
  }
  sheet.clearContents();
  range.setValues(defaultValues);
}

function sort(sheetname){
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var range=sheet.getRange("B3:E");
  range.sort([{column: 2, ascending: false},{column: 4, ascending: false}]);
  
}
function classification(sheetname){
  var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var counter=1;
  if(lastColumn==2) return 0;//若其中一個項目(PT/VA/Other)為空，則直接離開classification
  var Joblist=sheet.getSheetValues(3, 2,lastRow-3+1, lastColumn-2+1);
  Logger.log('lastColumn %s',lastColumn);
  Logger.log('lastRow %s',lastRow);
  Logger.log('Joblist %s',Joblist[0][0]);
  Logger.log('Joblist.legth %s %s',Joblist.length,Joblist[0].length);
  var JobtoOneblock="";
  var joblist_ex="";
  var joblist_type="";
  //將制式輸出改為可閱讀模式 ex:
  for(var i=0;i<Joblist.length;i++){
    //如果是pt or va類型
    if(Joblist[i][0].match('PT')||Joblist[i][0].match('VA')){
      //分出主協辦類型
      if(Joblist[i][0].match('協辦')){
        if(Joblist[i][0].split('_')[0] == joblist_ex ){
           var Jobdetail='\t第'+Joblist[i][0].split('_')[3]+'次'+Joblist[i][0].split('_')[4]+'\('+dateformat(Joblist[i][2],Joblist[i][3])+'\)\n';
        }else {           
           var Jobdetail='\('+counter+'\)'+Joblist[i][0].split('_')[0]+'\(協辦\)\n'+'\t第'+Joblist[i][0].split('_')[3]+'次'+Joblist[i][0].split('_')[4]+'\('+dateformat(Joblist[i][2],Joblist[i][3])+'\)\n';
           counter+=1;
        }
      }else {
        if(Joblist[i][0].split('_')[0] ==joblist_ex ){
          var Jobdetail='\t第'+Joblist[i][0].split('_')[3]+'次'+Joblist[i][0].split('_')[4]+'\('+dateformat(Joblist[i][2],Joblist[i][3])+'\)\n';
        }else {
          var Jobdetail='\('+counter+'\)'+Joblist[i][0].split('_')[0]+'\(主辦\)\n'+'\t第'+Joblist[i][0].split('_')[3]+'次'+Joblist[i][0].split('_')[4]+'\('+dateformat(Joblist[i][2],Joblist[i][3])+'\)\n';
          counter+=1;
        }
      }
      joblist_ex=Joblist[i][0].split('_')[0];
      
    }else /*if(Joblist[i][0].match('其它'))*/{
      //其它，若只有一個項次可分類，就只秀出其中一項加上分類
      if (Joblist[i][0].split('_')[2]){
        
        if(Joblist[i][0].split('_')[1] ==joblist_ex){
           counter+=1;
           var Jobdetail='\t\('+counter+'\)'+Joblist[i][0].split('_')[2]+'\('+dateformat(Joblist[i][2],Joblist[i][3])+'\)\n';
        }else {
           counter=1
           var Jobdetail='\b'+Joblist[i][0].split('_')[1]+'\n\t'+'\('+counter+'\)'+Joblist[i][0].split('_')[2]+'\('+dateformat(Joblist[i][2],Joblist[i][3])+'\)\n';
        }
      }else{

        var Jobdetail='\b'+Joblist[i][0].split('_')[1]+'\n\t\('+dateformat(Joblist[i][2],Joblist[i][3])+'\)\n';
      }
      joblist_ex=Joblist[i][0].split('_')[1];
    }
    
    //var Job_range = sheet.getRange(20+i,2,Jobdetail.length,Jobdetail[0].length);
    //var Job_range = sheet.getRange(20,2,1,1);
    JobtoOneblock=JobtoOneblock+Jobdetail;
    //Job_range.setValues(JobtoOneblock);
  }
  Logger.log(JobtoOneblock);
  var Job_range = sheet.getRange("B1");
  Job_range.setValue(JobtoOneblock);

}
//轉換標準時間格式為 ex: 11/12~11/20
function dateformat(date1,date2){
  Logger.log('startdate =%s',date1);
  Logger.log('enddate =%s',date2);
  date1=Utilities.formatDate(new Date(date1), "GMT+8", "MM/dd");
  date2=Utilities.formatDate(new Date(date2), "GMT+8", "MM/dd");
  //var date1 =Utilities.formatDate(new Date(date1), "GMT+8", "yyyy/MM/dd");
  //如果日期一樣就只輸出單日 ex:11/12
  if(date2 != date1){
    var formatvalue =date1+'~'+date2;
  }else{
    //var formatvalue =Utilities.formatDate(new Date(date1), "GMT+8", "yyyy/MM/dd");
    var formatvalue =date1;
  }
  return formatvalue;
  
}

function main(){
  var PT_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PT");
  var VA_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VA");
  var Others_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Others");
  var main_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("查詢日期輸入");
  calread();
  sort("PT");
  classification("PT");
  sort("VA");
  classification("VA");
  sort("Others");
  classification("Others");
  
  var list_result='1.滲透測試\(PT\)\n'+PT_sheet.getRange("B1").getValue()+'\n2.主機弱掃\(VA\)\n'+VA_sheet.getRange("B1").getValue()+'\n3.其它\n'+Others_sheet.getRange("B1").getValue();
  main_sheet.getRange("B5").setValue(list_result)
}