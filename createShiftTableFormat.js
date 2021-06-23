const START_PASS = "38"; // start row number from pass prediciton sheet //have to change each time
const END_PASS = "88"; // End row number from pass prediciton sheet //have to change each time

const SHIFT_DAYS = 7; //how many days does this shift table show
const CUT_OFF_MEL = 0;//[deg]
//const CUT_OFF_TIME = new Date().setTime(8*60*1000); //[ms]
const CUT_OFF_TIME = 0;// [min]
const CUT_OFF_HOUR_START = 0; // o'clock
const CUT_OFF_HOUR_END = 6// o'clock
//change url every time 

const URL = "";//debug
//const URL = "";//actual operation


function fetchPassInfo() {
  var ss = SpreadsheetApp.openByUrl("");
  var sheet = ss.getSheetByName("参照");
  var passInfoRange = sheet.getRange(START_PASS, 1, END_PASS - START_PASS + 1, 8)  //該当のパス群の行を指定する必要あり 
  var passInfo = passInfoRange.getValues();
  var passInfoByDay = []
  var tmp = []
  for(var i=0; i<passInfo.length; i++){
    tmp.push(passInfo[i]);
    if(i==passInfo.length-1){
      passInfoByDay.push(tmp);
      continue;
    }
    else if((passInfo[i+1][0].getTime()-passInfo[i][0].getTime())>36000000){
      passInfoByDay.push(tmp);
      tmp = [];
    }
  }   

  return passInfoByDay;
}



function fetchShiftsheet(sheet_name){
  
  var ss = SpreadsheetApp.openByUrl(URL);
  var sheet = ss.getSheetByName(sheet_name);
  Logger.log(sheet);
  if(sheet){
    return sheet;
  }
  else{
    sheet=ss.insertSheet();
    sheet.setName(sheet_name);
    return sheet;
  }

}

function createShiftTable(sheet,numMaxMember,color){
  var targetSheet = fetchShiftsheet(sheet);
  var passInfoByDay = fetchPassInfo();
  var passNumList = ['①', '②', '③', '④', '➄', '⑥', '⑦', '⑧']; //そのパス群で何番目のパスかを示す数字
  var dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'];
  
  var numColumn = 0;



  targetSheet.setColumnWidths(1, SHIFT_DAYS*2, 45);
  targetSheet.getRange(1,1,1+numMaxMember*8,1+SHIFT_DAYS*2).setHorizontalAlignment('center');
  targetSheet.setTabColor(color);

  for(var numColumn=0; numColumn<SHIFT_DAYS; numColumn++){

    day = (passInfoByDay[numColumn][0][0].getMonth()+1) + "/" + passInfoByDay[numColumn][0][0].getDate() + "(" + dayOfWeek[passInfoByDay[numColumn][0][0].getDay()] + ")";
    targetSheet.getRange(1,1+numColumn*2).setValue(day);
    targetSheet.getRange(1,1+numColumn*2).setBackground(color);
    targetSheet.getRange(1,1+numColumn*2,1+numMaxMember*8,2).setBorder(true, true, true, true, true, true,'black',SpreadsheetApp.BorderStyle.SOLID);
    targetSheet.getRange(1,1+numColumn*2,1+numMaxMember*8,2).setBorder(true, true, true, true, null, null,'black',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    //targetSheet.getRange(1,1+numMaxMember*8).setBorder(true, true, true, true, true, true,SpreadsheetApp.BorderStyle.SOLID);
    
    targetSheet.getRange(1,1+numColumn*2,1,2).merge();

    for(var eachPass = 0; eachPass < passInfoByDay[numColumn].length; eachPass++ ){
      targetSheet.getRange(2+eachPass*numMaxMember,1+numColumn*2).setValue(passNumList[eachPass]);
      targetSheet.getRange(3+eachPass*numMaxMember,1+numColumn*2).setValue(passInfoByDay[numColumn][eachPass][0]);
      //targetSheet.getRange(3+eachPass*numMaxMember,1+numColumn*2).setValue(Utilities.formatDate(passInfoByDay[numColumn][eachPass][0], "JST", "HH:mm"));
      targetSheet.getRange(3+eachPass*numMaxMember,1+numColumn*2).setNumberFormat("HH:mm");

      //Logger.log(passInfoByDay[numColumn][eachPass][2].getMinutes());//.getMinutes()
      //Logger.log(CUT_OFF_TIME)
      //(passInfoByDay[numColumn][eachPass][3] < CUT_OFF_MEL)
      //(passInfoByDay[numColumn][eachPass][2].getMinutes() < CUT_OFF_TIME)
      //
      if(CUT_OFF_HOUR_START < passInfoByDay[numColumn][eachPass][0].getHours() &&passInfoByDay[numColumn][eachPass][0].getHours() < CUT_OFF_HOUR_END  ){
        
        targetSheet.getRange(2+eachPass*numMaxMember,2+numColumn*2,2,1).setBackground("#999999");
      }
            
    }

  } 

}

function main(){
  sheet = fetchShiftsheet("従事者");
  sheet.clear();
  sheet = fetchShiftsheet("責任者");
  sheet.clear();  

  numMaxMember = 3 
  createShiftTable("従事者",numMaxMember,"#cfe2f3");

  numMaxMember = 3 
  createShiftTable("責任者",numMaxMember,"#fce5cd");
  Logger.log("execution");
}
