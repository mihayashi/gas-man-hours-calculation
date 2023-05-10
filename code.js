
/**
 * 前月のteamspirit工数実績を集計し、スプレッドシートに結果を出力する
 * @returns {any}
 */
function main(){
  //処理開始時刻
  const s = new Date();

  // teamspiritトークン取得 
  let resData = getToken();

  // const startDate = "2021-11-01";
  // 処理開始日・終了日を取得
  const startEndDate = getstartEndDate();
  const startDate = startEndDate.startDate;
  const endDate = startEndDate.endDate;
  const yyyymm = startDate.slice(0,7);
  const targetMonth = yyyymm.replace("-","/");

  //対象empIDリスト取得
  const targets = getTargets(resData,targetMonth);
 
  
  // 対象者ごとに対象タスクリスト取得
  targets.forEach(arr=>{
    getData(arr[0],arr[1],startDate,endDate,resData);
  });

  // ジョブ集計
  let jobDataSums = aggregatingJobs(jobData);
  // Logger.log(jobDataSums);

  // シート作成
  let newSheet = "";
  try{
    newSheet = createNewSheet(startDate);
  }catch(error){
    Logger.log("error "+error.name + ": "+ error.message);
    let e = new Date();
    let pastTime = Math.round(((e - s)/1000)*10)/10;  //小数点2位以下四捨五入
    Logger.log("処理時間(s):"+ pastTime);

    //send the result to google Chat
    const errorText = `teamspirit_kosuShukei:\nhttps://docs.google.com/spreadsheets/d/xxxxx \n工数集計中止: 対象年月: ${targetMonth}, ${error.message}`;
    sendChat(errorText);
    return;
  }
  Logger.log("シート名: "+ newSheet.getSheetName());

  //詳細データ出力
  console.log(data.length +'件');
  newSheet.getRange(newSheet.getLastRow()+1,1,data.length, data[0].length).setValues(data);


  // スプレッドシートにジョブ集計結果を書き込み
  newSheet.getRange(2,10,jobDataSums.length, jobDataSums[0].length).setValues(jobDataSums);

  // 集計エリアをジョブ名でソート
  let jobSumRange = newSheet.getRange("J2:L");
  jobSumRange.sort(10);


  // 処理結果出力
  let e = new Date();
  let pastTime = Math.round(((e - s)/1000)*10)/10;  //小数点2位以下四捨五入
  Logger.log("対象者: "+targets.length + "件, 対象タスク数: "+ data.length + "件, 対象ジョブ数: "+ jobDataSums.length + "件, 処理時間(s): "+ pastTime);

  //send the result to google Chat
  const resultText = `teamspirit_kosuShukei:\nhttps://docs.google.com/spreadsheets/d/xxxxx \n工数集計完了: 対象年月: ${targetMonth}, 対象者: ${targets.length}件, 対象タスク数: ${data.length}件, 対象ジョブ数: ${jobDataSums.length}件, 処理時間(s): ${pastTime}`;
  sendChat(resultText);

}

/**
 * ジョブ集計
 * @param {any} jobData
 * @returns {any}
 */
function aggregatingJobs(jobData){

  let group = jobData.reduce(function (result, current) {
  let element = result.find(function (p) {
    return p[0] === current[0]
  });
  if (element) {
    element.count ++; // count
    element[1] += current[1]; // sum
    element[2] += current[2]; // sum
  } else {
    result.push([current[0],current[1],current[2]]);
  }
  return result;
}, []);

let newGroup = group.map(value =>{
  let minutes = value[2];
  let hour = Math.floor(minutes / 60);
  let min = minutes % 60;
  let time = hour + ":" + ("00"+min).slice(-2);

  return [value[0],value[1],time];
});
Logger.log(newGroup);
return newGroup;

}


/**
 * teamspiritクエリー実行
 * @param {any} empID
 * @param {any} empName
 * @param {any} startDate
 * @param {any} endDate
 * @param {any} resData
 * @returns {any}
 */
function getData(empID,empName, startDate,endDate,resData){
  let records = query(resData,"SELECT Name, teamspirit__JobId__c, teamspirit__TaskNote__c, MajorClassification__c, MiddleClass__c, teamspirit__Process__c, teamspirit__WorkDate__c, teamspirit__TimeH__c, teamspirit__Time__c, teamspirit__CalcTime__c, teamspirit__JobApplyId__c, teamspirit__EmpId__c FROM teamspirit__AtkEmpWork__c WHERE teamspirit__EmpId__c = '"+empID+"' AND teamspirit__WorkDate__c >= "+startDate +" AND teamspirit__WorkDate__c < "+endDate);
  // let data = [];
  // Logger.log(records);
  // Logger.log(empName);
  for(let i = 0; i < records.length; i++){

    let date = records[i].teamspirit__WorkDate__c;
    let daiBunrui = records[i].MajorClassification__c;
    let chuBunrui = records[i].MiddleClass__c;
    let kotei = records[i].teamspirit__Process__c;
    let time = records[i].teamspirit__TimeH__c;
    let minutes = records[i].teamspirit__Time__c;

    let jobName = "";
    //大分類1文字目がアルファベットか
    let isAlphabet = alphaBetsPattern.test(daiBunrui);

    if(isAlphabet){

    // ジョブ出力用配列
    data.push([empID,empName, date, daiBunrui,chuBunrui, kotei,time,minutes]);

    // ジョブ集計用配列
    let chu = "";
    if(chuBunrui){
      chu = "_"+chuBunrui;
    }else{
      chu = "";
    }
    // jobName = daiBunrui + "_"+chuBunrui;
    jobName = daiBunrui + chu;
    jobData.push([jobName,time,minutes]);
    }

               }
  // console.log(data.length +'件');
  // newSheet.getRange(newSheet.getLastRow()+1,1,data.length, data[0].length).setValues(data);

}

/**
 * クエリー処理
 * @param {any} resData
 * @param {any} q
 * @returns {any}
 */
function query(resData,q){
  
  q = encodeURIComponent(q);
  let sessionInfo = JSON.parse(resData);
  let res = UrlFetchApp.fetch(
    sessionInfo.instance_url + "/services/data/v45.0/query/?q=" + q, {
    "method":"GET",
    "headers":{
    "Authorization":"Bearer " + sessionInfo.access_token
    }
    });
  
  let queryResult = JSON.parse(res.getContentText());
  return queryResult.records;
  
}

/**
 * teamspirit トークン取得
 * @returns {any}
 */
function getToken() {
  let res = UrlFetchApp.fetch(
    "https://login.salesforce.com/services/oauth2/token",
    {
      "method":"POST",
      "payload" : {
        "grant_type": "password",
        "client_id": "xxxxx",
        "client_secret":"xxxxx",
        "username":"メールアドレス",
        "password":"xxxxx"
      },
      "muteHttpExceptions":true
    });
  if(res.getResponseCode() == 200){
//    let prop = PropertiesService.getScriptProperties();
//    prop.setProperty("session_info", res.getContentText());
    resData = res.getContentText();
    // Logger.log(resData);
    return resData; 
  }
}

/**
 * 工数集計対象者リスト取得
 * @returns {any}
 */
function getTargets(resData,targetMonth){
  let records = query(resData,"SELECT teamspirit__EmpId__c, teamspirit__EmpName__c FROM teamspirit__AtkJobApply__c WHERE teamspirit__YearMonthS__c = '"+ targetMonth+"'");

  for(let i = 0; i < records.length; i++){
    let empID = records[i].teamspirit__EmpId__c;
    let empName = records[i].teamspirit__EmpName__c;

    empData.push([empID, empName]);
  }
  Logger.log(empData);
  return empData;
}
// function getTargets(){
//   // const ss = SpreadsheetApp.openById(ssid).getSheetByName("list");
//   // const ssRows = ss.getRange(1,2,ss.getLastRow(), ss.getLastColumn()-1).getDisplayValues();
//   //従業員リスト取得
//   let empSheet = SpreadsheetApp.openById(ssid).getSheetByName('list');
//   let empArr = empSheet.getDataRange().getValues();

//   //「対象：y」のメンバーを抽出
//   const sheetTargetsArr = empArr.filter(emp => emp[2] ==="y");
//   Logger.log(sheetTargetsArr.length + "件");
  
//   //「対象：y」を除いた配列を戻り値とする
//   const targetsArr = sheetTargetsArr.map(arr => arr.splice(0,2));
//   Logger.log(targetsArr);
//   return targetsArr;

// }

/**
 * 前月1日 + 1か月後の1日を返す
 * @returns {any}
 */
function getstartEndDate(){
  //todayを任意の日付とすることも可。そこから1か月が指定期間となる
  const today = dayjs.dayjs();
  const lastMonthFirstDay = today.add(-1, 'month').date(1);
  const lastMonthFirstDayString = lastMonthFirstDay.format('YYYY-MM-DD');
  const nextMonthFirstDay = lastMonthFirstDay.add(1, 'month');
  const nextMonthFirstDayString = nextMonthFirstDay.format('YYYY-MM-DD');

  const startEndDate = {

    startDate: lastMonthFirstDayString,
    endDate: nextMonthFirstDayString

  }
  return startEndDate;
}

/**
 * 当月シート作成
 * @param {any} startDate
 * @returns {any}
 */
function createNewSheet(startDate){
  const newSheetName = startDate.slice(0,4) + startDate.slice(5,7) ;
  const newSheet = templateSheet.copyTo(spreadSheet);
  newSheet.setName(newSheetName);
  spreadSheet.setActiveSheet(newSheet);
  spreadSheet.moveActiveSheet(1);

  return newSheet;

}


/**
 * 処理結果をgoogle Chatスペース「kosuShukei」に送信
 *
 * @param {*} text
 */
function sendChat(text){
  const body = {'text': text};
  const params = {
    'method': 'post',
    contentType: 'application/json',
    'payload': JSON.stringify(body)
  };
  UrlFetchApp.fetch(url, params);
}
