let empData = [];
let resData = "";
let data = [];
let jobData = [];

//頭1文字アルファベット
const alphaBetsPattern = /^[a-zA-Z]/;

const ssid = "xxxxxxxxxxxx";
const spreadSheet = SpreadsheetApp.openById(ssid);
const templateSheet = spreadSheet.getSheetByName("template");

// google chat通知用webhook
const url = "xxxxxxxxxxxxx";