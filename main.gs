// テストの際はgetFormats関数とgetLatestKeysSheet関数の2つのIDを書き換えて行う

// 現在時刻の取得
//  00:00 - 06:00(直前) -> midnight
//  06:00 - 12:00(直前) -> morning
//  12:00 - 18:00(直前) -> afternoon
//  18:00 - 24:00(直前) -> night
const getTimeframe = () => {
  const hours = new Date().getHours();
  var res = "default";
  if(hours < 6){
    res = "midnight";
  }else if(hours < 12){
    res = "morning";
  }else if(hours < 18){
    res= "aftenoon";
  }else{
    res = "night";
  }
  console.log(res);
  return res;
}

// スクレイピングするサイトのフォーマット情報を取得
const getFormats = () => {
  const folderId = 'JSONフォルダを入れたフォルダのID'; // 処理するフォルダのIDを指定(本番用)
  // const folderId = 'JSONフォルダを入れたフォルダのID'; // テスト用のフォルダID
  const folder = DriveApp.getFolderById(folderId); // フォルダを取得
  const files = folder.getFiles(); // フォルダ内のファイルを取得
  
  var res=[];
  const timeframe = getTimeframe();

  while (files.hasNext()) {
    var file = files.next(); // 次のファイルを取得
    var json = JSON.parse(file.getBlob().getDataAsString()); // ファイルをjsonとして読み込み
    if("Timeframe" in json){  // フォーマットjsonに"Timeframe"の項目がある場合は現在のtimeframeと一致するかをチェック
      if(json.Timeframe == timeframe){
        res.push(json);  //Timeframeが一致する場合jsonを配列に追加
      }
    }else{  // フォーマットjsonに"Timeframe"の項目がない場合は現在のtimeframeが"night"の場合に実行
        if(timeframe == "night"){
        res.push(json);  //Timeframeが一致する場合jsonを配列に追加
      }
    }
  }
  return res;
}

// スクレイピングデータを格納するスプレッドシートを取得
const getSheet = (sheetId) => {
  const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  return sheet;
}

// スクレイピングデータを格納するスプレッドシートのkeyカラム情報を保存したスプレッドシートを取得
// DBのkey: auto_incrementの代用
const getLatestKeysSheet = () =>{
  const id = 'スプシID';  // 本番用
  // const id = 'スプシID'; // テスト用
  const sheet = SpreadsheetApp.openById(id).getSheets()[0];
  return sheet;
}

// keyカラム情報を保存したスプレッドシートから該当作品の直近のキー番号のセルを取得
const getLatestKeysRange = (latestKeysSheet, pageName) => {
  const lastRow = latestKeysSheet.getLastRow();
  for(let i = 1; i <= lastRow; i++){
    const value = latestKeysSheet.getRange(i,1).getValues();
    if(value == pageName){
      return latestKeysSheet.getRange(i,2);
    }
  }
  return latestKeysSheet.getRange(1,10);  // error
}

// CloudFunctions上でスクレイピングを行う
const cloudFunctionScraping = (page,token) => {
  const url = "cloudfunctionsの実行URL";
  const html = UrlFetchApp.fetch(url, {
    headers: {
      "Authorization": token,
      "Content-Type": "application/json",
      "method": ["GET","POST"],
      "muteHttpExceptions" : true
    },
    "payload" : JSON.stringify(page),
  })
  const res = JSON.parse(html.getContentText());
  return res;
}

// JSON->二次元配列
const jsonToTable = (jsonData) => {
  // データが存在しない場合は空の配列を返す
  if (!jsonData || !jsonData.length) {
    return [];
  }
  
  // データを格納する二次元配列を作成する
  const tableData = [];
  
  // 各行のデータを配列に変換して追加する
  for (let i = 0; i < jsonData.length; i++) {
    const rowData = [];
    for (const key in jsonData[i]) {
      rowData.push(jsonData[i][key]);
    }
    tableData.push(rowData);
  }
  
  return tableData;
}

// シートの追加 + 列名入力
const addSheet = (spreadSheet,lastSheet) =>{
  const newSheet = spreadSheet.insertSheet();
  const columnNames=lastSheet.getRange(1, 1, 1, lastSheet.getLastColumn()).getValues();
  newSheet.getRange(1, 1, 1, lastSheet.getLastColumn()).setValues(columnNames);
  return;
}

// 各フォーマットごとにスクレイピングを行いスプレッドシートに格納する関数
const functionForRun = (page,token) =>{
  var res=[]
  res.push("Scraping Starts in : "+page.PageName);
  try{
    // スプシ（ID -> ファイル -> 一番右のシート -> 最終行）を取得
    const sheetId = page.SheetId;
    const spreadSheet = SpreadsheetApp.openById(sheetId);
    const numSheets = spreadSheet.getNumSheets();
    const sheet = spreadSheet.getSheets()[numSheets - 1];
    let lastRow = sheet.getLastRow();

    if(lastRow==0){lastRow++;}  // 列名記入忘れ対策
    res.push("Get Sheet in : "+page.PageName);


    // Cloud Functions上でスクレイピングの実行
    data = cloudFunctionScraping(page,token);
    // console.log(data);

    // データ挿入の下準備 JSON->配列
    const dataArray = jsonToTable(data);
    const numRows = dataArray.length;
    const numCols = dataArray[0].length;

    // データを挿入
    sheet.getRange(lastRow + 1, 2, numRows, numCols).setValues(dataArray);

    // 各データの1列目に一意のキー番号を入れる
    const latestKeysSheet = getLatestKeysSheet();
    const latestKeysRange = getLatestKeysRange(latestKeysSheet, page.PageName);
    const latestKeyNum = parseInt(latestKeysRange.getValues());
    const columnNumbers = [];
    for (let i = 1; i <= numRows; i++) {
      columnNumbers.push([latestKeyNum + i]);
    }
    sheet.getRange(lastRow + 1, 1, numRows, 1).setValues(columnNumbers);
    
    // 最新キー番号を更新
    const newLatestKeyNum = latestKeyNum + numRows;
    latestKeysRange.setValues([[ newLatestKeyNum ]]);

    // 行数が1,000,000を超えてたら新しいシートを追加
    if(sheet.getLastRow()>1000000){
      addSheet(spreadSheet,sheet);
    }
    res.push("Scraping Completed in : " + page.PageName);
    return(res);
  }catch(e){
    res.push("An error occurred in " + page.PageName);
    res.push(e.message);
    throw(res);
  }
}


// メイン関数
const main = () => {
  // format_jsonフォルダからフォーマット(=page)の配列(=pages)を取得
  let pages;
  try{
    pages = getFormats();
  }catch(e){
    console.error(e);
  }

  // 並列処理したい関数を配列workersArrayに格納
  // GCPへのアクセストークンはRunAll内部で生成せずに引数として渡さないと動かないため注意
  const token = "Bearer "+ ScriptApp.getIdentityToken();
  var workersArray=[]
  for(const page of pages){
    workersArray.push({
      functionName: "functionForRun",
      arguments: [page,token],
    });
  }
  var results = RunAll.Do(workersArray);
  var errors = [];
  results.forEach((res)=>{
    Logger.log(res);
    if("error" in JSON.parse(res.getContentText())){
      errors.push(JSON.parse(res.getContentText()).error);
    }
  });
  if(errors.length !== 0){
    Logger.log(errors);
    const errorMessage = "Apps Scriptでエラーが発生しました。確認してください。エラー数: " + errors.length;
    throw(errorMessage);
  }
  return;
}
