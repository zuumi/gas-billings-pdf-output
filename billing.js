function billing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const billList = ss.getSheetByName("リスト");
  lastRow = billList.getLastRow();
  var listRange = billList.getDataRange().getValues();//スプレッドシートのデータを二次元配列として取得
  var list = [];

  //会社を取得
  for (var i=1; i<listRange.length; i++){ 
    list.push(listRange[i][4]);//配列torihikiにmyRange[i][4]を追加
  }
  // 会社の重複を削除
  var listCompany = list.filter(function(value, i, self){ 
    return self.indexOf(value) === i;
  });

  for (i=0;i<listCompany.length;i++) {
    addSheet(listCompany[i],listRange);
  }
}

function addSheet(companyName,listRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const format = ss.getSheetByName("請求書フォーマット");

  //同じ名前のシートがなければ作成
  var sheet = SpreadsheetApp.getActive().getSheetByName(companyName)
  if(sheet) {
    ss.deleteSheet(sheet);
  }

  const newSheet = format.copyTo(ss);
  newSheet.setName(companyName);
  newSheet.getRange(6,6).setValue(companyName);
  getlist = [];

  for (i=0;i<listRange.length;i++) {
    if (listRange[i][4] === companyName) {
      getlist.push(listRange[i]);
    }
  }
  for (i=0;i<getlist.length;i++) {
    newSheet.getRange(14 + i, 2).setValue(getlist[i][0]);
    newSheet.getRange(14 + i, 2).setBorder(true, true , true, true, false, false);
    newSheet.getRange(14 + i, 3).setValue(getlist[i][1]);
    newSheet.getRange(14 + i, 3).setBorder(true, true , true, true, false, false);
    newSheet.getRange(14 + i, 6).setValue(getlist[i][2]);
    newSheet.getRange(14 + i, 6).setBorder(true, true , true, true, false, false);
    newSheet.getRange(14 + i, 7).setValue(getlist[i][3]);
    newSheet.getRange(14 + i, 7).setBorder(true, true , true, true, false, false);
  }
  sheetlastRow = newSheet.getLastRow();

  const folderId = exportSheetToPDF(companyName,companyName,sheetlastRow);
  const cSheet = ss.getSheetByName('会社情報');
  var lastRow = cSheet.getLastRow();
  for (i=2;i<=lastRow;i++) {
    let companyNameTest = cSheet.getRange(i,1).getValue();
    if (companyName === companyNameTest){
      var address = cSheet.getRange(i,3).getValue();
    }
  }
  sendEmail(folderId,companyName,address);
}

function exportSheetToPDF(sheetName,pdfName,sheetlastRow){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folderURL = ss.getSheetByName('リスト').getRange(2,8).getValue();//出力先のフォルダーのURLを入力
  const rss = ss.getSheetByName(sheetName);
  const sheetId = rss.getSheetId();

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';
  const lefthead = 'A1';
  const rightfoot = 'H' + sheetlastRow;

  const options = 'exportFormat=pdf&format=pdf&gid=' + sheetId  //PDFにするシートの「シートID」
  + '&portrait=true'  //true(縦) or false(横)
  + '&size=A4&fitw=true'//true(幅を用紙に合わせる) or false(原寸大)
  + '&gridlines=false' //グリッドラインの表示有無
  + '&range='+ lefthead + '%3A' + rightfoot;

  const requestUrl = url + options;
  const token = ScriptApp.getOAuthToken();
  
  const params = {
    'headers' : {'Authorization':'Bearer ' + token},
    'muteHttpExceptions' : true
  };
  
  const response = UrlFetchApp.fetch(requestUrl, params);
  
  //Blobオブジェクトを作成
  const blob = response.getBlob().setName(pdfName + '.pdf'); //PDFファイル名を設定
  
  //指定のフォルダにPDFファイルを作成
  const folderId = folderURL.slice(39);
  const folder = DriveApp.getFolderById(folderId);
  folder.createFile(blob);
  return folderId;
}

function sendEmail(folderId,fileName,address) {
  //メールの件名を記述する
  let mailTitle = "添付テストメール";
  //メール本文を記述する
  let mailText = "画像イメージの添付ファイル付きメールです。";
  const id = DriveApp.getFolderById(folderId).getFilesByName(fileName+'.pdf').next().getId();
  //Googleドライブから画像イメージを取得する
  let attachImg = DriveApp.getFileById(id).getBlob();
  //オプションで添付ファイルを設定する
  let options = {
  "attachments":attachImg,
  };
  //MailAppで宛先、件名、本文、添付ファイルを引数にしてメールを送付
  MailApp.sendEmail(address, mailTitle, mailText, options);
}
