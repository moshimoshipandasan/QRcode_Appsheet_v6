/*
内容:QRコードリーダー
作成日:2021/11/24
作成者:青山学院中等部 安藤昇
備考：プログラムの配布・改変などする場合にはメールにてご連絡ください。
email:gigaschool2020@gmail.com（安藤宛）
*/
function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('QRコード', [
      {name: 'QRコード作成', functionName: 'QrlabelFunction'},
      {name: '集計準備', functionName: 'totalling0'},
    ]);
}

function QrlabelFunction(){
  var today = Utilities.formatDate(new Date(),"JST", "yyyy/MM/dd");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName("t氏名");
  var lastrow = sht.getLastRow(); // 最終行を取得
  var lastcol = sht.getLastColumn(); // 最終列を取得
  var range = sht.getRange(1, 1, lastrow, lastcol);
  var values = range.getValues(); // 情報をオンメモリに保持
  for (var i=1; i < lastrow; i++){
    data = values[i][0];
    if(data == ""){
      values[i][0] = getRndStr();
    }
    values[i][4] = false;
    // QRコードに含めるデータを取得 (A列の値)
    var qrData = values[i][0];
    // QR Server APIを使用してQRコード画像を生成する数式を作成
    // ENCODEURL関数でデータをURLエンコードする
    var qrc1 = '=image(\"https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=\"&ENCODEURL(\"' + qrData + '\"))';
    values[i][3] = qrc1;
  }
  range.setValues(values); //スプレッドシートに書き戻し
}

function getRndStr(){
  var str = "abcdefghijklmnopqrstuvwxyz0123456789";
  var len = 8;
  var result = "";
  for(var i=0;i<len;i++){
    result += str.charAt(Math.floor(Math.random() * str.length));
  }
  return result;
}

function totalling0() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht0 = ss.getSheetByName('tQR読込み');
  var sht1 = ss.getSheetByName('集計');
  sht1.clear();
  var lastrow0 = sht0.getLastRow();
//  var lastcol0 = sht0.getLastColumn();
  var range0 = sht0.getRange(1, 11, lastrow0, 5);
  var values0 = range0.getValues(); // 情報をオンメモリに保持
  var range1 = sht1.getRange(1, 1, lastrow0, 5);
  range1.setValues(values0); //スプレッドシートに書き戻し
}
