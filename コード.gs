var CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('CHANNEL_ACCESS_TOKEN');
var USER_ID = PropertiesService.getScriptProperties().getProperty('USER_ID');
var SS_ID = PropertiesService.getScriptProperties().getProperty('SS_ID');



//フォームの作成
function createForm(){
  var ss = SpreadsheetApp.openById(SS_ID);//スプレッドシートの確保
  var date_setting = ss.getSheetByName('概要').getRange('B１:B5').getValues();//概要のデータ取得
  var formTitle = String(date_setting[0][0])+"月の鳩世話日程調整"; //フォームのタイトル
  var formDescription = date_setting[0][1]; //フォームの概要
  var week = ["(月)","(火)","(水)","(木)","(金)","(土)","(日)"]; //曜日
  var ss_form = ss.getSheetByName('概要').getRange(7,2);
  var FORM_ID_old = ss_form.getValue();//古いフォームのid取得
  var form = FormApp.create(formTitle);//：フォームの作成
  var FORM_ID = form.getId();//フォームのID取得
  var FORM_URL = form.getPublishedUrl();//作成したフォームのurlを作成
  var FOLDER_ID = '1EJVKXdQ_HJ6LyOYyM4hhF_t80vBYFM9V';//保存するファイルのID
　　　　var formFile = DriveApp.getFileById(FORM_ID);//保存するファイルの定義
  var FORM_FILE_old =DriveApp.getFileById(FORM_ID_old);//
  DriveApp.getFolderById(FOLDER_ID).addFile(formFile);//フォルダーに保存
  DriveApp.getFolderById(FOLDER_ID).removeFile(FORM_FILE_old);//
  DriveApp.getRootFolder().removeFile(formFile);//古いフォームを削除
  ss_form.setValue(FORM_ID);//古いフォームのidをスプレッドシートに残す
  form.setDescription(formDescription);//フォームに概要を加える
  
  //氏名の記入
  form.addTextItem()
   .setTitle('氏名')
   .setRequired(true);
  
  //シフトの記入
  for (var i=date_setting[2][0]; i<=date_setting[3][0]; i++){
    var days = String(i)+'日'+week[(i+(date_setting[4][0]-date_setting[2][0]))%7]
      form.addScaleItem()
      .setTitle(days)
      .setBounds(1,3)
      .setLabels('行けないよ','行けるよ')
      .setRequired(true);
    if((i-date_setting[2]+1)%7==0){
      form.addPageBreakItem();//7日ごとの改ページ
    }
 }
  //フォームの回答をスプレッドシートに紐付け
  form
  .setDestination(FormApp.DestinationType.SPREADSHEET, SS_ID);
  Logger.log(FORM_URL);
  return FORM_URL;
  }

//フォームのURLをbotに送信(push)
function sendForm(){
  var date = new Date();
  if(date.getDate()<14){
    var postData = {
      'to':USER_ID,
      'messages':[{
        'type': 'text',
        'text':'今月の鳩世話シフトです\n14日の正午までに答えてください\n'
      },{
        'type': 'text',
        'text':createForm()
      }]
      
    };
    
    var push_url = 'https://api.line.me/v2/bot/message/push';
    var headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    };
    
    var options = {
      'method': 'post',
      'headers': headers,
      'payload':JSON.stringify(postData),
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(push_url, options);
  }
  else{
    var postData = {
      'to':USER_ID,
      'messages':[{
        'type': 'text',
        'text':'今月の鳩世話シフトです\n３０日の正午までに答えてください\n'
      },{
        'type': 'text',
        'text':createForm()
      }]
      
    };
    
    var push_url = 'https://api.line.me/v2/bot/message/push';
    var headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    };
    
    var options = {
      'method': 'post',
      'headers': headers,
      'payload':JSON.stringify(postData),
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(push_url, options);
  }
}

//シフトを自動で調整し、メールを送信する
function getAnswer(){
  var ss = SpreadsheetApp.openById(SS_ID);//
  var ss_answer = ss.getSheetByName('概要').getRange('C8');//シフトの結果の場所
　　　　var date = new Date();
  var sheetName = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();//
  var max = 0;//鳩世話できる度合いの最高値
  var l = 0;
  var count_c = 0;//確定者の数
  var count_c_past;
  var answer=String(ss.getSheetByName('概要').getRange(2,2).getValue())+"月の鳥シフト\n\n";
  var candidate = [];//各日の候補者の配列
  var cd=[];//候補者の二次配列
  var row_resp=[];//各日の行ける度合いの配列
  var confirm = [];//各日の確定案
  var count_w = 0;//何週目かどうか
  var roop= 0;// ループ変数
  var week_divide = 7;
  var sheet = ss.getSheets()[0];//最新の回答結果
  var lastrow = sheet.getLastRow();//行の数
  var lastcolumn = sheet.getLastColumn();//列の数
  var response = sheet.getDataRange().getValues();//フォームの回答のデータ
  var past_cd = [];
  var do_times = [];//鳩世話できる日の数
  var times=[];
  var count_n = 0;
  Logger.log(response);
  sheet.setName(sheetName);// シート名を日付に変更
  for (var j=2; j<lastcolumn; j++){　
    candidate = [];
    row_resp = [];
    //各日の行ける度合いが最高値の人を、候補者配列に加える
    for (var k=1; k<lastrow; k++){
      row_resp.push(response[k][j]);
    }
    max = Math.max.apply(null,row_resp);
    for (k=1; k<lastrow; k++){ 
      if (response[k][j] == max){
        candidate.push(k);
      }
    }
    cd.push(candidate);
  }
  Logger.log(cd);
//候補者を絞るときに配列をいじるためコピーを取っておく。候補がどうしても割れない時はランダムにするときのためのコピー
  for(var i=0; i<cd.length;　i++){
    past_cd[i]=cd[i].slice();
  }
  var num_w = Math.floor(cd.length/7);//週の数
  //確定配列にあらかじめ0を日数分加えておく
  for (var j=0; j<cd.length; j++){
    confirm.push(0);
  }
  
 //一人一人の鳩世話できる日の数を配列化
  for(i=1; i<lastrow; i++){
    times=[];
    for(j=0; j<cd.length; j++){
      count_n = cd[j].indexOf(i);
      if(count_n>=0){
        times.push(j)
      }
    }
    do_times.push(times);  
      }
  Logger.log(do_times);
  
   //候補者を絞る作業
  while(count_c!=cd.length){
    //候補者が一人ならば確定者に選び、他の日からその候補者を削除する
    while(count_c>=0){
      count_c_past=count_c;
      for(var j=0; j<cd.length; j++){     
        if(cd[j].length == 1){
          count_c+=1;
          roop+=1;
          confirm[j] = cd[j][0];
          for(k=0; k<cd.length; k++){
            for(l=0; l<cd[k].length; l++){
              if(cd[k][l]==confirm[j]){
                cd[k].splice(l,1);
              }
            }
          }
        }
      }
      Logger.log(confirm);
      if(count_c-count_c_past==0){
        count_c=-1;
      }
    }
    for(j=0; j<cd.length; j++){
      if(cd[j].length == 2){
        if(do_times[cd[j][0]-1].length > do_times[cd[j][1]-1].length){
          cd[j].splice(cd[j][1]);
        }else if(do_times[cd[j][1]-1].length > do_times[cd[j][0]-1].length){
          cd[j].splice(cd[j][0]);
        }else{
          cd[j].splice(cd[j][Math.floor(Math.random())]);
        }
      }
    }
    roop+=1;
    if(roop>confirm.length){
      for(j=0; j<cd.length; j++){  
        if(confirm[j]==0){
          confirm[j] = past_cd[j][Math.floor(Math.random*(past_cd[j].length-1))]
        }
      }
    }
    Logger.log(confirm);
    }
  
      
//それぞれの数字に対応する人の名前と日付と曜日を加えて
 for (var i=2; i<confirm.length+2; i++){
    answer += response[0][i]+response[confirm[i-2]][1]+"\n";
  }
  Logger.log(answer);
  ss_answer.setValue(answer);
  //MailApp.sendEmail('pikatyu112@gmail.com', '1年　鳩シフト完成', answer);
}
//シフトの調整結果をbotに送る前に自分自身にメールする。　ここはそのうちメールじゃなくてLINEbotで自分自身だけに結果を送ってok OR noopで全体に送るか決められるようにしたい
//ここは5分毎にトリガーがかかってるから未読のメールがきたら検索に引っかかって動き出す仕組み
//本来はメールのトリガーなんてないけど頑張って作った
function sendAnswer(){
  var ss = SpreadsheetApp.openById(SS_ID);
  var searchTitle = '(is:unread "1年　鳩シフト完成")';//未読のメールを検索
  var consentMessage = 'ok';
  var retryMessage = 'no';
  var ss_answer = ss.getSheetByName('概要').getRange('C8').getValue();//シフトの調整結果のデータ
  var myThread = GmailApp.search(searchTitle, 0, 1)[0];
  var myMessage = GmailApp.getMessagesForThread(myThread);
  var myResponse = myMessage[1].getPlainBody();//返信したときの本文
  Logger.log(myMessage[1].getPlainBody());
  //メールにokと返信すればbotにシフトの結果を送信する
  if (myResponse.indexOf(retryMessage) != -1){
    getAnswer();
  }
  //メールにnoと返信すればもう一度getAnswerしてくれる。
  if (myResponse.indexOf(consentMessage) != -1){
    var postData = {
      'to':USER_ID,
      'messages':[{
      'type': 'text',
      'text':ss_answer,
      }]
    };
  
    var push_url = 'https://api.line.me/v2/bot/message/push';
    var headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    };
  
    var options = {
      'method': 'post',
      'headers': headers,
      'payload':JSON.stringify(postData),
      'muteHttpExceptions': true
    };
  
    var response = UrlFetchApp.fetch(push_url, options); 
  }
  myMessage[1].markRead(); //メールを既読にする
}

