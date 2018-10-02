var CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('CHANNEL_ACCESS_TOKEN');
var MY_USER_ID = PropertiesService.getScriptProperties().getProperty('USER_ID');
var SS_ID = PropertiesService.getScriptProperties().getProperty('SS_ID');

var CHANNEL_ACCESS_TOKEN = 'ogcYgd4GW2fyc3bFOY/jYV59QoDz85GnC0FyDSD70BxDD3GO3oA5GymxROBSqyltCG0w2xa2TRotzS8FbqHbULEptuYrxK41Tv2e23Bl66t76KztnOVr3dtg7+twbNG0BJmWmjJUsDO3D2skPNVI2gdB04t89/1O/w1cDnyilFU=';
var MY_USER_ID = 'Ue5d8796e14aebe4bdc272cbb4af025c0';//自分のLINEのID
var SS_ID = '1SIstpbAZkycH5NqBiAeXNFJ-6BHcTzy2rJlgbM7T8yg';//スプレッドシートのID
var ss = SpreadsheetApp.openById(SS_ID);


//フォームの作成
function createForm(){
  var date_setting = ss.getSheetByName('概要').getRange(1,2,5,1).getValues();//概要のデータ取得
  Logger.log(date_setting)
  var formTitle = String(date_setting[1])+"月の鳩世話日程調整"; //フォームのタイトル
  var formDescription = date_setting[0]; //フォームの概要
  var week = ["(月)","(火)","(水)","(木)","(金)","(土)","(日)"]; //曜日
  var ss_form = ss.getSheetByName('概要').getRange(7,2);
  var FORM_ID_old = ss_form.getValue();//古いフォームのid取得
  var form = FormApp.create(formTitle);//：フォームの作成
  var FORM_ID = form.getId();//フォームのID取得
  var FORM_URL = form.getPublishedUrl();//作成したフォームのurlを作成
  var FOLDER_ID = '1_LdvtROHXW6jRDqMWuqpW-30OTT33eQ8';//保存するファイルのID
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
    var days = String(i)+'日'+week[(i-date_setting[2][0]+date_setting[4][0])%7];
    form.addScaleItem()
    .setTitle(days)
    .setBounds(1,3)
    .setLabels('行けないよ','行けるよ')
    .setRequired(true);
    if((i-date_setting[2][0]+1)%7==0){
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
  var FormUrl =createForm();
  var userGrade1 = ss.getSheetByName('ユーザー管理').getRange(1,1,16,1).getValues();
  var date = new Date();
  for(var i=0; i<16; i++){
    Logger.log(userGrade1[i]);
    if(date.getDate()<14){
      var postData = {
        'to':userGrade1[i][0],
        'messages':[{
          'type': 'text',
          'text':'今月の鳩世話シフトです\n14日の正午までに答えてください'
        },{
          'type': 'text',
          'text':FormUrl
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
        'to':userGrade1[i][0],
        'messages':[{
          'type': 'text',
          'text':'今月の鳩世話シフトです\n29日の正午までに答えてください'
        },{
          'type': 'text',
          'text':FormUrl
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
}

//シフトを自動で調整し、メールを送信する
function getAnswer(){
  var ss_answer = ss.getSheetByName('概要').getRange('C8');//シフトの結果の場所
  var date = new Date();
  var sheetName = date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate();
  var max = 0;//鳩世話できる度合いの最高値
  var l = 0;
  var count_c = 0;//確定者の数
  var count_c_past=0;
  var answer=String(ss.getSheetByName('概要').getRange(2,2).getValue())+"月の鳥シフト\n\n";
  var candidate = [];//各日の候補者の配列
  var cd=[];//候補者の二次配列
  var row_resp=[];//各日の行ける度合いの配列
  var confirm = [];//各日の確定案
  var count_w = 0;//何週目かどうか
  var sheet = ss.getSheets()[0];//最新の回答結果
  var lastrow = sheet.getLastRow();//行の数
  var lastcolumn = sheet.getLastColumn();//列の数
  var response = sheet.getDataRange().getValues();//フォームの回答のデータ
  var past_cd = [];
  var do_times = [];//鳩世話できる日の数
  var times=0;
  var count_n = 0;
  var cd_n;//候補者の数の配列
  var min_day=0;//候補者の数の最小値の日の配列
  var max_p=0;//最も出れる日の多い人
  var retry_check=0;//二週目の必要性の確認
  var min=15;
  sheet.setName(sheetName);// シート名を日付に変更
  
  //候補者を選出
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
  //候補者を絞るときに配列をいじるためコピーを取っておく。確定者が定まらない時はランダムに選定するため、そのコピーをとる
  for(var i=0; i<cd.length; i++){
    past_cd[i]=cd[i].slice();
  }
  var num_w = Math.floor(cd.length/7);//週の数
  //確定配列にあらかじめ0を日数分加えておく
  for (var j=0; j<cd.length; j++){
    confirm.push(0);
  }  
  
  //一人一人の鳩世話できる日の数を配列化
  for(i=1; i<lastrow; i++){
    times=0;
    for(j=0; j<cd.length; j++){
      count_n = cd[j].indexOf(i);
      if(count_n>=0){
        times+=1;
      }
    }
    do_times.push(times);  
  }
  Logger.log(do_times);
  Logger.log('~~振り分け開始~~');
  //候補者を絞る作業
  while(count_c!=cd.length){
    i=0;
    count_c_past=0;
    //候補者が一人ならば確定者に選び、他の日からその候補者を削除する
    while(count_c_past>=0){
      count_c_past=count_c;
      for(var j=0; j<cd.length; j++){     
        if(cd[j].length == 1){
          confirm[j] = cd[j][0];
          count_c+=1;
          Logger.log('一人確定しました  '+count_c+'/'+cd.length);
          Logger.log(confirm);
          for(k=0; k<cd.length; k++){
            for(l=0; l<cd[k].length; l++){
              if(cd[k][l]==confirm[j]){
                cd[k].splice(l,1);
              }
            }
          }
        }
      }
      if(count_c-count_c_past==0){
        count_c_past=-1;
      }
    }
    
    //それぞれの候補者の数の配列
    cd_n = [];
    for(j=0; j<cd.length; j++){
      cd_n.push(cd[j].length);
    }
    Logger.log(cd_n);
    
    
    //候補者の数が最も少ない日を選別
    min=15;
    for(i=0; i<cd.length; i++){
      if(cd_n[i]>1 && min>cd_n[i]){
        min=cd_n[i];
        min_day=i;
      }
    }
    Logger.log(min_day);
    
    //さらに最も多く行ける日がある人を選別
    max=0;
    for(i=0; i<cd[min_day].length; i++){
      if(do_times[cd[min_day][i]-1]>=max){
        max=do_times[cd[min_day][i]-1]
        max_p=cd[min_day][i];
      }
    }
    Logger.log(cd[min_day]);
    
    //候補者が二人以上の時に最も多く行ける日がある人を候補者に絞る
    for(i=cd_n[min_day]-1; i>=0; i-=1){
      if(cd[min_day][i]==max_p){
        continue;
      }   
      cd[min_day].splice(i,1)
      Logger.log(cd[min_day]);
    }
    Logger.log(cd);
    
    //二週目に入るか確認するためにcdの全ての長さを合計し、もし０ならば二週目開始
    retry_check=0;
    for(i=0; i<cd.length; i++){
      retry_check+=cd[i].length;
    }
    //1周目の確定選定が終わり、同じ人を選ばなくては行けなくなったときに2周目の確定選定を行う
    //ここでは先ほどとルールが違い、1周目の日付との差sが最も大きい人を選定し、候補から除外することで3周目の選定を防ぐ。 
    if(retry_check==0){
      Logger.log('~二週目~');
      for(i=0; i<cd.length; i++){
        data_length=[];
        if(confirm[i]==0){
          for(j=0; j<past_cd[i].length; j++){
            for(k=0; k<confirm.length; k++){ 
              if(past_cd[i][j]==confirm[k]){
                data_length.push(Math.pow((k-i), 2));
              }
            }
          }
          for(j=0; j<past_cd[i].length; j++){
            if(data_length[j]==Math.max.apply(null, data_length)){
              confirm[i]=past_cd[i][j];
              count_c+=1;
              Logger.log('一人確定しました  '+count_c+'/'+cd.length);
              Logger.log(confirm);
            } 
          }
          for(j=0; j<past_cd.length; j++){
            for(k=0; k<past_cd[j].length; k++){
              if(past_cd[j][k]==confirm[k]){
                past_cd[j].splice(k,1);
              }
            }
          }
        }
      }
    }
  } 
  Logger.log('~~振り分け完了~~');
  
  //それぞれの数字に対応する人の名前と日付と曜日を加えて
  for (var i=2; i<confirm.length+2; i++){
    answer += response[0][i]+response[confirm[i-2]][1]+"\n";
  }
  Logger.log(answer);
  ss_answer.setValue(answer);
  
  var postData = {
    'to':MY_USER_ID,
    'messages':[{
      'type': 'text',
      'text':answer,
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


//シフト結果を送信
function sendAnswer(){
  
  var ss_answer = ss.getSheetByName('概要').getRange('C8').getValue();//シフトの調整結果のデータ
  var userGrade1 = ss.getSheetByName('ユーザー管理').getRange(1,1,16,1).getValues();
  for(var i=0; i<16; i++){
    var postData = {
      'to':userGrade1[i][0],
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
}



