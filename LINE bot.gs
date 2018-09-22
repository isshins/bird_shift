//イベントの振り分け
function doPost(e) {
  
  var events = JSON.parse(e.postData.contents).events;
  events.forEach(function(event) {
    if(event.type == "message") {
      reply(event);
      addLog(event.source.userId, event.message.text);
    } else if(event.type == "follow") {
      follow(event.source.userId);
    } else if(event.type == "unfollow") {
      unFollow(event.source.userId);
    }
 });
  
}

    //ログの管理
function addLog(userid,message){
  
  var ss = SpreadsheetApp.openById(SS_ID);  //スプレッドシートの指定
  var sheet = ss.getSheetByName('ログ管理'); //シートを取得する
  var url = 'https://api.line.me/v2/bot/profile/' + userid;
  var response = UrlFetchApp.fetch(url, {
    'headers': {
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    }
  });
  var UserName = JSON.parse(response.getContentText()).displayName;
  sheet.appendRow([userid,UserName,message]); //ユーザーIDをシートに追加する
  
}


//オウム返し
function reply(e) {

  //返信先のデータ及び返信メッセージの指定
  var replyMessage = {
    "replyToken" : e.replyToken,
    "messages" : [
      {
        "type" : "text",
        "text" : ((e.message.type=="text") ? e.message.text : "Text以外は返せません・・・")
      }
    ]
  };

  //情報を詰めて、エンドポイントを蹴飛ばす
  var replyData = {
    "method" : "post",
    "headers" : {
      "Content-Type" : "application/json",
      "Authorization" : "Bearer " + CHANNEL_ACCESS_TOKEN
    },
    "payload" : JSON.stringify(replyMessage)
  };
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", replyData);

}

/* 友達追加されたらユーザーIDを登録する */
function follow(userid) {

  var ss = SpreadsheetApp.openById(SS_ID);  //スプレッドシートの指定
  var sheet = ss.getSheetByName('ユーザー管理'); //シートを取得する
  var url = 'https://api.line.me/v2/bot/profile/' + userid;
  var response = UrlFetchApp.fetch(url, {
    'headers': {
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    }
  });
  var UserName = JSON.parse(response.getContentText()).displayName;
  sheet.appendRow([userid,UserName]); //ユーザーIDをシートに追加する

}

/* アンフォローされたら削除する */
function unFollow(userid){


  var ss = SpreadsheetApp.openById(SS_ID); 
  var sheet = ss.getSheetByName('ユーザー管理'); //シートを取得する
  var result = findRow(sheet, userid, 2);
  if(result > 0){
    sheet.deleteRows(result);
  }

}

//列の検索
function findRow(sheet,val,col){

  var data = sheet.getDataRange().getValues(); 
  Logger.log(data);
  for(var i=0; i < data.length; i++){
    if(data[i][col-1] === val){
      return i+1;
    }
  }
  return 0;

}


//何か打ちたい時
function PushSomething(){
  
  var text = 'あ';
  var postData = {
    'to':MY_USER_ID,
    'messages':[{
      'type': 'text',
      'text':text,
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
