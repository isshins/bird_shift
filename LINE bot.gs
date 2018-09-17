function doPost(e) {
  var events = JSON.parse(e.postData.contents).events;
  events.forEach(function(event) {
    if(event.type == "message") {
      reply(event);
    } else if(event.type == "follow") {
      follow(event);
    } else if(event.type == "unfollow") {
      unFollow(event);
    }
 });
}

function reply(e) {
  //返信先のデータ及び返信メッセージの指定
  var message = {
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
    "payload" : JSON.stringify(message)
  };
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", replyData);
}