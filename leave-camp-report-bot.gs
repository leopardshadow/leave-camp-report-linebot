
var CHANNEL_ACCESS_TOKEN = TOKEN; 
var SHEET_ID = SHEET_ID; 
var SpreadSheet = SpreadsheetApp.openById(SHEET_ID);
var Sheet = SpreadSheet.getSheetByName("工作表1");


function doPost(e) {
 
  var msg = JSON.parse(e.postData.contents);
  
  try {
      
    // 取出 replayToken 和發送的訊息文字
    var replyToken = msg.events[0].replyToken;
    var userMessage = msg.events[0].message.text;
    
    if (typeof replyToken === 'undefined') {
      return;
    }
    
    //
    switch(userMessage) {
      case ':clear':
      case '/clear':
      case '清空':
        clear_all_report();
        send_msg(replyToken, '已清空');
        return;

      case '/list':
      case ':list':
        send_msg(replyToken, print_report_list());
        return;

      /*
      case '/help':
      case ':help':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, intro_text);
        return;
        */

      default:
        console.log('not special command')
    }
    
    //
    re = /報告班長(.*)\s日期[:：](\d{1,2}[\/\.]?\d{1,2})\s時間[:：](\d{1,2}[:：]?\d{1,2})\s(\d{3})[ ]?([\s\S]*)/
    //            1.           2:日期                           3.時間                  4.號碼。    5.內容  

    found = userMessage.match(re)

    if (found.length == undefined)
      throw "Format ERROR";


    num = parseInt(found[4], 10);
    
    // Sheet.getRange(2, 2).setValue(msg_info);
    // Sheet.getRange(3, 3).setValue(num.toString(10));
    
    // for( i=4 ; i<msg_split.length ; i++)
    //   msg_info = msg_info + '  ' + msg_split[i]
    
    msg_info = found[4] + " " + found[5].replace(/\n/g, ' ');

    
    if(num>=42 && num<=54)
      Sheet.getRange(num-41, 2).setValue(msg_info);
    else if(num == 166)
      Sheet.getRange(14, 2).setValue(msg_info);
    else
      throw "number out of range";
    
    Sheet.getRange(15, 2).setValue( print_report_list() );
    
    send_msg(replyToken, print_report_list());
    
  }
  catch(err) {
    console.log(err);
  }
}




function print_report_list() {
  
  var report = '';

  for(i = 1; i <= 14; i++){
    if (Sheet.getRange(i, 2).getValue() == '' )
      report = report + Sheet.getRange(i, 1).getValue() + '\r\n';
    else
      report = report + Sheet.getRange(i, 2).getValue() + '\r\n';
  }
  
  return report;
}



function clear_all_report() {
    
  for(i = 1; i <= 15; i++){
      Sheet.getRange(i, 2).setValue('');
    
  }
}


function send_msg(replyToken, text) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': [{
        'type': 'text',
        'text': text,
      }],
    }),
  });
  
}




