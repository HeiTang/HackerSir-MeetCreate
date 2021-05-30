function  onOpen(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [
    {name: "加入事件",functionName: "addEvents"}
  ];
  sheet.addMenu('建立日曆事件', menuItems);
}

function addEvents(){
  // 讀取「第七屆會議列表」
  var Meet = sheet_7th_Meet;
  var sheet_List = Meet.getSheetByName('第七屆會議列表');
  var range = sheet_List.getDataRange();
  var values = range.getValues();
  
  // 讀取「第七屆 社團幹部 - 20200630 版」中幹部 Mail，並合併成字串
  var Directory = sheet_7th_Directory;
  var values1 = Directory.getRange('通訊錄!G2:G26').getValues();
  var guests = 'heitang.me@gmail.com';
  for (var i = 0; i < values1.length; i++){
    guests = guests + ',' + values1[i][0];
  }
  
  
  for (var i = 1; i < values.length; i++){
    var Status = sheet_List.getRange(i+1,9).getValues();
    if (Status != '已發布'){
      // 建立 Google Calendar 事件
      var calendar = calendar_HackerSir;
      var eventTitle = '[' + values[i][2] + '] ' + values[i][1];
      var startTime = values[i][5];
      var endTime = values[i][6];
      var eventColor = 11;
      var eventTime = 60;
      var options = {description: values[i][3], location: values[i][4],sendInvites: true, guests:guests};
      var event = calendar.createEvent(eventTitle, startTime, endTime, options).setColor(eventColor).addPopupReminder(eventTime);
      var eventId = event.getId();
      
      // 變更狀態
      sheet_List.getRange(i+1,9).setValue('已發布');
      sheet_List.getRange(i+1,10).setValue(eventId);
      
      // 建立訊息格式
      var startTime_format = Utilities.formatDate(values[i][5], 'GMT+8' , 'yyyy/MM/dd HH:mm');
      var LineText = '';
      LineText = 
        '✗ 受文者：全體幹部(含顧問)\n' + 
        '✗ 主旨：' + eventTitle + '\n' +
        '✗ 時間：' + startTime_format + '\n' +
        '✗ 地點：' + values[i][4] + '\n' +
        '✗ 形式：' + values[i][3] + '\n\n' +
        '☀ 會議紀錄/大綱：' + '\n' +
        'https://hackmd.io/@HackerSir/第七屆會議記錄' + '\n\n' +
        '☀ 會議請假表單：' + '\n' +
        'https://url.hackersir.org/leave' + '\n\n' +
        '☀ 無法參與請『務必』填寫請假表單' + '\n\n' +
        '🔥 看完請至 GoMail 收信，並選擇自己的出席情況 ( ´Д`)y━･~~';
       
      sendToLine(LineText);
      sendToDiscord(LineText);
    }
  }
}

// 建立 Line Notify 通知
function sendToLine(LineText){
  var token = lineToken;
  var options = 
      {
        "method":"post",
        "headers":{"Authorization":"Bearer " + token},
        "payload":{'message':'\n' + LineText}
      };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify",options);
}

// 建立 Discord Webhook 通知
function sendToDiscord(LineText) {
  var discordUrl = discordToken;

  var options = {
        method: "post",
        payload:{content: '[會議公告]\n\n' + LineText} 
    };
  var response = UrlFetchApp.fetch(discordUrl, options);
}
