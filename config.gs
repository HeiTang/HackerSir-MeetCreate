// Line Notify 、 Discord Webhook : 通知
var lineToken = '';
var discordToken = '';

// Excel 「逢甲大學黑客社 會議建立」 > 「第七屆會議列表」
var sheet_7th_Meet = SpreadsheetApp.openById('');
// Excel 「第七屆 社團幹部 - 20200630 版」 > 「通訊錄」
var sheet_7th_Directory = SpreadsheetApp.openById('');

// Google Calendar「黑客社幹部行事曆」
var calendar_HackerSir = CalendarApp.getCalendarById('');