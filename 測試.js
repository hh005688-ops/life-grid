/**
 * 獨立測試函數：用來檢查到底是試算表還是行事曆在扯後腿
 */
function testDebug() {
  Logger.log("--- 開始診斷測試 ---");
  
  // 1. 測試讀取試算表
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("隨機排班");
    Logger.log("【檢查 1 - 試算表】分頁狀態: " + (sheet ? "成功找到「隨機排班」" : "❌ 找不到分頁！"));
    if (sheet) {
      Logger.log("【檢查 1 - 試算表】最後一列數: " + sheet.getLastRow());
    }
  } catch (e) {
    Logger.log("【檢查 1 - 試算表】發生錯誤: " + e.message);
  }
  
  // 2. 測試讀取 Google 行事曆
  try {
    const calendarId = "zh.taiwan#holiday@group.v.calendar.google.com";
    Logger.log("【檢查 2 - 行事曆】嘗試連線台灣節日行事曆...");
    const cal = CalendarApp.getCalendarById(calendarId);
    Logger.log("【檢查 2 - 行事曆】物件取得: " + (cal ? "成功" : "❌ 失敗 (可能無權限或ID失效)"));
    
    if (cal) {
      const events = cal.getEvents(new Date(2026, 8, 1), new Date(2026, 8, 30));
      Logger.log("【檢查 2 - 行事曆】成功抓取 2026/9 月事件數量: " + events.length);
    }
  } catch (e) {
    Logger.log("【檢查 2 - 行事曆】發生錯誤: " + e.message);
  }
  
  Logger.log("--- 診斷測試結束 ---");
}

/**
 * 專門用來測試日曆抓取結果的除錯函式
 */
function testFetchHolidays() {
  const targetYear = 2026;
  const startDate = new Date(targetYear, 0, 1);
  const endDate = new Date(targetYear, 11, 31, 23, 59, 59);
  
  Logger.log('=== 開始測試讀取日曆 ===');
  
  // 測試 1：直接用官方 ID 抓
  const holidayCalendarId = 'zh.taiwan#holiday@group.v.calendar.google.com';
  const cal = CalendarApp.getCalendarById(holidayCalendarId);
  if (cal) {
    Logger.log('成功找到官方台灣節假日日曆: ' + cal.getName());
    const events = cal.getEvents(startDate, endDate);
    Logger.log('找到的事件總數: ' + events.length);
    
    // 印出 9 月份的事件來瞧瞧
    events.forEach(function(e) {
      const d = e.getStartTime();
      if (d.getMonth() === 8) { // 9月 (0-indexed)
        Logger.log('日期: ' + d.getDate() + ' -> 事件: ' + e.getTitle());
      }
    });
  } else {
    Logger.log('【警告】找不到官方台灣節假日日曆 (CalendarApp.getCalendarById 回傳 null)');
  }
  
  // 測試 2：檢查帳號底下的所有日曆
  const allCals = CalendarApp.getAllCalendars();
  Logger.log('帳號底下的日曆總數: ' + allCals.length);
  allCals.forEach(function(c, index) {
    Logger.log(`[${index + 1}] 名稱: ${c.getName()} | ID: ${c.getId()}`);
  });
  
  Logger.log('=== 測試結束 ===');
}