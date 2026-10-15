/**
 * =================================================================
 * [模組] 台灣國定假日抓取模組
 * 職責：專門負責向公開 API 請求該年度國定假日資料
 * =================================================================
 */
/**function getTaiwanHolidaysMap(year) {
  try {
    const apiUrl = "https://api.pin-yi.me/taiwan-calendar/" + year;
    const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    
    if (response.getResponseCode() !== 200) return {};
    
    const holidayData = JSON.parse(response.getContentText());
    const holidayMap = {};
    
    for (let j = 0; j < holidayData.length; j++) {
      const item = holidayData[j];
      // 篩選出國定假日，且排除一般的星期六、星期日
      if (item.isHoliday === true && item.description && item.description !== "星期六" && item.description !== "星期日") {
        holidayMap[item.date_format] = item.description;
      }
    }
    
    return holidayMap;
    
  } catch (error) {
    Logger.log("抓取國定假日發生例外 (降級略過): " + error.message);
    return {};
  }
}*/