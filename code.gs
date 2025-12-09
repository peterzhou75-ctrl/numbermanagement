var SHEET_NAME = "工作表1";
var SALES_SHEET_NAME = "業務清單";
var LOG_SHEET_NAME = "異動記錄";
var WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxppaeG1P2zNqjGBjVwbNRp5Wx8EoivWNHuZCqu5vxJkT4N0Jhe7TveyV_gvl5EJIh8/exec"; // 記得重新部署後更新這裡

function doGet(e) {
  if (e.parameter.page == "manage") {
    return HtmlService.createTemplateFromFile('Manage').evaluate()
      .setTitle('測試門號管理系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('測試門號保留系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// 取得業務清單供下拉選單使用
function getSalesList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_SHEET_NAME);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var sales = [];
  // 從第 2 列開始
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      sales.push({ name: data[i][0], email: data[i][1] });
    }
  }
  return sales;
}

// 取得可用的號碼
function getAvailableNumbers(subject, count) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var available = [];
  
  // 限制只讀取到第 52 列 (如果資料少於 52 則讀全部)
  var limit = Math.min(data.length, 52);

  for (var i = 1; i < limit; i++) {
    // 檢查: 號碼申請主體(Col B[1]), 外線數量(Col C[2]), Title(Col E[4]) 是否為空
    if (data[i][1] == subject && data[i][2] == count && data[i][4] == "") {
      available.push(data[i][3]);
    }
  }
  return available;
}

// 提交預留請求
function reserveNumbers(formObject) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var numbersToReserve = formObject.selectedNumbers.split(',');
  
  var limit = Math.min(data.length, 52);

  for (var i = 1; i < limit; i++) {
    if (numbersToReserve.indexOf(String(data[i][3])) > -1) {
      sheet.getRange(i + 1, 5).setValue(formObject.clientName);
      sheet.getRange(i + 1, 6).setValue(formObject.ownerName);
      sheet.getRange(i + 1, 9).setValue(formObject.ownerEmail);
      sheet.getRange(i + 1, 10).setValue(formObject.startDate);
      sheet.getRange(i + 1, 11).setValue(formObject.endDate);
    }
  }
  return "預留成功！";
}

// 取得所有號碼資訊 (限制到 52 列)
function getAllNumbers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getDisplayValues(); 
  var allNumbers = [];
  
  var limit = Math.min(data.length, 52);

  for (var i = 1; i < limit; i++) {
    // 回傳所有資料，前端再根據 Owner 判斷是否可編輯
    // 若 Title (Col E[4]) 不為空，代表有人使用
    if (data[i][4] != "") {
      allNumbers.push({
        row: i + 1,
        number: data[i][3],
        title: data[i][4],
        owner: data[i][5], // 用於權限判斷
        start: data[i][9],
        end: data[i][10]
      });
    }
  }
  return allNumbers;
}

// 處理更新與寫入異動記錄
function processUpdates(items, operatorName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  var resultMessages = [];
  var timestamp = new Date();

  items.forEach(function(item) {
    // 再次驗證權限：確保操作者是該號碼的 Owner
    var currentOwner = sheet.getRange(item.row, 6).getValue();
    if (currentOwner != operatorName) {
      resultMessages.push("錯誤：您無權限修改號碼 " + item.number);
      return;
    }

    if (item.action == "extend") {
      sheet.getRange(item.row, 11).setValue(item.newEndDate);
      resultMessages.push("號碼 " + item.number + " 已延展");
      // 寫入 Log
      logSheet.appendRow([timestamp, item.number, item.title, operatorName, "延展測試至 " + item.newEndDate]);

    } else if (item.action == "end") {
      sheet.getRange(item.row, 5).clearContent(); // Title
      sheet.getRange(item.row, 6).clearContent(); // Owner
      sheet.getRange(item.row, 9).clearContent(); // Email
      sheet.getRange(item.row, 10).clearContent(); // Start
      sheet.getRange(item.row, 11).clearContent(); // End
      
      resultMessages.push("號碼 " + item.number + " 已結束");
      // 寫入 Log
      logSheet.appendRow([timestamp, item.number, item.title, operatorName, "結束測試"]);
    }
  });
  
  return resultMessages;
}

// 每日檢查 (更新：包含到期日顯示)
function sendDailyReminders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0); // 歸零時間
  
  // 限制只讀取到第 52 列 (如果資料少於 52 則讀全部)
  var limit = Math.min(data.length, 52);

  for (var i = 1; i < limit; i++) {
    var endDateRaw = data[i][10]; // K欄 Test End
    var title = data[i][4];       // E欄 Title
    var email = data[i][8];       // I欄 Email
    
    // 必須有結束日、標題、Email 才執行
    if (endDateRaw && title && email) {
      var endDate = new Date(endDateRaw);
      endDate.setHours(0, 0, 0, 0);
      
      // 格式化日期為 yyyy-MM-dd (例如 2023-10-25)
      var formattedDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

      // 計算差距天數
      var diffTime = endDate.getTime() - today.getTime();
      var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      // 判斷條件：剩餘 2 天 或 當天 (0 天)
      if (diffDays == 2 || diffDays == 0) {
        var statusText = (diffDays == 0) ? "今天到期！" : "即將到期 (剩 2 天)";
        var subject = "【測試號碼提醒】" + title + " " + statusText;
        var link = WEB_APP_URL + "?page=manage";
        
        // 內文增加「測試到期日」
        var body = title + " 的測試號碼 " + statusText + "。\n" +
                   "測試到期日：" + formattedDate + "\n\n" +
                   "請確認是否要延展或結束測試：\n" + 
                   "管理系統連結: " + link;
        
        GmailApp.sendEmail(email, subject, body);
        Logger.log("已發送提醒給: " + email + " 關於 " + title);
      }
    }
  }
}
