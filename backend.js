/**
 * 英文單字自動出題與測驗系統 - Google Apps Script 後端 (backend.js)
 * 
 * 部署說明：
 * 1. 在 Google 試算表中，點擊「擴充功能」->「Apps Script」。
 * 2. 將此檔案程式碼完整複製貼入 Apps Script 的 `Code.gs` 中。
 * 3. 執行一次 `setupSheet()` 函式以初始化試算表標頭與工作表結構。
 * 4. 點擊右上角「部署」->「新增部署」。
 * 5. 類型選擇「網頁應用程式 (Web App)」。
 * 6. 設定：
 *    - 說明：英文單字測驗系統 API
 *    - 執行身份：我 (Me)
 *    - 誰能存取：所有人 (Anyone) [重要：否則前端會發生 CORS 或權限錯誤]
 * 7. 部署後複製「網頁應用程式 URL」，貼至前端 index.html / admin.html 的 GAS 設定區域。
 */

// 初始化試算表結構
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 初始化成績紀錄表
  var scoreSheet = ss.getSheetByName("成績紀錄");
  if (!scoreSheet) {
    scoreSheet = ss.insertSheet("成績紀錄");
  }
  if (scoreSheet.getLastRow() === 0) {
    scoreSheet.appendRow([
      "測驗時間", // A欄
      "測驗回數", // B欄 (支援補考標記)
      "班級",     // C欄 (帳號@前字串第4、5碼)
      "座號",     // D欄 (帳號@前字串最後2碼)
      "姓名",     // E欄
      "測驗成績", // F欄
      "錯題記錄"  // G欄 (JSON格式或單字清單)
    ]);
    scoreSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e2e8f0");
  }

  // 2. 初始化試題與出題設定表
  var configSheet = ss.getSheetByName("出題設定");
  if (!configSheet) {
    configSheet = ss.insertSheet("出題設定");
  }
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow([
      "回數ID",       // A欄: e.g. R001
      "測驗名稱",     // B欄: e.g. 第一章 核心單字測驗
      "涵蓋單元",     // C欄: e.g. Unit 1, Unit 2
      "總題數",       // D欄: e.g. 10
      "單字題目清單", // E欄: JSON 字串 (含抽出的固定單字清單，供測驗與補考使用)
      "建立時間"      // F欄
    ]);
    configSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e2e8f0");
  }
}

// 處理 HTTP GET 請求 (取得測驗清單或指定回數的試題)
function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : "getQuizList";
    
    if (action === "getQuizList") {
      return responseJSON(getQuizList());
    } else if (action === "getQuizWords") {
      var roundId = e.parameter.roundId;
      return responseJSON(getQuizWords(roundId));
    }
    
    return responseJSON({ status: "success", message: "API 連線正常", action: action });
  } catch (error) {
    return responseJSON({ status: "error", message: error.toString() });
  }
}

// 處理 HTTP POST 請求 (提交成績與儲存出題設定)
function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }

    var action = data.action;

    if (action === "submitResult") {
      return responseJSON(saveScoreResult(data));
    } else if (action === "saveQuizConfig") {
      return responseJSON(saveQuizConfig(data));
    }

    return responseJSON({ status: "error", message: "未知的 action 指令: " + action });
  } catch (error) {
    return responseJSON({ status: "error", message: error.toString() });
  }
}

/**
 * 儲存測驗成績至「成績紀錄」試算表
 */
function saveScoreResult(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("成績紀錄");
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName("成績紀錄");
  }

  var account = (data.account || "").trim();
  var name = (data.name || "學生").trim();
  var roundTitle = (data.roundTitle || "測驗").trim();
  var isMakeUp = !!data.isMakeUp; // 是否為補考
  var score = Number(data.score) || 0;
  var wrongWords = data.wrongWords ? (typeof data.wrongWords === 'string' ? data.wrongWords : JSON.stringify(data.wrongWords)) : "無錯題";
  
  // 補考標記處理
  var finalRoundText = isMakeUp ? (roundTitle + " (補考)") : roundTitle;
  
  // 帳號解析：取帳號 @ 前字串
  var username = account;
  if (account.indexOf("@") !== -1) {
    username = account.split("@")[0];
  }
  
  // C欄：班級（取帳號@前字串第4、5碼，1-based 索引）
  var classCode = "";
  if (username.length >= 5) {
    classCode = username.substring(3, 5);
  } else {
    classCode = data.className || "00";
  }

  // D欄：座號（取帳號@前字串最後2碼）
  var seatNo = "";
  if (username.length >= 2) {
    seatNo = username.slice(-2);
  } else {
    seatNo = data.seatNo || "00";
  }

  var timestamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

  // 寫入列
  sheet.appendRow([
    timestamp,     // A欄: 測驗時間
    finalRoundText,// B欄: 測驗回數 (含補考標記)
    classCode,     // C欄: 班級
    seatNo,        // D欄: 座號
    name,          // E欄: 姓名
    score,         // F欄: 測驗成績
    wrongWords     // G欄: 錯題記錄
  ]);

  return {
    status: "success",
    message: "成績已成功登錄！",
    data: {
      timestamp: timestamp,
      round: finalRoundText,
      classCode: classCode,
      seatNo: seatNo,
      name: name,
      score: score
    }
  };
}

/**
 * 儲存教師建立的測驗設定至「出題設定」
 */
function saveQuizConfig(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("出題設定");
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName("出題設定");
  }

  var roundId = data.roundId || ("R" + String(new Date().getTime()).slice(-6));
  var title = data.title || "單字測驗";
  var units = Array.isArray(data.units) ? data.units.join(", ") : (data.units || "全單元");
  var questionCount = data.questionCount || 10;
  var wordListJSON = JSON.stringify(data.words || []);
  var timestamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    roundId,
    title,
    units,
    questionCount,
    wordListJSON,
    timestamp
  ]);

  return {
    status: "success",
    message: "出題設定已儲存成功！",
    roundId: roundId
  };
}

/**
 * 取得所有已發布的測驗回數清單
 */
function getQuizList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("出題設定");
  if (!sheet || sheet.getLastRow() <= 1) {
    return { status: "success", quizzes: [] };
  }

  var values = sheet.getDataRange().getValues();
  var quizzes = [];

  for (var i = 1; i < values.length; i++) {
    quizzes.push({
      roundId: values[i][0],
      title: values[i][1],
      units: values[i][2],
      questionCount: values[i][3],
      createdAt: values[i][5]
    });
  }

  return {
    status: "success",
    quizzes: quizzes
  };
}

/**
 * 取得指定測驗回數的題目單字清單 (確保補考時能讀取同一個抽出的固定單字)
 */
function getQuizWords(roundId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("出題設定");
  if (!sheet || sheet.getLastRow() <= 1) {
    return { status: "error", message: "找不到出題設定" };
  }

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(roundId)) {
      var words = [];
      try {
        words = JSON.parse(values[i][4]);
      } catch(e) {
        words = [];
      }
      return {
        status: "success",
        roundId: values[i][0],
        title: values[i][1],
        units: values[i][2],
        questionCount: values[i][3],
        words: words
      };
    }
  }

  return { status: "error", message: "找不到對應的回數ID: " + roundId };
}

/**
 * 輔助回應函式 (帶 CORS 標頭與 JSON 格式)
 */
function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
