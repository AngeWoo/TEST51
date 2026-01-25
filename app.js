/**
 * 系統初始化與權限檢查 (部署前請務必手動執行一次此函式)
 */
function initDrive() {
  Logger.log("Recruitment API 系統權限已啟動");
  Logger.log("目前時區：台北 (GMT+8)");
}

/**
 * 處理 GET 請求
 * 用於獲取題目、後台數據列表、候選人全方位詳情
 */
function doGet(e) {
  const action = e.parameter.action;
  const password = e.parameter.password;
  const userName = e.parameter.userName;

  if (action === 'getQuestions') return jsonResponse(getQuestionsData());
  if (action === 'getResults') return jsonResponse(getResultsData(password));
  if (action === 'getFullCandidateInfo') return jsonResponse(getFullCandidateInfo(userName, password));

  return HtmlService.createHtmlOutput("Recruitment API Service Online");
}

/**
 * 處理 POST 請求
 * 用於提交履歷、更新履歷、提交同意書與測驗答案
 */
function doPost(e) {
  let params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    params = e.parameter || {};
  }

  const action = params.action;
  try {
    // 提交新履歷
    if (action === 'submitResume') return jsonResponse(submitResumeData(params.resumeData));
    
    // 修改現有履歷 (背景執行更新)
    if (action === 'updateResume') return jsonResponse(updateResumeData(params.resumeData));
    
    // 提交簽名同意書
    if (action === 'submitConsent') return jsonResponse(submitConsentData(params.userName, params.signatureData));
    
    // 提交測驗答案
    if (action === 'submitAnswers') return jsonResponse(submitAnswersData(params.userName, params.answers));
    
    return jsonResponse({ success: false, message: '未知操作' });
  } catch (err) {
    return jsonResponse({ success: false, message: 'Server Error: ' + err.toString() });
  }
}

/**
 * 格式化 JSON 輸出
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 1. 取得測驗題目
 * 結構：A:ID, B:內容, C:跳過, D,E,F,G:選項
 */
function getQuestionsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Questions');
  const data = sheet.getDataRange().getValues();
  const questions = data.slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]),
    question: r[1],
    options: [r[3], r[4], r[5], r[6]]
  }));
  return { success: true, data: questions };
}

/**
 * 2. 提交個人履歷 (16 個欄位)
 */
function submitResumeData(resumeDataJson) {
  const d = typeof resumeDataJson === 'string' ? JSON.parse(resumeDataJson) : resumeDataJson;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resumes');
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  
  sheet.appendRow([
    ts, d.name, d.gender, d.birthday, d.phone, d.email, 
    d.address, d.edu, d.school, d.major, d.expYears, d.job, 
    d.expDetail, d.skills, d.intro, d.dept
  ]);
  return { success: true };
}

/**
 * 3. 更新履歷資料 (修正錯誤使用)
 * 根據姓名尋找 Resumes 表中最後一筆紀錄並覆寫
 */
function updateResumeData(resumeDataJson) {
  const d = typeof resumeDataJson === 'string' ? JSON.parse(resumeDataJson) : resumeDataJson;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resumes');
  const data = sheet.getDataRange().getValues();
  const searchName = String(d.name).trim();
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === searchName) {
      const row = i + 1;
      const newValues = [[
        ts, d.name, d.gender, d.birthday, d.phone, d.email, 
        d.address, d.edu, d.school, d.major, d.expYears, d.job, 
        d.expDetail, d.skills, d.intro, d.dept
      ]];
      sheet.getRange(row, 1, 1, 16).setValues(newValues);
      return { success: true };
    }
  }
  return { success: false, message: "找不到原始資料" };
}

/**
 * 4. 提交答案並自動計分與關聯
 */
function submitAnswersData(userName, answersJson) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchName = String(userName).trim();
  
  // A. 跨表找職位與單位 (Resumes L欄與P欄)
  const rSheet = ss.getSheetByName('Resumes');
  const rData = rSheet.getDataRange().getValues();
  let position = "未指定", department = "未指定";
  for (let j = rData.length - 1; j >= 1; j--) {
    if (String(rData[j][1]).trim() === searchName) {
      position = rData[j][11]; // L 欄 (Index 11)
      department = rData[j][15]; // P 欄 (Index 15)
      break;
    }
  }

  // B. 批改得分 (Answers A:ID, B:Ans, C:Score)
  const ansSheet = ss.getSheetByName('Answers');
  const ansData = ansSheet.getDataRange().getValues();
  const ansMap = {};
  for (let i = 1; i < ansData.length; i++) {
    ansMap[String(ansData[i][0]).trim()] = {
      correct: String(ansData[i][1]).trim(),
      score: Number(ansData[i][2]) || 10
    };
  }

  let totalScore = 0;
  const userAns = typeof answersJson === 'string' ? JSON.parse(answersJson) : answersJson;
  for (let qId in userAns) {
    if (ansMap[qId] && String(userAns[qId]).trim() === ansMap[qId].correct) {
      totalScore += ansMap[qId].score;
    }
  }

  // C. 寫入 Results (姓名, 職務, 單位, 分數, 時間, 答案紀錄)
  const resultsSheet = ss.getSheetByName('Results');
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  resultsSheet.appendRow([userName, position, department, totalScore, ts, JSON.stringify(userAns)]);

  return { success: true, score: totalScore };
}

/**
 * 5. 管理端：獲取數據列表
 */
function getResultsData(password) {
  if (password !== 'admin') return { success: false, message: '密碼錯誤' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  const data = sheet.getDataRange().getValues();
  
  return { 
    success: true, 
    data: data.slice(1).map(r => ({
      name: r[0],
      position: r[1],
      dept: r[2],
      score: r[3],
      timestamp: r[4],
      answers: r[5]
    })) 
  };
}

/**
 * 6. 管理端：獲取候選人全方位詳情 (履歷陣列 + 分數 + 簽名)
 */
function getFullCandidateInfo(userName, password) {
  if (password !== 'admin') return { success: false, message: '權限不足' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = { resumeArray: [], score: 0, signature: "", timestamp: "" };
  const cleanName = String(userName).trim();

  // A. 找履歷
  const rSheet = ss.getSheetByName('Resumes');
  const rData = rSheet.getDataRange().getValues();
  for (let i = rData.length - 1; i >= 1; i--) {
    if (String(rData[i][1]).trim() === cleanName) {
      res.resumeArray = rData[i]; break;
    }
  }

  // B. 找分數與考試時間
  const resultsData = ss.getSheetByName('Results').getDataRange().getValues();
  for (let i = resultsData.length - 1; i >= 1; i--) {
    if (String(resultsData[i][0]).trim() === cleanName) {
      res.score = resultsData[i][3];
      res.timestamp = resultsData[i][4];
      break;
    }
  }

  // C. 找簽名
  const cData = ss.getSheetByName('Consent').getDataRange().getValues();
  for (let i = cData.length - 1; i >= 1; i--) {
    if (String(cData[i][1]).trim() === cleanName) {
      res.signature = cData[i][3]; break;
    }
  }

  return { success: true, data: res };
}

/**
 * 7. 提交同意書與簽名
 */
function submitConsentData(userName, sig) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Consent');
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  sheet.appendRow([ts, userName, '已同意', sig]);
  return { success: true };
}