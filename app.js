/**
 * 權限與 API 初始化
 */
function initDrive() {
  Logger.log("API 系統權限已啟動，目前時區為台北 (GMT+8)");
}

function doGet(e) {
  const action = e.parameter.action;
  const password = e.parameter.password;
  const userName = e.parameter.userName;

  if (action === 'getQuestions') return jsonResponse(getQuestionsData());
  if (action === 'getResults') return jsonResponse(getResultsData(password));
  if (action === 'getFullCandidateInfo') return jsonResponse(getFullCandidateInfo(userName, password));

  return HtmlService.createHtmlOutput("Recruitment API Service is Online");
}

function doPost(e) {
  let params;
  try { params = JSON.parse(e.postData.contents); } catch (err) { params = e.parameter || {}; }
  const action = params.action;
  try {
    if (action === 'submitResume') return jsonResponse(submitResumeData(params.resumeData));
    if (action === 'submitConsent') return jsonResponse(submitConsentData(params.userName, params.signatureData));
    if (action === 'submitAnswers') return jsonResponse(submitAnswersData(params.userName, params.answers));
    return jsonResponse({ success: false, message: 'Unknown Action' });
  } catch (err) {
    return jsonResponse({ success: false, message: 'Server Error: ' + err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getQuestionsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Questions');
  const data = sheet.getDataRange().getValues();
  return { success: true, data: data.slice(1).filter(r => r[0]).map(r => ({ id: String(r[0]), question: r[1], options: [r[3], r[4], r[5], r[6]] })) };
}

function submitResumeData(json) {
  const d = typeof json === 'string' ? JSON.parse(json) : json;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resumes');
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  sheet.appendRow([ts, d.name, d.gender, d.birthday, d.phone, d.email, d.address, d.edu, d.school, d.major, d.expYears, d.job, d.expDetail, d.skills, d.intro, d.dept]);
  return { success: true };
}

function submitAnswersData(userName, ans) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rData = ss.getSheetByName('Resumes').getDataRange().getValues();
  let pos = "未指定", dept = "未指定";
  const searchName = String(userName).trim();
  for (let j = rData.length - 1; j >= 1; j--) {
    if (String(rData[j][1]).trim() === searchName) {
      pos = rData[j][11]; dept = rData[j][15]; break;
    }
  }
  const ansSheet = ss.getSheetByName('Answers');
  const ansData = ansSheet.getDataRange().getValues();
  const ansMap = {};
  for (let i = 1; i < ansData.length; i++) { ansMap[String(ansData[i][0]).trim()] = { correct: String(ansData[i][1]).trim(), score: Number(ansData[i][2]) || 10 }; }
  let score = 0;
  const userAns = typeof ans === 'string' ? JSON.parse(ans) : ans;
  for (let qId in userAns) { if (ansMap[qId] && String(userAns[qId]).trim() === ansMap[qId].correct) score += ansMap[qId].score; }
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  ss.getSheetByName('Results').appendRow([userName, pos, dept, score, ts, JSON.stringify(userAns)]);
  return { success: true, score: score };
}

function getResultsData(pw) {
  if (pw !== 'admin') return { success: false };
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results').getDataRange().getValues();
  return { success: true, data: data.slice(1).map(r => ({ name: r[0], position: r[1], dept: r[2], score: r[3], timestamp: r[4], answers: r[5] })) };
}

/**
 * 核心修正：強化搜尋邏輯
 */
function getFullCandidateInfo(userName, pw) {
  if (pw !== 'admin') return { success: false, message: "密碼驗證失敗" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = { resumeArray: [], score: 0, signature: "", timestamp: "" };
  
  const searchName = String(userName).trim(); // 強制去空白

  // 1. 搜尋履歷
  const rSheet = ss.getSheetByName('Resumes');
  const rData = rSheet.getDataRange().getValues();
  for (let i = rData.length - 1; i >= 1; i--) {
    const sheetName = String(rData[i][1]).trim();
    if (sheetName === searchName) {
      res.resumeArray = rData[i];
      break;
    }
  }

  // 2. 搜尋成績
  const resultsData = ss.getSheetByName('Results').getDataRange().getValues();
  for (let i = resultsData.length - 1; i >= 1; i--) {
    if (String(resultsData[i][0]).trim() === searchName) {
      res.score = resultsData[i][3];
      res.timestamp = resultsData[i][4];
      break;
    }
  }

  // 3. 搜尋簽名
  const cData = ss.getSheetByName('Consent').getDataRange().getValues();
  for (let i = cData.length - 1; i >= 1; i--) {
    if (String(cData[i][1]).trim() === searchName) {
      res.signature = cData[i][3];
      break;
    }
  }

  // 如果找不到履歷資料，回傳錯誤以利偵錯
  if (res.resumeArray.length === 0) {
    return { success: false, message: "在 Resumes 工作表中找不到姓名為 [" + searchName + "] 的資料" };
  }

  return { success: true, data: res };
}

function submitConsentData(name, sig) {
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Consent').appendRow([ts, name, '已同意', sig]);
  return { success: true };
}