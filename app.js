/**
 * 炎洲集團人才招募 E 化系統 - Google Apps Script 後端
 * 版本：4.0（Phase 2 提升效率）
 *
 * Phase 1 功能：面試評分、錄取決策、候選人狀態追蹤
 * Phase 2 新增：
 * - 面試排程管理（Schedules Sheet）
 * - 批次通知發送（Notifications Sheet + MailApp）
 * - 一鍵歸檔 PDF（Google Drive PDF 生成）
 */

// ==================== 系統設定 ====================

function getAdminPassword() {
  try {
    const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    return pw || 'admin';
  } catch (e) {
    return 'admin';
  }
}

function getInterviewerPassword() {
  try {
    const pw = PropertiesService.getScriptProperties().getProperty('INTERVIEWER_PASSWORD');
    return pw || 'interviewer';
  } catch (e) {
    return 'interviewer';
  }
}

function setAdminPassword() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', 'your_secure_password_here');
  Logger.log("管理密碼已更新");
}

function setInterviewerPassword() {
  PropertiesService.getScriptProperties().setProperty('INTERVIEWER_PASSWORD', 'your_interviewer_password_here');
  Logger.log("面試官密碼已更新");
}

function initDrive() {
  Logger.log("Recruitment API 系統權限已啟動 v4.0");
  Logger.log("目前時區：台北 (GMT+8)");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Resumes', 'Questions', 'Answers', 'Results', 'Consent', 'AuditLog', 'Interviews', 'Decisions', 'Schedules', 'Notifications'];
  requiredSheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
      Logger.log("已自動建立工作表：" + name);
    }
  });
}

// ==================== 輸入驗證與清理 ====================

function sanitizeInput(str) {
  if (str == null) return '';
  return String(str).trim().replace(/[<>]/g, '').substring(0, 5000);
}

function isValidEmail(email) {
  if (!email) return true;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isValidPhone(phone) {
  if (!phone) return true;
  return /^[0-9\-\+\s]{7,15}$/.test(phone);
}

function validateResumeData(data) {
  const errors = [];
  if (!data.name || !data.name.trim()) errors.push('姓名為必填');
  if (!data.phone || !data.phone.trim()) errors.push('電話為必填');
  if (!data.email || !data.email.trim()) errors.push('Email 為必填');
  if (!data.job || !data.job.trim()) errors.push('應徵職務為必填');
  if (!data.dept || !data.dept.trim()) errors.push('需求單位為必填');
  if (data.email && !isValidEmail(data.email)) errors.push('Email 格式不正確');
  if (data.phone && !isValidPhone(data.phone)) errors.push('電話格式不正確');
  return errors;
}

// ==================== 審計日誌 ====================

function logAudit(action, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('AuditLog');
    if (!sheet) {
      sheet = ss.insertSheet('AuditLog');
      sheet.appendRow(['時間', '操作', '詳情']);
    }
    const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm:ss");
    sheet.appendRow([ts, action, detail]);
  } catch (e) {
    Logger.log("審計日誌寫入失敗：" + e.toString());
  }
}

// ==================== 共用查找工具 ====================

function findCandidate(data, appId, name, appIdCol, nameCol) {
  for (let i = data.length - 1; i >= 1; i--) {
    if (appId && String(data[i][appIdCol] || '').trim() === String(appId).trim()) return i;
  }
  if (!appId && name) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][nameCol] || '').trim() === String(name).trim()) return i;
    }
  }
  return -1;
}

// ==================== ID 生成工具 ====================

function generateApplicationId() {
  var now = new Date();
  var ts = Utilities.formatDate(now, "GMT+8", "yyMMddHHmm");
  var rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return 'APP-' + ts + '-' + rand;
}

function generateScheduleId() {
  var now = new Date();
  var ts = Utilities.formatDate(now, "GMT+8", "yyMMddHHmm");
  var rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return 'SCH-' + ts + '-' + rand;
}

function generateNotificationId() {
  var now = new Date();
  var ts = Utilities.formatDate(now, "GMT+8", "yyMMddHHmm");
  var rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return 'NTF-' + ts + '-' + rand;
}

// ==================== API 路由 ====================

function doGet(e) {
  const action = e.parameter.action;
  const password = e.parameter.password;
  const userName = e.parameter.userName;
  const applicationId = e.parameter.applicationId;
  const query = e.parameter.query;

  try {
    if (action === 'getQuestions') return jsonResponse(getQuestionsData());
    if (action === 'getResults') return jsonResponse(getResultsData(password));
    if (action === 'getFullCandidateInfo') return jsonResponse(getFullCandidateInfo(userName, password, applicationId));

    // Phase 1
    if (action === 'searchCandidates') return jsonResponse(searchCandidatesForInterview(password, query));
    if (action === 'getInterviewData') return jsonResponse(getInterviewData(password));

    // Phase 2
    if (action === 'getSchedules') return jsonResponse(getSchedulesData(password));
    if (action === 'getInterviewerSchedule') return jsonResponse(getInterviewerScheduleData(password, e.parameter.interviewerName));
    if (action === 'getNotificationLog') return jsonResponse(getNotificationLogData(password));
    if (action === 'getArchiveLog') return jsonResponse(getArchiveLogData(password));
    if (action === 'getEmailQuota') return jsonResponse(getEmailQuotaData(password));

    return HtmlService.createHtmlOutput("Recruitment API Service Online - v4.0");
  } catch (err) {
    logAudit('ERROR', 'doGet: ' + err.toString());
    return jsonResponse({ success: false, message: '伺服器錯誤，請稍後再試' });
  }
}

function doPost(e) {
  let params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    params = e.parameter || {};
  }

  const action = params.action;
  try {
    if (action === 'submitResume') return jsonResponse(submitResumeData(params.resumeData));
    if (action === 'updateResume') return jsonResponse(updateResumeData(params.resumeData));
    if (action === 'submitConsent') return jsonResponse(submitConsentData(params.userName, params.signatureData, params.applicationId));
    if (action === 'submitAnswers') return jsonResponse(submitAnswersData(params.userName, params.answers, params.applicationId));

    // Phase 1
    if (action === 'submitInterview') return jsonResponse(submitInterviewData(params));
    if (action === 'submitDecision') return jsonResponse(submitDecisionData(params));

    // Phase 2
    if (action === 'createSchedule') return jsonResponse(createScheduleData(params));
    if (action === 'updateSchedule') return jsonResponse(updateScheduleData(params));
    if (action === 'deleteSchedule') return jsonResponse(deleteScheduleData(params));
    if (action === 'sendNotifications') return jsonResponse(sendNotificationsData(params));
    if (action === 'generatePdf') return jsonResponse(generatePdfData(params));
    if (action === 'generateBulkPdf') return jsonResponse(generateBulkPdfData(params));

    return jsonResponse({ success: false, message: '未知操作' });
  } catch (err) {
    logAudit('ERROR', 'doPost [' + action + ']: ' + err.toString());
    return jsonResponse({ success: false, message: '伺服器錯誤，請稍後再試' });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== 原有核心功能 ====================

function getQuestionsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Questions');
  if (!sheet) return { success: false, message: '題庫尚未設定' };
  const data = sheet.getDataRange().getValues();
  const questions = data.slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]),
    question: String(r[1]),
    options: [r[3], r[4], r[5], r[6]].map(o => o ? String(o) : '')
  }));
  if (questions.length === 0) return { success: false, message: '尚無測驗題目' };
  return { success: true, data: questions };
}

function submitResumeData(resumeDataJson) {
  const d = typeof resumeDataJson === 'string' ? JSON.parse(resumeDataJson) : resumeDataJson;
  const errors = validateResumeData(d);
  if (errors.length > 0) return { success: false, message: '資料驗證失敗：' + errors.join('、') };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resumes');
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  const appId = d.applicationId || generateApplicationId();

  if (d.email) {
    const existingData = sheet.getDataRange().getValues();
    const now = new Date();
    for (let i = existingData.length - 1; i >= 1; i--) {
      if (String(existingData[i][5]).trim().toLowerCase() === String(d.email).trim().toLowerCase()) {
        const rowDate = new Date(existingData[i][0]);
        if ((now - rowDate) / (1000 * 60 * 60) < 24) {
          return { success: true, applicationId: existingData[i][16] || '', duplicate: true, message: '您已於 24 小時內提交過申請' };
        }
      }
    }
  }

  sheet.appendRow([
    ts, sanitizeInput(d.name), sanitizeInput(d.gender), sanitizeInput(d.birthday),
    sanitizeInput(d.phone), sanitizeInput(d.email), sanitizeInput(d.address),
    sanitizeInput(d.edu), sanitizeInput(d.school), sanitizeInput(d.major),
    sanitizeInput(d.expYears), sanitizeInput(d.job), sanitizeInput(d.expDetail),
    sanitizeInput(d.skills), sanitizeInput(d.intro), sanitizeInput(d.dept), appId
  ]);
  logAudit('SUBMIT_RESUME', '姓名：' + sanitizeInput(d.name) + '，追蹤碼：' + appId);
  return { success: true, applicationId: appId };
}

function updateResumeData(resumeDataJson) {
  const d = typeof resumeDataJson === 'string' ? JSON.parse(resumeDataJson) : resumeDataJson;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resumes');
  const data = sheet.getDataRange().getValues();
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  const appId = d.applicationId;

  const idx = findCandidate(data, appId, d.name, 16, 1);
  if (idx === -1) return { success: false, message: "找不到原始資料" };

  sheet.getRange(idx + 1, 1, 1, 17).setValues([[
    ts, sanitizeInput(d.name), sanitizeInput(d.gender), sanitizeInput(d.birthday),
    sanitizeInput(d.phone), sanitizeInput(d.email), sanitizeInput(d.address),
    sanitizeInput(d.edu), sanitizeInput(d.school), sanitizeInput(d.major),
    sanitizeInput(d.expYears), sanitizeInput(d.job), sanitizeInput(d.expDetail),
    sanitizeInput(d.skills), sanitizeInput(d.intro), sanitizeInput(d.dept),
    appId || data[idx][16] || ''
  ]]);
  logAudit('UPDATE_RESUME', '姓名：' + sanitizeInput(d.name) + '，追蹤碼：' + (appId || '(依姓名)'));
  return { success: true };
}

function submitAnswersData(userName, answersJson, applicationId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchName = sanitizeInput(userName);
  const appId = applicationId || '';

  const rData = ss.getSheetByName('Resumes').getDataRange().getValues();
  let position = "未指定", department = "未指定";
  const rIdx = findCandidate(rData, appId, searchName, 16, 1);
  if (rIdx > 0) {
    position = rData[rIdx][11] || '未指定';
    department = rData[rIdx][15] || '未指定';
  }

  const resultsSheet = ss.getSheetByName('Results');
  const existingResults = resultsSheet.getDataRange().getValues();
  for (let i = existingResults.length - 1; i >= 1; i--) {
    if (appId && String(existingResults[i][6] || '').trim() === appId) {
      return { success: true, score: existingResults[i][3], duplicate: true, message: '此申請已提交過測驗' };
    }
  }

  const ansData = ss.getSheetByName('Answers').getDataRange().getValues();
  const ansMap = {};
  for (let i = 1; i < ansData.length; i++) {
    ansMap[String(ansData[i][0]).trim()] = { correct: String(ansData[i][1]).trim(), score: Number(ansData[i][2]) || 10 };
  }

  let totalScore = 0, totalPossible = 0;
  const userAns = typeof answersJson === 'string' ? JSON.parse(answersJson) : answersJson;
  for (let qId in ansMap) totalPossible += ansMap[qId].score;
  for (let qId in userAns) {
    if (ansMap[qId] && String(userAns[qId]).trim() === ansMap[qId].correct) totalScore += ansMap[qId].score;
  }

  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  resultsSheet.appendRow([searchName, position, department, totalScore, ts, JSON.stringify(userAns), appId]);
  logAudit('SUBMIT_ANSWERS', '姓名：' + searchName + '，分數：' + totalScore + '/' + totalPossible + '，追蹤碼：' + appId);
  return { success: true, score: totalScore };
}

function getResultsData(password) {
  if (password !== getAdminPassword()) return { success: false, message: '密碼錯誤' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  if (!sheet) return { success: true, data: [] };
  const data = sheet.getDataRange().getValues();
  logAudit('VIEW_RESULTS', '管理員查閱了候選人列表');
  return {
    success: true,
    data: data.slice(1).map(r => ({
      name: r[0], position: r[1], dept: r[2], score: r[3],
      timestamp: r[4], answers: r[5], applicationId: r[6] || ''
    }))
  };
}

function getFullCandidateInfo(userName, password, applicationId) {
  if (password !== getAdminPassword()) return { success: false, message: '權限不足' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = { resumeArray: [], score: 0, signature: "", timestamp: "", applicationId: applicationId || "", interview: null, decision: null, email: "" };
  const cleanName = sanitizeInput(userName);
  const appId = applicationId || '';

  const rData = ss.getSheetByName('Resumes').getDataRange().getValues();
  const rIdx = findCandidate(rData, appId, cleanName, 16, 1);
  if (rIdx > 0) { res.resumeArray = rData[rIdx]; res.applicationId = rData[rIdx][16] || appId; res.email = String(rData[rIdx][5] || ''); }

  const resultsData = ss.getSheetByName('Results').getDataRange().getValues();
  const resIdx = findCandidate(resultsData, appId, cleanName, 6, 0);
  if (resIdx > 0) { res.score = resultsData[resIdx][3]; res.timestamp = resultsData[resIdx][4]; }

  const cSheet = ss.getSheetByName('Consent');
  if (cSheet) {
    const cData = cSheet.getDataRange().getValues();
    const cIdx = findCandidate(cData, appId, cleanName, 4, 1);
    if (cIdx > 0) res.signature = cData[cIdx][3];
  }

  const iSheet = ss.getSheetByName('Interviews');
  if (iSheet) {
    const iData = iSheet.getDataRange().getValues();
    for (let i = iData.length - 1; i >= 1; i--) {
      const matchById = appId && String(iData[i][1] || '').trim() === appId;
      const matchByName = !appId && String(iData[i][2] || '').trim() === cleanName;
      if (matchById || matchByName) {
        res.interview = {
          interviewerName: iData[i][3],
          scores: iData[i][4] ? (typeof iData[i][4] === 'string' ? JSON.parse(iData[i][4]) : iData[i][4]) : {},
          totalAvg: iData[i][5],
          recommendation: iData[i][6],
          notes: iData[i][7],
          timestamp: iData[i][0]
        };
        break;
      }
    }
  }

  const dSheet = ss.getSheetByName('Decisions');
  if (dSheet) {
    const dData = dSheet.getDataRange().getValues();
    const dIdx = findCandidate(dData, appId, cleanName, 1, 2);
    if (dIdx > 0) {
      res.decision = { decision: dData[dIdx][3], combinedScore: dData[dIdx][4], decidedBy: dData[dIdx][5], timestamp: dData[dIdx][0] };
    }
  }

  logAudit('VIEW_CANDIDATE', '管理員查閱：' + cleanName + '，追蹤碼：' + (appId || '(依姓名)'));
  return { success: true, data: res };
}

function submitConsentData(userName, sig, applicationId) {
  if (!userName || !sig) return { success: false, message: '缺少必要資料' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Consent');
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  sheet.appendRow([ts, sanitizeInput(userName), '已同意', sig, applicationId || '']);
  logAudit('SUBMIT_CONSENT', '姓名：' + sanitizeInput(userName) + '，追蹤碼：' + (applicationId || ''));
  return { success: true };
}

// ==================== Phase 1：面試評分系統 ====================

function searchCandidatesForInterview(password, query) {
  if (password !== getInterviewerPassword() && password !== getAdminPassword()) {
    return { success: false, message: '密碼錯誤' };
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultsData = ss.getSheetByName('Results').getDataRange().getValues();
  const searchTerm = query ? String(query).trim().toLowerCase() : '';
  const candidates = [];
  for (let i = 1; i < resultsData.length; i++) {
    const name = String(resultsData[i][0] || '');
    const appId = String(resultsData[i][6] || '');
    if (searchTerm && !name.toLowerCase().includes(searchTerm) && !appId.toLowerCase().includes(searchTerm)) continue;
    candidates.push({ name: name, position: String(resultsData[i][1] || ''), dept: String(resultsData[i][2] || ''), examScore: Number(resultsData[i][3]) || 0, timestamp: resultsData[i][4], applicationId: appId });
  }
  return { success: true, data: candidates };
}

function submitInterviewData(params) {
  const pw = params.interviewerPassword || params.password;
  if (pw !== getInterviewerPassword() && pw !== getAdminPassword()) {
    return { success: false, message: '面試官密碼錯誤' };
  }
  const appId = params.applicationId || '';
  const candidateName = sanitizeInput(params.candidateName);
  const interviewerName = sanitizeInput(params.interviewerName);
  const recommendation = params.recommendation;
  const notes = sanitizeInput(params.notes);
  if (!candidateName) return { success: false, message: '請選擇候選人' };
  if (!interviewerName) return { success: false, message: '請輸入面試官姓名' };

  const validDimensions = ['professional', 'communication', 'teamwork', 'potential', 'stability'];
  let scores;
  try { scores = typeof params.scores === 'string' ? JSON.parse(params.scores) : params.scores; }
  catch (e) { return { success: false, message: '評分資料格式錯誤' }; }

  let total = 0;
  for (let d = 0; d < validDimensions.length; d++) {
    const dim = validDimensions[d];
    const val = Number(scores[dim]);
    if (!val || val < 1 || val > 5 || val !== Math.floor(val)) return { success: false, message: '每個評分維度須為 1-5 的整數' };
    total += val;
  }
  const totalAvg = Math.round((total / 5) * 10) / 10;

  const validRecs = ['強烈推薦', '推薦', '待議', '不推薦'];
  if (!validRecs.includes(recommendation)) return { success: false, message: '請選擇錄用建議' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const iSheet = ss.getSheetByName('Interviews') || ss.insertSheet('Interviews');
  const iData = iSheet.getDataRange().getValues();
  for (let i = iData.length - 1; i >= 1; i--) {
    const sameCandidate = (appId && String(iData[i][1] || '').trim() === appId) || (!appId && String(iData[i][2] || '').trim() === candidateName);
    const sameInterviewer = String(iData[i][3] || '').trim() === interviewerName;
    if (sameCandidate && sameInterviewer) {
      const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
      iSheet.getRange(i + 1, 1, 1, 8).setValues([[ts, appId, candidateName, interviewerName, JSON.stringify(scores), totalAvg, recommendation, notes]]);
      logAudit('UPDATE_INTERVIEW', '面試官：' + interviewerName + '，候選人：' + candidateName + '，均分：' + totalAvg);
      return { success: true, totalAvg: totalAvg, updated: true };
    }
  }
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  iSheet.appendRow([ts, appId, candidateName, interviewerName, JSON.stringify(scores), totalAvg, recommendation, notes]);
  logAudit('SUBMIT_INTERVIEW', '面試官：' + interviewerName + '，候選人：' + candidateName + '，均分：' + totalAvg + '，建議：' + recommendation);
  return { success: true, totalAvg: totalAvg };
}

function getInterviewData(password) {
  if (password !== getAdminPassword()) return { success: false, message: '密碼錯誤' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const interviews = {}, decisions = {};

  const iSheet = ss.getSheetByName('Interviews');
  if (iSheet) {
    const iData = iSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      const key = String(iData[i][1] || '').trim() || String(iData[i][2] || '').trim();
      if (!key) continue;
      interviews[key] = {
        applicationId: String(iData[i][1] || ''), candidateName: String(iData[i][2] || ''),
        interviewerName: String(iData[i][3] || ''),
        scores: iData[i][4] ? (typeof iData[i][4] === 'string' ? JSON.parse(iData[i][4]) : iData[i][4]) : {},
        totalAvg: Number(iData[i][5]) || 0, recommendation: String(iData[i][6] || ''),
        notes: String(iData[i][7] || ''), timestamp: iData[i][0]
      };
    }
  }

  const dSheet = ss.getSheetByName('Decisions');
  if (dSheet) {
    const dData = dSheet.getDataRange().getValues();
    for (let i = 1; i < dData.length; i++) {
      const key = String(dData[i][1] || '').trim() || String(dData[i][2] || '').trim();
      if (!key) continue;
      decisions[key] = {
        applicationId: String(dData[i][1] || ''), decision: String(dData[i][3] || ''),
        combinedScore: Number(dData[i][4]) || 0, decidedBy: String(dData[i][5] || ''),
        timestamp: dData[i][0]
      };
    }
  }
  return { success: true, interviews: interviews, decisions: decisions };
}

function submitDecisionData(params) {
  if (params.password !== getAdminPassword()) return { success: false, message: '權限不足' };
  const appId = params.applicationId || '';
  const candidateName = sanitizeInput(params.candidateName);
  const decision = params.decision;
  const combinedScore = Number(params.combinedScore) || 0;
  const decidedBy = sanitizeInput(params.decidedBy || '管理員');

  const validDecisions = ['錄取', '備取', '不錄取'];
  if (!validDecisions.includes(decision)) return { success: false, message: '無效的決策選項' };
  if (!candidateName) return { success: false, message: '缺少候選人資訊' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dSheet = ss.getSheetByName('Decisions') || ss.insertSheet('Decisions');
  const dData = dSheet.getDataRange().getValues();
  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");

  const dIdx = findCandidate(dData, appId, candidateName, 1, 2);
  if (dIdx > 0) {
    dSheet.getRange(dIdx + 1, 1, 1, 6).setValues([[ts, appId, candidateName, decision, combinedScore, decidedBy]]);
    logAudit('UPDATE_DECISION', '候選人：' + candidateName + '，決策：' + decision + '，綜合分：' + combinedScore);
  } else {
    dSheet.appendRow([ts, appId, candidateName, decision, combinedScore, decidedBy]);
    logAudit('SUBMIT_DECISION', '候選人：' + candidateName + '，決策：' + decision + '，綜合分：' + combinedScore);
  }
  return { success: true };
}

// ==================== Phase 2：面試排程管理 ====================
// Schedules: A:Timestamp, B:ScheduleId, C:ApplicationId, D:CandidateName,
//            E:InterviewDate, F:InterviewTime, G:Duration, H:Location, I:InterviewerName, J:Status

function getSchedulesData(password) {
  if (password !== getAdminPassword() && password !== getInterviewerPassword()) return { success: false, message: '密碼錯誤' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Schedules');
  if (!sheet) return { success: true, data: [] };
  const data = sheet.getDataRange().getValues();
  return {
    success: true,
    data: data.slice(1).filter(r => r[0]).map(r => ({
      timestamp: r[0], scheduleId: String(r[1] || ''), applicationId: String(r[2] || ''),
      candidateName: String(r[3] || ''), interviewDate: String(r[4] || ''),
      interviewTime: String(r[5] || ''), duration: Number(r[6]) || 30,
      location: String(r[7] || ''), interviewerName: String(r[8] || ''),
      status: String(r[9] || '待面試')
    }))
  };
}

function getInterviewerScheduleData(password, interviewerName) {
  if (password !== getAdminPassword() && password !== getInterviewerPassword()) return { success: false, message: '密碼錯誤' };
  if (!interviewerName) return { success: false, message: '請提供面試官姓名' };
  const result = getSchedulesData(password);
  if (!result.success) return result;
  const searchName = sanitizeInput(interviewerName);
  result.data = result.data.filter(s => s.interviewerName === searchName);
  return result;
}

function createScheduleData(params) {
  if (params.password !== getAdminPassword() && params.password !== getInterviewerPassword()) return { success: false, message: '權限不足' };
  const candidateName = sanitizeInput(params.candidateName);
  const appId = params.applicationId || '';
  const interviewDate = sanitizeInput(params.interviewDate);
  const interviewTime = sanitizeInput(params.interviewTime);
  const duration = Number(params.duration) || 30;
  const location = sanitizeInput(params.location);
  const interviewerName = sanitizeInput(params.interviewerName);

  if (!candidateName) return { success: false, message: '請指定候選人' };
  if (!interviewDate) return { success: false, message: '請指定面試日期' };
  if (!interviewTime) return { success: false, message: '請指定面試時間' };
  if (!location) return { success: false, message: '請指定面試地點' };
  if (!interviewerName) return { success: false, message: '請指定面試官' };
  if (![30, 45, 60, 90].includes(duration)) return { success: false, message: '時長須為 30/45/60/90 分鐘' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Schedules') || ss.insertSheet('Schedules');
  const data = sheet.getDataRange().getValues();

  // 檢查同時段同地點衝突
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4] || '') === interviewDate && String(data[i][5] || '') === interviewTime &&
        String(data[i][7] || '') === location && String(data[i][9] || '') !== '已取消') {
      return { success: false, message: '該時段與地點已有排程（' + String(data[i][3] || '') + '）' };
    }
  }

  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  const scheduleId = generateScheduleId();
  sheet.appendRow([ts, scheduleId, appId, candidateName, interviewDate, interviewTime, duration, location, interviewerName, '待面試']);
  logAudit('CREATE_SCHEDULE', '候選人：' + candidateName + '，日期：' + interviewDate + ' ' + interviewTime + '，地點：' + location + '，面試官：' + interviewerName);
  return { success: true, scheduleId: scheduleId };
}

function updateScheduleData(params) {
  if (params.password !== getAdminPassword() && params.password !== getInterviewerPassword()) return { success: false, message: '權限不足' };
  const scheduleId = params.scheduleId;
  if (!scheduleId) return { success: false, message: '缺少排程 ID' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Schedules');
  if (!sheet) return { success: false, message: '排程表不存在' };
  const data = sheet.getDataRange().getValues();

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() === scheduleId) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return { success: false, message: '找不到該排程' };

  const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
  const row = data[rowIdx];
  const updated = [
    ts,
    scheduleId,
    params.applicationId !== undefined ? params.applicationId : row[2],
    params.candidateName ? sanitizeInput(params.candidateName) : row[3],
    params.interviewDate ? sanitizeInput(params.interviewDate) : row[4],
    params.interviewTime ? sanitizeInput(params.interviewTime) : row[5],
    params.duration !== undefined ? Number(params.duration) : row[6],
    params.location ? sanitizeInput(params.location) : row[7],
    params.interviewerName ? sanitizeInput(params.interviewerName) : row[8],
    params.status ? sanitizeInput(params.status) : row[9]
  ];
  sheet.getRange(rowIdx + 1, 1, 1, 10).setValues([updated]);
  logAudit('UPDATE_SCHEDULE', '排程ID：' + scheduleId + '，更新欄位：' + Object.keys(params).filter(k => !['password','action','scheduleId'].includes(k)).join(','));
  return { success: true };
}

function deleteScheduleData(params) {
  if (params.password !== getAdminPassword() && params.password !== getInterviewerPassword()) return { success: false, message: '權限不足' };
  const scheduleId = params.scheduleId;
  if (!scheduleId) return { success: false, message: '缺少排程 ID' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Schedules');
  if (!sheet) return { success: false, message: '排程表不存在' };
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() === scheduleId) {
      // 標記為已取消（安全刪除）
      const ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm");
      sheet.getRange(i + 1, 1).setValue(ts);
      sheet.getRange(i + 1, 10).setValue('已取消');
      logAudit('CANCEL_SCHEDULE', '排程ID：' + scheduleId + '，候選人：' + String(data[i][3] || ''));
      return { success: true };
    }
  }
  return { success: false, message: '找不到該排程' };
}

// ==================== Phase 2：批次通知發送 ====================
// Notifications: A:Timestamp, B:NotificationId, C:ApplicationId, D:CandidateName,
//                E:Email, F:TemplateType, G:Subject, H:Body, I:Status, J:ErrorMessage

function getEmailTemplates() {
  return {
    '面試通知': {
      subject: '【炎洲集團】面試通知 — {{position}}',
      body: '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;padding:30px;">' +
        '<h2 style="color:#991b1b;border-bottom:3px solid #ef4444;padding-bottom:10px;">炎洲集團 面試通知</h2>' +
        '<p>{{name}} 您好，</p>' +
        '<p>感謝您應徵本集團「<strong>{{position}}</strong>」職務（{{dept}}）。</p>' +
        '<p>我們誠摯邀請您於以下時間前來面試：</p>' +
        '<table style="border-collapse:collapse;width:100%;margin:15px 0;">' +
        '<tr><td style="padding:8px 15px;background:#fdf2f2;font-weight:bold;width:80px;">日期</td><td style="padding:8px 15px;">{{date}}</td></tr>' +
        '<tr><td style="padding:8px 15px;background:#fdf2f2;font-weight:bold;">時間</td><td style="padding:8px 15px;">{{time}}</td></tr>' +
        '<tr><td style="padding:8px 15px;background:#fdf2f2;font-weight:bold;">地點</td><td style="padding:8px 15px;">{{location}}</td></tr>' +
        '</table>' +
        '<p>請攜帶身分證件及相關證照，提前 10 分鐘抵達。</p>' +
        '<p style="margin-top:30px;color:#64748b;font-size:13px;">炎洲集團 人力資源部</p></div>'
    },
    '錄取通知': {
      subject: '【炎洲集團】錄取通知 — {{position}}',
      body: '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;padding:30px;">' +
        '<h2 style="color:#22c55e;border-bottom:3px solid #22c55e;padding-bottom:10px;">炎洲集團 錄取通知</h2>' +
        '<p>{{name}} 您好，</p>' +
        '<p>恭喜您通過「<strong>{{position}}</strong>」（{{dept}}）職務面試！</p>' +
        '<p>我們正式通知您 <strong style="color:#22c55e;font-size:18px;">已被錄取</strong>。</p>' +
        '<p>您的綜合評分為：<strong>{{score}} 分</strong></p>' +
        '<p>請於收到此通知後 <strong>3 個工作日內</strong> 回覆是否接受錄取。</p>' +
        '<p>如有任何疑問，歡迎來電洽詢。</p>' +
        '<p style="margin-top:30px;color:#64748b;font-size:13px;">炎洲集團 人力資源部</p></div>'
    },
    '備取通知': {
      subject: '【炎洲集團】備取通知 — {{position}}',
      body: '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;padding:30px;">' +
        '<h2 style="color:#f59e0b;border-bottom:3px solid #f59e0b;padding-bottom:10px;">炎洲集團 備取通知</h2>' +
        '<p>{{name}} 您好，</p>' +
        '<p>感謝您參加「<strong>{{position}}</strong>」（{{dept}}）面試。</p>' +
        '<p>經過審慎評估，目前您的申請列為 <strong style="color:#f59e0b;">備取</strong>。</p>' +
        '<p>如正取人員未能報到，我們將優先通知您遞補。</p>' +
        '<p>再次感謝您的參與，祝一切順利。</p>' +
        '<p style="margin-top:30px;color:#64748b;font-size:13px;">炎洲集團 人力資源部</p></div>'
    },
    '不錄取通知': {
      subject: '【炎洲集團】面試結果通知 — {{position}}',
      body: '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;padding:30px;">' +
        '<h2 style="color:#64748b;border-bottom:3px solid #e2e8f0;padding-bottom:10px;">炎洲集團 面試結果通知</h2>' +
        '<p>{{name}} 您好，</p>' +
        '<p>感謝您應徵「<strong>{{position}}</strong>」（{{dept}}）並參加面試。</p>' +
        '<p>經過審慎評估，很遺憾此次未能錄取。</p>' +
        '<p>這並不代表您的能力不足，僅是此次職缺的條件較為特殊。</p>' +
        '<p>未來如有適合的職缺，我們將主動與您聯繫。</p>' +
        '<p>祝您求職順利！</p>' +
        '<p style="margin-top:30px;color:#64748b;font-size:13px;">炎洲集團 人力資源部</p></div>'
    }
  };
}

function substituteVars(text, vars) {
  return text.replace(/\{\{(\w+)\}\}/g, function(match, key) {
    return vars[key] !== undefined ? vars[key] : match;
  });
}

function getCandidateInfoForTemplate(appId, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var info = { name: name || '', position: '', dept: '', score: '', date: '', time: '', location: '', company: '炎洲集團' };

  // 從 Resumes 取得職務、部門
  const rSheet = ss.getSheetByName('Resumes');
  if (rSheet) {
    const rData = rSheet.getDataRange().getValues();
    const rIdx = findCandidate(rData, appId, name, 16, 1);
    if (rIdx > 0) {
      info.position = String(rData[rIdx][11] || '');
      info.dept = String(rData[rIdx][15] || '');
    }
  }

  // 從 Results 取得分數
  const resSheet = ss.getSheetByName('Results');
  if (resSheet) {
    const resData = resSheet.getDataRange().getValues();
    const resIdx = findCandidate(resData, appId, name, 6, 0);
    if (resIdx > 0) info.score = String(resData[resIdx][3] || '');
  }

  // 從 Schedules 取得排程 (最新一筆)
  const sSheet = ss.getSheetByName('Schedules');
  if (sSheet) {
    const sData = sSheet.getDataRange().getValues();
    for (let i = sData.length - 1; i >= 1; i--) {
      const match = (appId && String(sData[i][2] || '').trim() === appId) ||
                     (!appId && String(sData[i][3] || '').trim() === name);
      if (match && String(sData[i][9] || '') !== '已取消') {
        info.date = String(sData[i][4] || '');
        info.time = String(sData[i][5] || '');
        info.location = String(sData[i][7] || '');
        break;
      }
    }
  }

  return info;
}

function sendNotificationsData(params) {
  if (params.password !== getAdminPassword()) return { success: false, message: '權限不足' };

  const recipients = params.recipients;
  if (!recipients || !Array.isArray(recipients) || recipients.length === 0) return { success: false, message: '無收件人' };
  if (recipients.length > 20) return { success: false, message: '單次最多 20 位收件人' };

  const templateType = sanitizeInput(params.templateType);
  const templates = getEmailTemplates();
  const template = templates[templateType];
  if (!template) return { success: false, message: '無效的通知模板類型' };

  // 檢查剩餘配額
  var remaining;
  try { remaining = MailApp.getRemainingDailyQuota(); } catch (e) { remaining = 100; }
  if (remaining < recipients.length) return { success: false, message: '今日郵件配額不足（剩餘 ' + remaining + ' 封）' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nSheet = ss.getSheetByName('Notifications') || ss.insertSheet('Notifications');
  const results = [];
  var sent = 0, failed = 0;

  for (var r = 0; r < recipients.length; r++) {
    var rec = recipients[r];
    var candidateInfo = getCandidateInfoForTemplate(rec.applicationId || '', rec.candidateName || '');
    var subject = substituteVars(params.customSubject || template.subject, candidateInfo);
    var body = substituteVars(params.customBody || template.body, candidateInfo);
    var ts = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm:ss");
    var nId = generateNotificationId();
    var email = sanitizeInput(rec.email);

    if (!email || !isValidEmail(email)) {
      nSheet.appendRow([ts, nId, rec.applicationId || '', rec.candidateName || '', email, templateType, subject, '', '失敗', 'Email 格式無效']);
      results.push({ name: rec.candidateName, status: '失敗', error: 'Email 格式無效' });
      failed++;
      continue;
    }

    try {
      MailApp.sendEmail({ to: email, subject: subject, htmlBody: body, name: '炎洲集團人力資源部' });
      nSheet.appendRow([ts, nId, rec.applicationId || '', rec.candidateName || '', email, templateType, subject, body.substring(0, 5000), '已發送', '']);
      results.push({ name: rec.candidateName, status: '已發送' });
      sent++;
    } catch (e) {
      nSheet.appendRow([ts, nId, rec.applicationId || '', rec.candidateName || '', email, templateType, subject, '', '失敗', e.toString()]);
      results.push({ name: rec.candidateName, status: '失敗', error: e.toString() });
      failed++;
    }
  }

  logAudit('SEND_NOTIFICATIONS', templateType + '：成功 ' + sent + ' 封，失敗 ' + failed + ' 封');
  return { success: true, results: results, sent: sent, failed: failed };
}

function getNotificationLogData(password) {
  if (password !== getAdminPassword()) return { success: false, message: '密碼錯誤' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Notifications');
  if (!sheet) return { success: true, data: [] };
  const data = sheet.getDataRange().getValues();
  return {
    success: true,
    data: data.slice(1).filter(r => r[0]).map(r => ({
      timestamp: r[0], notificationId: String(r[1] || ''), applicationId: String(r[2] || ''),
      candidateName: String(r[3] || ''), email: String(r[4] || ''), templateType: String(r[5] || ''),
      subject: String(r[6] || ''), status: String(r[8] || ''), errorMessage: String(r[9] || '')
    }))
  };
}

function getEmailQuotaData(password) {
  if (password !== getAdminPassword()) return { success: false, message: '密碼錯誤' };
  try {
    var remaining = MailApp.getRemainingDailyQuota();
    return { success: true, remaining: remaining };
  } catch (e) {
    return { success: true, remaining: -1 };
  }
}

// ==================== Phase 2：PDF 歸檔 ====================

function buildCandidateReportHtml(candidateData) {
  var d = candidateData;
  var r = d.resumeArray || [];
  var examScore = d.score || 0;

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:sans-serif;padding:40px;max-width:800px;margin:0 auto;color:#1e293b;font-size:14px;}' +
    'h1{color:#991b1b;font-size:28px;text-align:center;border-bottom:4px solid #991b1b;padding-bottom:15px;margin-bottom:5px;}' +
    '.subtitle{text-align:center;color:#64748b;font-size:13px;margin-bottom:30px;}' +
    '.section{margin:25px 0;}' +
    '.section-title{font-size:16px;font-weight:bold;color:#991b1b;border-left:5px solid #ef4444;padding-left:10px;margin-bottom:15px;}' +
    'table{width:100%;border-collapse:collapse;margin:10px 0;}' +
    'td{padding:8px 12px;border:1px solid #e2e8f0;font-size:13px;}' +
    'td.label{background:#fdf2f2;font-weight:bold;width:100px;color:#991b1b;}' +
    '.score-box{text-align:center;padding:20px;margin:10px 0;}' +
    '.score-big{font-size:48px;font-weight:900;color:#991b1b;}' +
    '.bar-wrap{margin:8px 0;}' +
    '.bar-label{display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px;}' +
    '.bar{height:14px;background:#e2e8f0;border-radius:7px;overflow:hidden;}' +
    '.bar-fill{height:100%;border-radius:7px;}' +
    '.decision-badge{display:inline-block;padding:8px 25px;border-radius:10px;font-size:18px;font-weight:900;margin-top:10px;}' +
    '.footer{margin-top:40px;text-align:center;color:#94a3b8;font-size:11px;border-top:2px solid #e2e8f0;padding-top:15px;}' +
    '</style></head><body>';

  // Header
  html += '<h1>炎洲集團 候選人評估報告</h1>';
  html += '<div class="subtitle">Generated: ' + Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm") + '</div>';

  // 基本資料
  html += '<div class="section"><div class="section-title">個人資料</div><table>';
  var fields = [
    ['姓名', r[1] || ''], ['性別', r[2] || ''], ['生日', r[3] || ''], ['電話', r[4] || ''],
    ['Email', r[5] || ''], ['地址', r[6] || ''], ['學歷', r[7] || ''], ['學校', r[8] || ''],
    ['科系', r[9] || ''], ['年資', r[10] || ''], ['應徵職務', r[11] || ''], ['需求單位', r[15] || '']
  ];
  for (var i = 0; i < fields.length; i += 2) {
    html += '<tr><td class="label">' + fields[i][0] + '</td><td>' + fields[i][1] + '</td>';
    if (i + 1 < fields.length) {
      html += '<td class="label">' + fields[i+1][0] + '</td><td>' + fields[i+1][1] + '</td>';
    }
    html += '</tr>';
  }
  html += '</table>';

  // 經歷、技能、自傳
  if (r[12]) html += '<table><tr><td class="label">工作經歷</td><td>' + String(r[12]).replace(/\n/g, '<br>') + '</td></tr></table>';
  if (r[13]) html += '<table><tr><td class="label">專業技能</td><td>' + String(r[13]).replace(/\n/g, '<br>') + '</td></tr></table>';
  if (r[14]) html += '<table><tr><td class="label">自傳</td><td>' + String(r[14]).replace(/\n/g, '<br>') + '</td></tr></table>';
  html += '</div>';

  // 測驗成績
  html += '<div class="section"><div class="section-title">測驗成績</div>';
  html += '<div class="score-box"><div class="score-big">' + examScore + '<span style="font-size:20px;color:#64748b;"> / 100</span></div></div></div>';

  // 面試評分
  if (d.interview) {
    var iv = d.interview;
    var scores = iv.scores || {};
    var dimLabels = { professional: '專業能力', communication: '溝通表達', teamwork: '團隊協作', potential: '學習潛力', stability: '工作穩定度' };
    var barColors = ['#ef4444', '#f97316', '#8b5cf6', '#22c55e', '#06b6d4'];
    var ci = 0;

    html += '<div class="section"><div class="section-title">面試評分（面試官：' + (iv.interviewerName || '') + '）</div>';
    for (var key in dimLabels) {
      var val = scores[key] || 0;
      html += '<div class="bar-wrap"><div class="bar-label"><span>' + dimLabels[key] + '</span><span>' + val + ' / 5</span></div><div class="bar"><div class="bar-fill" style="width:' + (val * 20) + '%;background:' + barColors[ci] + ';"></div></div></div>';
      ci++;
    }
    html += '<table><tr><td class="label">面試均分</td><td>' + iv.totalAvg + ' / 5.0</td><td class="label">錄用建議</td><td>' + (iv.recommendation || '') + '</td></tr></table>';
    if (iv.notes) html += '<table><tr><td class="label">面試備註</td><td>' + String(iv.notes).replace(/\n/g, '<br>') + '</td></tr></table>';

    var combined = Math.round(examScore * 0.4 + iv.totalAvg * 20 * 0.6);
    html += '<div class="score-box" style="background:#fdf2f2;border-radius:15px;"><div style="color:#64748b;font-size:13px;margin-bottom:5px;">綜合評分</div><div class="score-big">' + combined + '</div><div style="color:#94a3b8;font-size:12px;">測驗 ' + examScore + ' x 40% + 面試 ' + iv.totalAvg + ' x 20 x 60%</div></div>';
    html += '</div>';
  }

  // 決策
  if (d.decision) {
    var dc = d.decision;
    var dcColor = dc.decision === '錄取' ? '#22c55e' : dc.decision === '備取' ? '#f59e0b' : '#ef4444';
    html += '<div class="section"><div class="section-title">錄取決策</div>';
    html += '<div style="text-align:center;"><div class="decision-badge" style="background:' + dcColor + '20;color:' + dcColor + ';">' + dc.decision + '</div>';
    html += '<div style="margin-top:10px;color:#64748b;font-size:13px;">由 ' + (dc.decidedBy || '') + ' 於 ' + (dc.timestamp || '') + ' 決定</div></div></div>';
  }

  // 簽名
  if (d.signature) {
    html += '<div class="section"><div class="section-title">數位簽名</div>';
    html += '<div style="text-align:center;padding:20px;"><img src="' + d.signature + '" style="max-width:250px;border-bottom:2px solid #333;"></div></div>';
  }

  html += '<div class="footer">炎洲集團人才招募 E 化系統 - 候選人評估報告 | 機密文件</div>';
  html += '</body></html>';
  return html;
}

function generatePdfData(params) {
  if (params.password !== getAdminPassword()) return { success: false, message: '權限不足' };
  var candidateName = sanitizeInput(params.candidateName);
  var appId = params.applicationId || '';
  if (!candidateName && !appId) return { success: false, message: '請指定候選人' };

  try {
    // 取得完整資料
    var infoRes = getFullCandidateInfo(candidateName || 'lookup', params.password, appId);
    if (!infoRes.success) return { success: false, message: '無法取得候選人資料' };
    var candidateData = infoRes.data;
    if (!candidateData.resumeArray || candidateData.resumeArray.length === 0) return { success: false, message: '找不到候選人履歷資料' };

    var actualName = String(candidateData.resumeArray[1] || candidateName);

    // 生成 HTML → PDF
    var htmlContent = buildCandidateReportHtml(candidateData);
    var pdfBlob = Utilities.newBlob(htmlContent, 'text/html', 'report.html').getAs('application/pdf');
    var fileName = actualName + '_' + (appId || 'NA') + '_評估報告.pdf';
    pdfBlob.setName(fileName);

    // 存到 Google Drive 指定資料夾
    var folders = DriveApp.getFoldersByName('招募系統PDF歸檔');
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('招募系統PDF歸檔');
    var file = folder.createFile(pdfBlob);

    logAudit('GENERATE_PDF', '候選人：' + actualName + '，檔案：' + fileName + '，連結：' + file.getUrl());
    return { success: true, fileUrl: file.getUrl(), fileName: fileName };
  } catch (e) {
    logAudit('ERROR', 'generatePdf: ' + e.toString());
    return { success: false, message: 'PDF 生成失敗：' + e.toString() };
  }
}

function generateBulkPdfData(params) {
  if (params.password !== getAdminPassword()) return { success: false, message: '權限不足' };
  var filter = params.filter || 'all'; // all, 錄取, 備取, 不錄取

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dSheet = ss.getSheetByName('Decisions');
  if (!dSheet) return { success: false, message: '尚無決策資料' };
  var dData = dSheet.getDataRange().getValues();

  var targets = [];
  for (var i = 1; i < dData.length; i++) {
    var decision = String(dData[i][3] || '');
    if (filter === 'all' || decision === filter) {
      targets.push({ applicationId: String(dData[i][1] || ''), candidateName: String(dData[i][2] || ''), decision: decision });
    }
  }

  if (targets.length === 0) return { success: false, message: '沒有符合條件的候選人' };

  // 限制每批 10 人（GAS 6 分鐘限制）
  var batchSize = Math.min(targets.length, 10);
  var processed = 0, errors = [];
  var folders = DriveApp.getFoldersByName('招募系統PDF歸檔');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('招募系統PDF歸檔');

  for (var j = 0; j < batchSize; j++) {
    try {
      var t = targets[j];
      var infoRes = getFullCandidateInfo(t.candidateName || 'lookup', params.password, t.applicationId);
      if (!infoRes.success) { errors.push(t.candidateName + '：資料取得失敗'); continue; }

      var htmlContent = buildCandidateReportHtml(infoRes.data);
      var pdfBlob = Utilities.newBlob(htmlContent, 'text/html', 'report.html').getAs('application/pdf');
      var fileName = t.candidateName + '_' + (t.applicationId || 'NA') + '_評估報告.pdf';
      pdfBlob.setName(fileName);
      folder.createFile(pdfBlob);
      processed++;
    } catch (e) {
      errors.push(t.candidateName + '：' + e.toString());
    }
  }

  logAudit('GENERATE_BULK_PDF', '篩選：' + filter + '，成功 ' + processed + '/' + batchSize + ' 份');
  return {
    success: true,
    processed: processed,
    total: targets.length,
    batchSize: batchSize,
    errors: errors,
    folderUrl: folder.getUrl(),
    hasMore: targets.length > batchSize
  };
}

function getArchiveLogData(password) {
  if (password !== getAdminPassword()) return { success: false, message: '密碼錯誤' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('AuditLog');
  if (!sheet) return { success: true, data: [] };
  var data = sheet.getDataRange().getValues();
  var archives = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '') === 'GENERATE_PDF' || String(data[i][1] || '') === 'GENERATE_BULK_PDF') {
      archives.push({ timestamp: data[i][0], action: String(data[i][1] || ''), detail: String(data[i][2] || '') });
    }
  }
  return { success: true, data: archives.reverse() };
}
