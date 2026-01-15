function doGet(e) {
  const page = e.parameter.page;
  
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const action = e.parameter.action;
  
  if (action === 'getQuestions') {
    return getQuestions();
  } else if (action === 'submitAnswers') {
    return submitAnswers(e.parameter.answers, e.parameter.userName);
  } else if (action === 'getResults') {
    return getResults(e.parameter.password);
  }
}

function getQuestions() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const questionsSheet = sheet.getSheetByName('Questions');
    const data = questionsSheet.getDataRange().getValues();
    
    const questions = [];
    for (let i = 1; i < data.length; i++) {
      // 從 D/E/F/G 欄位讀取選項，過濾空值
      const options = [data[i][3], data[i][4], data[i][5], data[i][6]]
        .filter(opt => opt !== null && opt !== undefined && opt !== '');
      
      questions.push({
        id: data[i][0],
        question: data[i][1],
        options: options
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: questions
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitAnswers(answersJson, userName) {
  try {
    const answers = JSON.parse(answersJson);
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const answersSheet = sheet.getSheetByName('Answers');
    const answerData = answersSheet.getDataRange().getValues();
    
    let totalScore = 0;
    const results = [];
    
    // 建立標準答案對照表
    const answerMap = {};
    for (let i = 1; i < answerData.length; i++) {
      // 確保題目ID是字串型別
      const id = String(answerData[i][0]);
      answerMap[id] = {
        correctAnswer: String(answerData[i][1]).trim(), // 轉成字串並去除空格
        score: Number(answerData[i][2]) // 確保是數字
      };
    }
    
    // 批改每一題
    for (const [questionKey, userAnswer] of Object.entries(answers)) {
      // 從 questionKey 提取題目ID (例如 "q1" -> "1")
      const questionId = questionKey.replace('q', '');
      const standard = answerMap[questionId];
      
      if (!standard) {
        Logger.log('找不到題目 ' + questionId + ' 的標準答案');
        Logger.log('可用的題目ID: ' + Object.keys(answerMap).join(', '));
        continue;
      }
      
      let earnedScore = 0;
      const trimmedUserAnswer = String(userAnswer).trim();
      const trimmedCorrectAnswer = standard.correctAnswer;
      
      // 除錯日誌
      Logger.log('題目 ' + questionId + ':');
      Logger.log('  使用者答案: "' + trimmedUserAnswer + '"');
      Logger.log('  正確答案: "' + trimmedCorrectAnswer + '"');
      Logger.log('  是否相同: ' + (trimmedUserAnswer === trimmedCorrectAnswer));
      
      // 比對答案（去除空格後比較）
      if (trimmedUserAnswer === trimmedCorrectAnswer) {
        earnedScore = standard.score;
        totalScore += earnedScore;
        Logger.log('  得分: ' + earnedScore);
      } else {
        Logger.log('  得分: 0');
      }
      
      results.push({
        questionId: questionId,
        userAnswer: userAnswer,
        correctAnswer: standard.correctAnswer,
        earned: earnedScore,
        total: standard.score,
        isCorrect: earnedScore > 0
      });
    }
    
    // 儲存結果
    const resultsSheet = sheet.getSheetByName('Results');
    const timestamp = new Date().toLocaleString('zh-TW');
    
    // D 欄只儲存答案，用 | 分隔
    const answersOnly = results.map(r => r.userAnswer).join(' | ');
    
    resultsSheet.appendRow([userName, totalScore, timestamp, answersOnly]);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      score: totalScore,
      details: results
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('submitAnswers 錯誤：' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getResults(password) {
  try {
    // 驗證密碼
    if (password !== 'admin') {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: '密碼錯誤'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const resultsSheet = sheet.getSheetByName('Results');
    const data = resultsSheet.getDataRange().getValues();
    
    const results = [];
    for (let i = 1; i < data.length; i++) {
      results.push({
        name: data[i][0],
        score: data[i][1],
        timestamp: data[i][2],
        answers: data[i][3]
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: results
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}