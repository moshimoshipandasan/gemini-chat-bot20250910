/**
 * ミニマル Gemini チャットボット
 * シンプルで実用的なGAS実装
 */

// ===== 設定 =====
const MINIMAL_MINIMAL_CONFIG = {
  API_MODEL: 'gemini-pro',
  API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',
  CACHE_DURATION: 21600, // 6時間（秒）
  MAX_HISTORY: 10, // 保持する会話履歴の最大数
  
  SHEETS: {
    PROMPT: 'プロンプト',
    LOG: 'ログ'
  }
};

// ===== メイン関数 =====

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('minimal-index')
    .setTitle('Gemini チャットボット')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * メッセージ処理
 * @param {string} userId - ユーザーID
 * @param {string} message - ユーザーメッセージ
 * @returns {string} AIの応答
 */
function processMessage(userId, message) {
  try {
    // 入力検証
    if (!userId || !message) {
      throw new Error('ユーザーIDまたはメッセージが無効です');
    }
    
    // 会話履歴の取得・更新
    const history = getConversationHistory(userId);
    history.push({ role: 'user', content: message });
    
    // 履歴を制限
    if (history.length > MINIMAL_CONFIG.MAX_HISTORY * 2) {
      history.splice(0, history.length - MINIMAL_CONFIG.MAX_HISTORY * 2);
    }
    
    // Gemini APIを呼び出し
    const response = callGeminiAPI(history);
    
    // 応答を履歴に追加
    history.push({ role: 'model', content: response });
    
    // 履歴を保存
    saveConversationHistory(userId, history);
    
    // ログに記録
    logToSheet(userId, 'user', message);
    logToSheet(userId, 'model', response);
    
    return response;
    
  } catch (error) {
    console.error('processMessage error:', error);
    logToSheet(userId, 'error', error.toString());
    throw error;
  }
}

/**
 * 会話履歴をクリア
 * @param {string} userId - ユーザーID
 */
function clearHistory(userId) {
  const cache = CacheService.getScriptCache();
  cache.remove('chat_' + userId);
}

// ===== Gemini API関連 =====

/**
 * Gemini APIを呼び出す
 * @param {Array} history - 会話履歴
 * @returns {string} AIの応答
 */
function callGeminiAPI(history) {
  const apiKey = getApiKey();
  const systemPrompt = getSystemPrompt();
  
  // APIリクエストの構築
  const messages = [
    { role: 'user', parts: [{ text: systemPrompt }] },
    ...history.map(msg => ({
      role: msg.role,
      parts: [{ text: msg.content }]
    }))
  ];
  
  const payload = {
    contents: messages,
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 1024
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const url = `${MINIMAL_CONFIG.API_BASE_URL}/${MINIMAL_CONFIG.API_MODEL}:generateContent?key=${apiKey}`;
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini API エラー: ' + response.getContentText());
  }
  
  const json = JSON.parse(response.getContentText());
  
  if (!json.candidates || !json.candidates[0]?.content?.parts?.[0]?.text) {
    throw new Error('Gemini APIから有効な応答がありません');
  }
  
  return json.candidates[0].content.parts[0].text;
}

/**
 * APIキーを取得
 */
function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('APIキーが設定されていません。スクリプトプロパティにGEMINI_API_KEYを設定してください。');
  }
  return apiKey;
}

// ===== スプレッドシート関連 =====

/**
 * システムプロンプトを取得（A1セルから）
 */
function getSystemPrompt() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MINIMAL_CONFIG.SHEETS.PROMPT);
  if (!sheet) {
    // プロンプトシートがない場合は作成
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(MINIMAL_CONFIG.SHEETS.PROMPT);
    newSheet.getRange('A1').setValue('あなたは親切なAIアシスタントです。ユーザーの質問に丁寧に答えてください。');
    return newSheet.getRange('A1').getValue();
  }
  
  const prompt = sheet.getRange('A1').getValue();
  return prompt || 'あなたは親切なAIアシスタントです。';
}

/**
 * ログをスプレッドシートに記録
 */
function logToSheet(userId, role, message) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MINIMAL_CONFIG.SHEETS.LOG);
    
    // ログシートがない場合は作成
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(MINIMAL_CONFIG.SHEETS.LOG);
      // ヘッダーを設定
      sheet.getRange('A1:E1').setValues([['タイムスタンプ', 'ユーザーID', '役割', 'メッセージ', 'トークン数']]);
      sheet.setFrozenRows(1);
    }
    
    // ログを追加
    const timestamp = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
    const tokenCount = Math.ceil(message.length / 4); // 簡易的なトークン数計算
    
    sheet.appendRow([timestamp, userId, role, message, tokenCount]);
    
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
}

// ===== キャッシュ管理 =====

/**
 * 会話履歴を取得
 */
function getConversationHistory(userId) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('chat_' + userId);
  
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      console.error('キャッシュ解析エラー:', e);
    }
  }
  
  return [];
}

/**
 * 会話履歴を保存
 */
function saveConversationHistory(userId, history) {
  const cache = CacheService.getScriptCache();
  cache.put('chat_' + userId, JSON.stringify(history), MINIMAL_CONFIG.CACHE_DURATION);
}

// ===== 初期設定用ヘルパー関数 =====

/**
 * 初期設定（手動実行用）
 * APIキーを設定してから実行してください
 */
function setup() {
  // プロンプトシートの作成
  const promptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MINIMAL_CONFIG.SHEETS.PROMPT);
  if (!promptSheet) {
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(MINIMAL_CONFIG.SHEETS.PROMPT);
    newSheet.getRange('A1').setValue('あなたは親切なAIアシスタントです。ユーザーの質問に丁寧に答えてください。');
    console.log('プロンプトシートを作成しました');
  }
  
  // ログシートの作成
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MINIMAL_CONFIG.SHEETS.LOG);
  if (!logSheet) {
    const newLogSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(MINIMAL_CONFIG.SHEETS.LOG);
    newLogSheet.getRange('A1:E1').setValues([['タイムスタンプ', 'ユーザーID', '役割', 'メッセージ', 'トークン数']]);
    newLogSheet.setFrozenRows(1);
    console.log('ログシートを作成しました');
  }
  
  // APIキーの確認
  try {
    getApiKey();
    console.log('APIキーが設定されています');
  } catch (e) {
    console.error('APIキーが設定されていません。スクリプトプロパティにGEMINI_API_KEYを設定してください。');
  }
}

/**
 * テスト用関数
 */
function test() {
  const result = processMessage('test_user', 'こんにちは');
  console.log('テスト結果:', result);
}