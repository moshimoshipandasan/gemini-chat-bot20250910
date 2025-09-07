/**
 * My Chatbot - Google Apps Script Web App
 * LINEスタイルのチャットインターフェースを提供するGemini AI搭載チャットボット
 */

/**
 * アプリケーション設定
 */
const CONFIG = {
  // Gemini API設定
  API: {
    MODEL: 'gemini-2.5-flash',
    BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',
    GENERATION_CONFIG: {
      temperature: 0.3,
      top_p: 0.9,
      top_k: 40,
      max_output_tokens: 8192
    }
  },
  
  // キャッシュ設定
  CACHE: {
    KEY: 'conversationHistory',
    DURATION: 21600 // 6時間（秒）
  },
  
  // 会話履歴設定
  CONVERSATION: {
    MAX_HISTORY_LENGTH: 10
  },
  
  // スプレッドシート設定
  SHEETS: {
    PROMPT: 'プロンプト',
    LOG: 'ログ',
    SETTINGS: '設定',
    SYSTEM_PROMPT_CELL: 'A1'
  },
  
  // ログシートの列設定
  LOG_COLUMNS: {
    HEADERS: ['タイムスタンプ', 'ユーザーID', '役割', 'メッセージ', 'トークン数'],
    WIDTHS: {
      TIMESTAMP: 180,
      USER_ID: 150,
      ROLE: 100,
      MESSAGE: 400,
      TOKEN_COUNT: 100
    }
  },
  
  // タイムゾーン
  TIMEZONE: 'Asia/Tokyo',
  
  // エラーメッセージ
  ERRORS: {
    NO_API_KEY: 'APIキーが設定されていません',
    API_CALL_FAILED: 'Gemini APIの呼び出しに失敗しました',
    NO_RESPONSE: 'Gemini APIからの応答がありません',
    INVALID_USER_ID: 'ユーザーIDが無効です',
    INVALID_MESSAGE: 'メッセージが無効です'
  }
};

/**
 * APIキーを取得
 * @returns {string} Gemini APIキー
 * @throws {Error} APIキーが設定されていない場合
 */
function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('apikey');
  if (!apiKey) {
    throw new Error(CONFIG.ERRORS.NO_API_KEY);
  }
  return apiKey;
}

/**
 * Gemini APIのURLを生成
 * @returns {string} 完全なAPI URL
 */
function getGeminiUrl() {
  return `${CONFIG.API.BASE_URL}/${CONFIG.API.MODEL}:generateContent?key=${getApiKey()}`;
}

/**
 * スプレッドシートにカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('チャットボット設定')
    .addItem('キャッシュをクリア', 'clearCache')
    .addItem('設定シートを初期化', 'reinitializeSettingsSheet')
    .addToUi();
}

/**
 * キャッシュをクリア
 * 会話履歴のキャッシュを削除する
 */
function clearCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CONFIG.CACHE.KEY);
    
    // 成功メッセージを表示
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'キャッシュをクリアしました。会話履歴がリセットされました。',
      'キャッシュクリア完了',
      3
    );
    
    console.log('キャッシュクリア実行: ' + new Date().toLocaleString('ja-JP'));
  } catch (error) {
    console.error('キャッシュクリアエラー:', error);
    
    // エラーメッセージを表示
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'キャッシュのクリアに失敗しました: ' + error.message,
      'エラー',
      5
    );
    
    throw error;
  }
}

/**
 * スプレッドシートのインスタンスを取得
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * システムプロンプトを取得
 * @returns {string} システムプロンプト
 */
function getSystemPrompt() {
  const sheet = getSpreadsheet().getSheetByName(CONFIG.SHEETS.PROMPT);
  if (!sheet) {
    throw new Error(`シート「${CONFIG.SHEETS.PROMPT}」が見つかりません`);
  }
  return sheet.getRange(CONFIG.SHEETS.SYSTEM_PROMPT_CELL).getValue();
}

/**
 * ログシートの取得または作成
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getLogSheet() {
  const ss = getSpreadsheet();
  let logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEETS.LOG);
    initializeLogSheet(logSheet);
  }
  
  return logSheet;
}

/**
 * 設定シートの取得または作成
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSettingsSheet() {
  const ss = getSpreadsheet();
  let settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
    initializeSettingsSheet(settingsSheet);
  }
  
  return settingsSheet;
}

/**
 * 設定シートを初期化
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 初期化するシート
 */
function initializeSettingsSheet(sheet) {
  // ヘッダーの設定
  const headers = [
    ['設定項目', '値', '説明'],
    ['テーマ', 'default', 'default(青空), dark, sunset, forest, ocean, lavender, midnight, sakura, custom'],
    ['文字サイズ', 'medium', 'small, medium, large, xlarge'],
    ['タイピングインジケーター', 'true', 'true/false'],
    ['通知音', 'false', 'true/false'],
    ['メッセージ削除可能', 'false', 'true/false'],
    ['最大履歴表示数', '50', '表示するメッセージの最大数'],
    ['自動スクロール', 'true', 'true/false'],
    ['送信ショートカット', 'Enter', 'Enter/Ctrl+Enter'],
    ['背景色', '#B3D9FF', 'カスタムテーマ用'],
    ['ヘッダー色', '#6FB7FF', 'カスタムテーマ用'],
    ['ユーザーメッセージ色', '#92E05D', 'カスタムテーマ用'],
    ['AIメッセージ色', '#FFFFFF', 'カスタムテーマ用'],
    ['フォントファミリー', 'Noto Sans JP', 'フォント名']
  ];
  
  sheet.getRange(1, 1, headers.length, 3).setValues(headers);
  sheet.setFrozenRows(1);
  
  // 列幅の設定
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 300);
  
  // ヘッダーの書式設定
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
}

/**
 * ログシートを初期化
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 初期化するシート
 */
function initializeLogSheet(sheet) {
  // ヘッダーの設定
  sheet.getRange('A1:E1').setValues([CONFIG.LOG_COLUMNS.HEADERS]);
  sheet.setFrozenRows(1);
  
  // 列幅の設定
  const widths = CONFIG.LOG_COLUMNS.WIDTHS;
  sheet.setColumnWidth(1, widths.TIMESTAMP);
  sheet.setColumnWidth(2, widths.USER_ID);
  sheet.setColumnWidth(3, widths.ROLE);
  sheet.setColumnWidth(4, widths.MESSAGE);
  sheet.setColumnWidth(5, widths.TOKEN_COUNT);
}

/**
 * チャットログを記録
 * @param {string} userId - ユーザーID
 * @param {string} role - 役割（'user' または 'model'）
 * @param {string} message - メッセージ内容
 */
function logChat(userId, role, message) {
  try {
    const logSheet = getLogSheet();
    const timestamp = new Date().toLocaleString('ja-JP', { timeZone: CONFIG.TIMEZONE });
    const tokenCount = estimateTokenCount(message);
    
    logSheet.appendRow([
      timestamp,
      userId,
      role,
      message,
      tokenCount
    ]);
  } catch (error) {
    console.error('ログの記録に失敗しました:', error);
  }
}

/**
 * トークン数を概算
 * @param {string} text - テキスト
 * @returns {number} 推定トークン数
 */
function estimateTokenCount(text) {
  if (!text) return 0;
  // 日本語は文字あたり約2トークン、英語は4文字あたり1トークンとして概算
  const japaneseChars = (text.match(/[\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf]/g) || []).length;
  const otherChars = text.length - japaneseChars;
  return Math.ceil(japaneseChars * 2 + otherChars / 4);
}

/**
 * キャッシュサービスのインスタンスを取得
 * @returns {GoogleAppsScript.Cache.Cache}
 */
function getCache() {
  return CacheService.getScriptCache();
}

/**
 * GETリクエストを処理
 * @param {Object} request - HTTPリクエストオブジェクト
 * @returns {GoogleAppsScript.HTML.HtmlOutput} HTML出力
 */
function doGet(request) {
  try {
    // 新しいセッションのためにキャッシュをクリア
    getCache().remove(CONFIG.CACHE.KEY);
    
    // 設定を取得
    const settings = getSettings();
    
    // HTMLテンプレートを作成
    const template = HtmlService.createTemplateFromFile('index');
    template.settings = JSON.stringify(settings); // 設定をテンプレートに渡す
    const output = template.evaluate();
    
    // セキュリティ設定
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    output.setTitle('My Chatbot');
    
    return output;
  } catch (error) {
    console.error('ページの読み込みに失敗しました:', error);
    return HtmlService.createHtmlOutput('エラーが発生しました。管理者にお問い合わせください。');
  }
}

/**
 * メッセージを処理
 * @param {string} userId - ユーザーID
 * @param {string} message - ユーザーのメッセージ
 * @returns {string} AIの応答
 * @throws {Error} 処理中にエラーが発生した場合
 */
function processMessage(userId, message) {
  // 入力検証
  if (!userId || typeof userId !== 'string') {
    throw new Error(CONFIG.ERRORS.INVALID_USER_ID);
  }
  if (!message || typeof message !== 'string') {
    throw new Error(CONFIG.ERRORS.INVALID_MESSAGE);
  }
  
  try {
    // 会話履歴を取得
    const conversationHistory = getConversationHistory();
    
    // ユーザーの会話履歴を初期化
    if (!conversationHistory[userId]) {
      conversationHistory[userId] = [];
    }
    
    // ユーザーメッセージを追加
    addMessageToHistory(conversationHistory[userId], 'user', message);
    logChat(userId, 'user', message);
    
    // 履歴を制限
    trimConversationHistory(conversationHistory[userId]);
    
    // Gemini APIを呼び出し
    const response = callGeminiAPI(userId, conversationHistory[userId]);
    
    // AIの応答を追加
    addMessageToHistory(conversationHistory[userId], 'model', response);
    logChat(userId, 'model', response);
    
    // 履歴を保存
    saveConversationHistory(conversationHistory);
    
    return response;
  } catch (error) {
    console.error('メッセージ処理エラー:', error);
    throw error;
  }
}

/**
 * 会話履歴を取得
 * @returns {Object} 会話履歴オブジェクト
 */
function getConversationHistory() {
  const cache = getCache();
  const cached = cache.get(CONFIG.CACHE.KEY);
  return cached ? JSON.parse(cached) : {};
}

/**
 * 会話履歴を保存
 * @param {Object} conversationHistory - 会話履歴オブジェクト
 */
function saveConversationHistory(conversationHistory) {
  const cache = getCache();
  cache.put(CONFIG.CACHE.KEY, JSON.stringify(conversationHistory), CONFIG.CACHE.DURATION);
}

/**
 * メッセージを履歴に追加
 * @param {Array} history - 会話履歴配列
 * @param {string} role - 役割（'user' または 'model'）
 * @param {string} text - メッセージテキスト
 */
function addMessageToHistory(history, role, text) {
  history.push({
    role: role,
    parts: [{ text: text }]
  });
}

/**
 * 会話履歴を最大長に制限
 * @param {Array} history - 会話履歴配列
 */
function trimConversationHistory(history) {
  if (history.length > CONFIG.CONVERSATION.MAX_HISTORY_LENGTH) {
    history.splice(0, history.length - CONFIG.CONVERSATION.MAX_HISTORY_LENGTH);
  }
}

/**
 * Gemini APIを呼び出す（リトライ機構付き）
 * @param {string} userId - ユーザーID
 * @param {Array} history - 会話履歴
 * @returns {string} AIの応答テキスト
 * @throws {Error} API呼び出しに失敗した場合
 */
function callGeminiAPI(userId, history) {
  const payload = buildApiPayload(history);
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  // リトライ設定
  const maxRetries = 3;
  const baseDelay = 1000; // 1秒
  let lastError = null;
  
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      // リトライの場合は指数バックオフで待機
      if (attempt > 0) {
        const delay = baseDelay * Math.pow(2, attempt - 1);
        console.log(`リトライ ${attempt}/${maxRetries - 1}: ${delay}ms 待機中...`);
        Utilities.sleep(delay);
      }
      
      const response = UrlFetchApp.fetch(getGeminiUrl(), options);
      const responseCode = response.getResponseCode();
      
      // 成功
      if (responseCode === 200) {
        const responseJson = JSON.parse(response.getContentText());
        return extractResponseText(responseJson);
      }
      
      // リトライ可能なエラー（429: Too Many Requests, 503: Service Unavailable）
      if (responseCode === 429 || responseCode === 503) {
        lastError = new Error(`API一時エラー (${responseCode}): リトライします`);
        console.warn(`リトライ可能なエラー: ${responseCode}`);
        continue;
      }
      
      // リトライ不可能なエラー
      console.error(`API エラー: ${responseCode}`);
      throw new Error(`${CONFIG.ERRORS.API_CALL_FAILED}: ${responseCode}`);
      
    } catch (error) {
      lastError = error;
      
      // ネットワークエラーの場合はリトライ
      if (error.toString().includes('UrlFetchApp') || 
          error.toString().includes('Network') ||
          error.toString().includes('Timeout')) {
        console.warn(`ネットワークエラー (試行 ${attempt + 1}/${maxRetries}):`, error);
        continue;
      }
      
      // その他のエラーは即座に投げる
      if (!error.message.includes('API一時エラー')) {
        console.error('Gemini API呼び出しエラー:', error);
        throw error;
      }
    }
  }
  
  // すべてのリトライが失敗した場合
  console.error('すべてのリトライが失敗しました:', lastError);
  throw new Error(`API呼び出しが${maxRetries}回失敗しました: ${lastError.message}`);
}

/**
 * API用のペイロードを構築
 * @param {Array} history - 会話履歴
 * @returns {Object} APIペイロード
 */
function buildApiPayload(history) {
  return {
    systemInstruction: {
      role: 'model',
      parts: [{ text: getSystemPrompt() }]
    },
    contents: history,
    generationConfig: CONFIG.API.GENERATION_CONFIG
  };
}

/**
 * APIレスポンスからテキストを抽出
 * @param {Object} responseJson - APIレスポンス
 * @returns {string} 応答テキスト
 * @throws {Error} 応答が無効な場合
 */
function extractResponseText(responseJson) {
  if (!responseJson?.candidates?.length || 
      !responseJson.candidates[0]?.content?.parts?.length) {
    throw new Error(CONFIG.ERRORS.NO_RESPONSE);
  }
  
  return responseJson.candidates[0].content.parts[0].text;
}


/**
 * 設定を取得
 * @returns {Object} 設定オブジェクト
 */
function getSettings() {
  try {
    const sheet = getSettingsSheet();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    // デフォルト設定から開始（すべての設定が確実に存在するようにする）
    const settings = Object.assign({}, getDefaultSettings());
    
    // シートから取得した設定で上書き（空でない値のみ）
    data.forEach(row => {
      if (row[0] && row[1] !== '') {
        settings[row[0]] = row[1];
      }
    });
    
    return settings;
  } catch (error) {
    console.error('設定の取得エラー:', error);
    return getDefaultSettings();
  }
}

/**
 * デフォルト設定を取得
 * @returns {Object} デフォルト設定
 */
function getDefaultSettings() {
  return {
    'テーマ': 'default',
    '文字サイズ': 'medium',
    'タイピングインジケーター': 'true',
    '通知音': 'false',
    'メッセージ削除可能': 'false',
    '最大履歴表示数': '50',
    '自動スクロール': 'true',
    '送信ショートカット': 'Enter',
    '背景色': '#B3D9FF',
    'ヘッダー色': '#6FB7FF',
    'ユーザーメッセージ色': '#92E05D',
    'AIメッセージ色': '#FFFFFF',
    'フォントファミリー': 'Noto Sans JP'
  };
}

/**
 * 設定を保存
 * @param {Object} settings - 保存する設定（部分的な更新も可能）
 * @returns {boolean} 成功/失敗
 */
function saveSettings(settings) {
  try {
    const sheet = getSettingsSheet();
    
    // 既存の設定を取得（デフォルト値も含む）
    const currentSettings = getSettings();
    
    // 新しい設定を既存の設定にマージ
    const mergedSettings = Object.assign({}, currentSettings, settings);
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    // マージされた設定値で更新
    data.forEach((row, index) => {
      const key = row[0];
      if (key && mergedSettings.hasOwnProperty(key)) {
        sheet.getRange(index + 2, 2).setValue(mergedSettings[key]);
      }
    });
    
    console.log('設定を保存しました:', mergedSettings);
    return true;
  } catch (error) {
    console.error('設定の保存エラー:', error);
    return false;
  }
}

/**
 * 会話履歴をクリア
 * @param {string} userId - ユーザーID（省略時は全ユーザー）
 */
function clearConversationHistory(userId = null) {
  const cache = getCache();
  
  if (userId) {
    const history = getConversationHistory();
    if (history[userId]) {
      delete history[userId];
      saveConversationHistory(history);
    }
  } else {
    cache.remove(CONFIG.CACHE.KEY);
  }
}

/**
 * 会話をエクスポート（複数形式対応）
 * @param {string} userId - ユーザーID
 * @param {string} format - エクスポート形式（'json', 'csv', 'text', 'drive'）
 * @returns {Object} エクスポート結果
 */
function exportConversation(userId, format = 'json') {
  try {
    const history = getConversationHistory();
    const userHistory = history[userId] || [];
    
    // ログからメッセージ詳細を取得
    const logSheet = getLogSheet();
    const logs = logSheet.getDataRange().getValues();
    const userLogs = logs.filter(row => row[1] === userId).slice(1); // ヘッダーを除く
    
    const exportData = {
      userId: userId,
      exportDate: new Date().toISOString(),
      messages: userHistory.map((msg, index) => ({
        role: msg.role,
        content: msg.parts[0].text,
        timestamp: userLogs[index] ? userLogs[index][0] : new Date().toISOString(),
        tokenCount: userLogs[index] ? userLogs[index][4] : estimateTokenCount(msg.parts[0].text)
      }))
    };
    
    switch (format.toLowerCase()) {
      case 'json':
        return exportAsJson(exportData);
      case 'csv':
        return exportAsCsv(exportData);
      case 'text':
        return exportAsText(exportData);
      case 'drive':
        return exportToDrive(exportData, userId);
      default:
        throw new Error(`サポートされていない形式: ${format}`);
    }
  } catch (error) {
    console.error('会話のエクスポートエラー:', error);
    throw error;
  }
}

/**
 * JSON形式でエクスポート
 */
function exportAsJson(data) {
  return {
    success: true,
    format: 'json',
    content: JSON.stringify(data, null, 2),
    mimeType: 'application/json',
    filename: `chat_export_${data.userId}_${Date.now()}.json`
  };
}

/**
 * CSV形式でエクスポート
 */
function exportAsCsv(data) {
  const headers = ['タイムスタンプ', '役割', 'メッセージ', 'トークン数'];
  const rows = [headers];
  
  data.messages.forEach(msg => {
    rows.push([
      msg.timestamp,
      msg.role === 'user' ? 'ユーザー' : 'AI',
      msg.content.replace(/"/g, '""'), // CSVエスケープ
      msg.tokenCount
    ]);
  });
  
  const csv = rows.map(row => 
    row.map(cell => `"${cell}"`).join(',')
  ).join('\n');
  
  return {
    success: true,
    format: 'csv',
    content: csv,
    mimeType: 'text/csv',
    filename: `chat_export_${data.userId}_${Date.now()}.csv`
  };
}

/**
 * テキスト形式でエクスポート
 */
function exportAsText(data) {
  let text = `チャット履歴エクスポート\n`;
  text += `ユーザーID: ${data.userId}\n`;
  text += `エクスポート日時: ${data.exportDate}\n`;
  text += `${'='.repeat(50)}\n\n`;
  
  data.messages.forEach(msg => {
    const role = msg.role === 'user' ? '👤 ユーザー' : '🤖 AI';
    text += `${role} (${msg.timestamp})\n`;
    text += `${msg.content}\n`;
    text += `トークン数: ${msg.tokenCount}\n`;
    text += `${'-'.repeat(30)}\n\n`;
  });
  
  return {
    success: true,
    format: 'text',
    content: text,
    mimeType: 'text/plain',
    filename: `chat_export_${data.userId}_${Date.now()}.txt`
  };
}

/**
 * Google Driveにエクスポート
 */
function exportToDrive(data, userId) {
  try {
    // Google Docsドキュメントを作成
    const doc = DocumentApp.create(`チャット履歴_${userId}_${new Date().toLocaleDateString('ja-JP')}`);
    const body = doc.getBody();
    
    // タイトルと基本情報
    const title = body.appendParagraph('チャット履歴');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    body.appendParagraph(`ユーザーID: ${data.userId}`);
    body.appendParagraph(`エクスポート日時: ${new Date().toLocaleString('ja-JP')}`);
    body.appendParagraph('');
    
    // メッセージを追加
    data.messages.forEach(msg => {
      const role = msg.role === 'user' ? 'ユーザー' : 'AI';
      const msgPara = body.appendParagraph(`【${role}】 ${msg.timestamp}`);
      msgPara.setBold(true);
      
      body.appendParagraph(msg.content);
      body.appendParagraph(`トークン数: ${msg.tokenCount}`);
      body.appendParagraph('');
    });
    
    // ドキュメントのURLを取得
    const url = doc.getUrl();
    
    return {
      success: true,
      format: 'drive',
      documentUrl: url,
      documentId: doc.getId(),
      message: `Google Docsにエクスポートしました: ${url}`
    };
  } catch (error) {
    console.error('Drive エクスポートエラー:', error);
    throw new Error('Google Driveへのエクスポートに失敗しました: ' + error.message);
  }
}

/**
 * 設定シートの状態を確認（デバッグ用）
 * @returns {Object} シートの状態情報
 */
function checkSettingsSheet() {
  try {
    const sheet = getSettingsSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow > 0 && lastCol > 0) {
      const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      return {
        status: 'exists',
        rows: lastRow,
        columns: lastCol,
        data: data,
        settings: getSettings()
      };
    } else {
      return {
        status: 'empty',
        rows: 0,
        columns: 0,
        data: [],
        settings: {}
      };
    }
  } catch (error) {
    return {
      status: 'error',
      error: error.toString()
    };
  }
}

/**
 * 設定シートを手動で初期化（メニューから実行用）
 */
function reinitializeSettingsSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // ユーザーに確認
  const response = ui.alert(
    '設定シートの初期化',
    '設定シートを初期化します。\n現在の設定はすべてデフォルト値にリセットされます。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('初期化をキャンセルしました。');
    return;
  }
  
  try {
    const ss = getSpreadsheet();
    const existingSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
    
    // 既存のシートがある場合は削除
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // 新しいシートを作成して初期化
    const newSheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
    initializeSettingsSheet(newSheet);
    
    // 成功メッセージ
    ui.alert(
      '初期化完了',
      '設定シートを正常に初期化しました。\nWebアプリをリロードすると新しい設定が反映されます。',
      ui.ButtonSet.OK
    );
    
    return '設定シートを再初期化しました';
  } catch (error) {
    ui.alert(
      'エラー',
      '設定シートの初期化中にエラーが発生しました:\n' + error.toString(),
      ui.ButtonSet.OK
    );
    return 'エラー: ' + error.toString();
  }
}