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
    SYSTEM_PROMPT_CELL: 'B5'
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
    .addItem('プロンプト自動生成', 'generateGeminiPrompt')
    .addSeparator()
    .addItem('設定シートを初期化', 'reinitializeSettingsSheet')
    .addToUi();
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
 * Gemini APIを呼び出す
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
  
  try {
    const response = UrlFetchApp.fetch(getGeminiUrl(), options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      console.error(`API エラー: ${responseCode}`);
      throw new Error(`${CONFIG.ERRORS.API_CALL_FAILED}: ${responseCode}`);
    }
    
    const responseJson = JSON.parse(response.getContentText());
    return extractResponseText(responseJson);
  } catch (error) {
    console.error('Gemini API呼び出しエラー:', error);
    if (error.message.includes(CONFIG.ERRORS.API_CALL_FAILED)) {
      throw error;
    }
    throw new Error(`${CONFIG.ERRORS.API_CALL_FAILED}: ${error.message}`);
  }
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
 * Gemini用プロンプトを自動生成
 * @returns {string} 生成されたプロンプト
 */
function generateGeminiPrompt() {
  try {
    const sheet = getSpreadsheet().getSheetByName(CONFIG.SHEETS.PROMPT);
    if (!sheet) {
      throw new Error(`シート「${CONFIG.SHEETS.PROMPT}」が見つかりません`);
    }
    
    const values = [
      sheet.getRange('B2').getValue(),
      sheet.getRange('B3').getValue(),
      sheet.getRange('B4').getValue()
    ].filter(v => v); // 空の値を除外
    
    if (values.length === 0) {
      throw new Error('プロンプト生成に必要な情報が入力されていません');
    }
    
    const baseText = `# タスクの説明：${values.join('、')}

## 役割・目標
あなたは、Gemini AIモデル用の効果的なプロンプトを自動生成するAIアシスタントです。目標は、ユーザーが指定したタスクに対して最適化された、明確で構造化されたプロンプトを作成することです。

## 視点・対象
- 主な対象：Gemini AIモデルを使用するユーザー
- 二次的な対象：Gemini AIモデル自体（プロンプトの受け手として）

## 制約条件
1. 生成されるプロンプトは、Gemini AIモデルの特性と制限を考慮に入れたものであること（マルチモーダル機能、コードの理解と生成能力、長文脈処理など）
2. プロンプトは明確で簡潔であること、ただし必要な詳細は省略しないこと
3. 特定の構造（役割・目標、制約条件など）を含めること
4. 倫理的で法的に問題のない内容であること
5. Geminiの機能と制限を正確に反映すること（トークン制限、知識のカットオフ日など）

## 処理手順 (Chain of Thought)
1. ユーザーの入力を分析し、要求されているタスクを特定する
2. タスクに適した役割と目標を定義する
3. 対象となる視点や読者を決定する
4. タスクに関連する制約条件をリストアップする
5. タスクを完了するための具体的な手順を考案する
6. 必要な入力情報を特定する
7. 期待される出力形式を決定する
8. 上記の要素を組み合わせて、構造化されたプロンプトを作成する
9. プロンプトを見直し、明確さと簡潔さを確認する
10. 必要に応じて微調整を行う

## 入力文
以下の形式で入力を受け付けます：

[タスクの説明]のGemini用プロンプトを役割・目標、視点・対象、制約条件、処理手順(CoT)、入力文、出力文を考慮して作成して

例：「レシピ生成」のGemini用プロンプトを役割・目標、視点・対象、制約条件、処理手順(CoT)、入力文、出力文を考慮して作成して

## 出力文
以下の構造に従ってプロンプトを生成します：

# [タスク名] Gemini Prompt

## 役割・目標
[役割と目標の説明]

## 視点・対象
[視点と対象の列挙]

## 制約条件
1. [制約条件1]
2. [制約条件2]
...

## 処理手順 (Chain of Thought)
1. [手順1]
2. [手順2]
...

## 入力文
[必要な入力情報の説明]
[入力例]

## 出力文
[期待される出力形式の説明]
[出力例]

このフォーマットに従って、要求されたタスクに最適化されたGemini用プロンプトを生成します。
`;

  const payload = {
    'contents': [{
      'parts': [{
        'text': baseText
      }]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(getGeminiUrl(), options);
    const responseJson = JSON.parse(response.getContentText());
    
    const generatedPrompt = extractResponseText(responseJson);
    sheet.getRange('B5').setValue(generatedPrompt);
    
    // 成功をユーザーに通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'プロンプトが正常に生成されました',
      '完了',
      3
    );
    
    return generatedPrompt;
  } catch (error) {
    console.error('プロンプト生成エラー:', error);
    const errorMessage = `プロンプト生成エラー: ${error.message}`;
    
    // エラーをユーザーに通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      errorMessage,
      'エラー',
      5
    );
    
    throw new Error(errorMessage);
  }
  } catch (error) {
    console.error('プロンプト生成関数エラー:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      error.message,
      'エラー',
      5
    );
    throw error;
  }
}

/**
 * 設定を取得
 * @returns {Object} 設定オブジェクト
 */
function getSettings() {
  try {
    const sheet = getSettingsSheet();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    const settings = {};
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
 * @param {Object} settings - 保存する設定
 * @returns {boolean} 成功/失敗
 */
function saveSettings(settings) {
  try {
    const sheet = getSettingsSheet();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    // 設定値を更新
    data.forEach((row, index) => {
      const key = row[0];
      if (key && settings.hasOwnProperty(key)) {
        sheet.getRange(index + 2, 2).setValue(settings[key]);
      }
    });
    
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
 * 会話をエクスポート
 * @param {string} userId - ユーザーID
 * @returns {Object} エクスポートデータ
 */
function exportConversation(userId) {
  try {
    const history = getConversationHistory();
    const userHistory = history[userId] || [];
    
    return {
      userId: userId,
      exportDate: new Date().toISOString(),
      messages: userHistory.map(msg => ({
        role: msg.role,
        content: msg.parts[0].text,
        timestamp: new Date().toISOString()
      }))
    };
  } catch (error) {
    console.error('会話のエクスポートエラー:', error);
    throw error;
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