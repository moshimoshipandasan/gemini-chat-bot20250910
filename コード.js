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
  
  // セッション管理設定
  SESSION: {
    ENABLED: true,
    MAX_SESSIONS_PER_USER: 5,
    SESSION_TIMEOUT: 3600000, // 1時間（ミリ秒）
    AUTO_SAVE_INTERVAL: 60000, // 1分ごとに自動保存
    PERSIST_TO_PROPERTIES: true,
    MAX_PERSISTENT_SESSIONS: 20
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
  
  // パフォーマンス設定
  PERFORMANCE: {
    BATCH_LOG_ENABLED: true,
    BATCH_LOG_SIZE: 10,
    CACHE_PROMPT: true,
    CACHE_SETTINGS: true,
    PROMPT_CACHE_DURATION: 3600, // 1時間
    SETTINGS_CACHE_DURATION: 600, // 10分
    MAX_CONCURRENT_REQUESTS: 5,
    REQUEST_QUEUE_TIMEOUT: 30000, // 30秒
    ENABLE_PERFORMANCE_MONITORING: true
  },
  
  // エラーメッセージ
  ERRORS: {
    NO_API_KEY: 'APIキーが設定されていません',
    API_CALL_FAILED: 'Gemini APIの呼び出しに失敗しました',
    NO_RESPONSE: 'Gemini APIからの応答がありません',
    INVALID_USER_ID: 'ユーザーIDが無効です',
    INVALID_MESSAGE: 'メッセージが無効です',
    // ユーザー向けメッセージ
    USER_RATE_LIMIT: '現在リクエストが多いため、少々お待ちください。30秒後に再度お試しください。',
    USER_SERVICE_UNAVAILABLE: 'AIサービスが一時的に利用できません。数分後に再度お試しください。',
    USER_NETWORK_ERROR: 'ネットワークエラーが発生しました。インターネット接続を確認してください。',
    USER_TIMEOUT: '応答時間が長すぎます。もう一度お試しください。',
    USER_API_ERROR: 'AI APIに問題が発生しています。管理者にお問い合わせください。',
    USER_GENERIC_ERROR: 'エラーが発生しました。しばらくしてから再度お試しください。'
  }
};

// パフォーマンス最適化: バッチログバッファ
let logBuffer = [];
let logBufferTimer = null;

// パフォーマンス最適化: キャッシュ
const performanceCache = {
  prompt: null,
  promptExpiry: 0,
  settings: null,
  settingsExpiry: 0,
  apiKey: null
};

// パフォーマンスメトリクス
const performanceMetrics = {
  apiCalls: 0,
  cacheHits: 0,
  cacheMisses: 0,
  averageResponseTime: 0,
  totalResponseTime: 0
};

/**
 * セッションマネージャークラス
 * 高度なセッション管理機能を提供
 */
class SessionManager {
  constructor() {
    this.cache = CacheService.getScriptCache();
    this.properties = PropertiesService.getUserProperties();
  }
  
  /**
   * 新しいセッションを作成
   * @param {string} userId - ユーザーID
   * @returns {Object} セッション情報
   */
  createSession(userId) {
    const sessionId = Utilities.getUuid();
    const session = {
      id: sessionId,
      userId: userId,
      startTime: new Date().toISOString(),
      lastActivity: new Date().toISOString(),
      messageCount: 0,
      tokenCount: 0,
      metadata: {},
      status: 'active'
    };
    
    this.saveSession(sessionId, session);
    return session;
  }
  
  /**
   * セッションを取得
   * @param {string} sessionId - セッションID
   * @returns {Object|null} セッション情報
   */
  getSession(sessionId) {
    if (!sessionId) return null;
    
    // まずキャッシュから取得
    let session = this.cache.get(`session_${sessionId}`);
    if (session) {
      return JSON.parse(session);
    }
    
    // キャッシュになければPropertiesから取得
    if (CONFIG.SESSION.PERSIST_TO_PROPERTIES) {
      session = this.properties.getProperty(`session_${sessionId}`);
      if (session) {
        const parsedSession = JSON.parse(session);
        // キャッシュに復元
        this.cache.put(`session_${sessionId}`, session, CONFIG.CACHE.DURATION);
        return parsedSession;
      }
    }
    
    return null;
  }
  
  /**
   * セッションを保存
   * @param {string} sessionId - セッションID
   * @param {Object} session - セッション情報
   */
  saveSession(sessionId, session) {
    const sessionJson = JSON.stringify(session);
    
    // キャッシュに保存
    this.cache.put(`session_${sessionId}`, sessionJson, CONFIG.CACHE.DURATION);
    
    // Propertiesにも保存（長期保存）
    if (CONFIG.SESSION.PERSIST_TO_PROPERTIES) {
      this.properties.setProperty(`session_${sessionId}`, sessionJson);
    }
  }
  
  /**
   * セッションを更新
   * @param {string} sessionId - セッションID
   * @param {Object} updates - 更新内容
   */
  updateSession(sessionId, updates) {
    const session = this.getSession(sessionId);
    if (session) {
      Object.assign(session, updates, {
        lastActivity: new Date().toISOString()
      });
      this.saveSession(sessionId, session);
    }
  }
  
  /**
   * ユーザーのすべてのセッションを取得
   * @param {string} userId - ユーザーID
   * @returns {Array} セッションリスト
   */
  getUserSessions(userId) {
    const sessions = [];
    const keys = this.properties.getKeys();
    
    for (const key of keys) {
      if (key.startsWith('session_')) {
        try {
          const session = JSON.parse(this.properties.getProperty(key));
          if (session.userId === userId) {
            sessions.push(session);
          }
        } catch (e) {
          console.error(`セッション取得エラー: ${key}`, e);
        }
      }
    }
    
    // 最終アクティビティでソート
    return sessions.sort((a, b) => 
      new Date(b.lastActivity) - new Date(a.lastActivity)
    );
  }
  
  /**
   * セッションをエクスポート
   * @param {string} sessionId - セッションID
   * @returns {string} エクスポートされたJSON
   */
  exportSession(sessionId) {
    const session = this.getSession(sessionId);
    if (!session) return null;
    
    // 会話履歴も含めてエクスポート
    const history = getConversationHistory(session.userId);
    
    return JSON.stringify({
      session: session,
      history: history[session.userId] || [],
      exportDate: new Date().toISOString(),
      version: '1.0'
    }, null, 2);
  }
  
  /**
   * セッションをインポート
   * @param {string} jsonData - インポートするJSONデータ
   * @returns {Object} インポートされたセッション
   */
  importSession(jsonData) {
    try {
      const data = JSON.parse(jsonData);
      const session = data.session;
      const history = data.history;
      
      // 新しいセッションIDを生成
      session.id = Utilities.getUuid();
      session.importDate = new Date().toISOString();
      
      // セッションを保存
      this.saveSession(session.id, session);
      
      // 履歴を復元
      if (history && history.length > 0) {
        const conversationHistory = getConversationHistory();
        conversationHistory[session.userId] = history;
        saveConversationHistory(conversationHistory);
      }
      
      return session;
    } catch (error) {
      console.error('セッションインポートエラー:', error);
      throw new Error('セッションのインポートに失敗しました');
    }
  }
  
  /**
   * タイムアウトしたセッションをクリーンアップ
   */
  cleanupTimedOutSessions() {
    const keys = this.properties.getKeys();
    const now = new Date();
    let cleanedCount = 0;
    
    for (const key of keys) {
      if (key.startsWith('session_')) {
        try {
          const session = JSON.parse(this.properties.getProperty(key));
          const lastActivity = new Date(session.lastActivity);
          
          if (now - lastActivity > CONFIG.SESSION.SESSION_TIMEOUT) {
            session.status = 'expired';
            this.saveSession(session.id, session);
            cleanedCount++;
          }
        } catch (e) {
          console.error(`セッションクリーンアップエラー: ${key}`, e);
        }
      }
    }
    
    console.log(`${cleanedCount}件のタイムアウトセッションをクリーンアップ`);
    return cleanedCount;
  }
}

// グローバルインスタンス
const sessionManager = new SessionManager();

/**
 * APIキーを取得（キャッシュ付き）
 * @returns {string} Gemini APIキー
 * @throws {Error} APIキーが設定されていない場合
 */
function getApiKey() {
  // キャッシュから取得
  if (performanceCache.apiKey) {
    return performanceCache.apiKey;
  }
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('apikey');
  if (!apiKey) {
    throw new Error(CONFIG.ERRORS.NO_API_KEY);
  }
  
  // キャッシュに保存
  performanceCache.apiKey = apiKey;
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
  ui.createMenu('チャットボット管理')
    .addItem('📊 統計情報を表示', 'showStatistics')
    .addSeparator()
    .addItem('💾 ログをエクスポート', 'exportLogs')
    .addItem('🗑️ 古いログを削除', 'cleanOldLogs')
    .addSeparator()
    .addItem('🔍 システムヘルスチェック', 'checkSystemHealth')
    .addItem('👥 アクティブユーザー一覧', 'showActiveUsers')
    .addItem('📊 パフォーマンスモニター', 'showPerformanceMetrics')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔐 セッション管理')
      .addItem('📁 セッション一覧', 'showSessionList')
      .addItem('💾 セッションをエクスポート', 'exportSessionDialog')
      .addItem('📂 セッションをインポート', 'importSessionDialog')
      .addItem('🧹 期限切れセッションをクリーンアップ', 'cleanupSessions'))
    .addSeparator()
    .addItem('🔄 キャッシュをクリア', 'clearCache')
    .addItem('⚙️ 設定シートを初期化', 'reinitializeSettingsSheet')
    .addToUi();
}

/**
 * 統計情報を表示
 * 使用状況の統計を表示する
 */
function showStatistics() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOG);
    if (!sheet || sheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert('統計情報', 'ログデータがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    
    // 統計を計算
    const stats = {
      totalMessages: data.length,
      uniqueUsers: new Set(data.map(row => row[1])).size,
      totalTokens: data.reduce((sum, row) => sum + (row[4] || 0), 0),
      userMessages: data.filter(row => row[2] === 'user').length,
      assistantMessages: data.filter(row => row[2] === 'assistant').length,
      firstMessage: data[0][0],
      lastMessage: data[data.length - 1][0]
    };
    
    // 日付範囲を計算
    const days = Math.ceil((new Date(stats.lastMessage) - new Date(stats.firstMessage)) / (1000 * 60 * 60 * 24)) || 1;
    
    const message = `
📊 チャットボット統計情報
━━━━━━━━━━━━━━━━━━━━━━━

📝 メッセージ統計:
  • 総メッセージ数: ${stats.totalMessages.toLocaleString()}
  • ユーザーメッセージ: ${stats.userMessages.toLocaleString()}
  • AIレスポンス: ${stats.assistantMessages.toLocaleString()}
  
👥 ユーザー統計:
  • ユニークユーザー数: ${stats.uniqueUsers}
  • 平均メッセージ/ユーザー: ${Math.round(stats.totalMessages / stats.uniqueUsers)}
  
🔢 トークン使用量:
  • 総トークン数: ${stats.totalTokens.toLocaleString()}
  • 平均トークン/メッセージ: ${Math.round(stats.totalTokens / stats.assistantMessages)}
  
📅 期間:
  • 開始日: ${new Date(stats.firstMessage).toLocaleString('ja-JP')}
  • 最終日: ${new Date(stats.lastMessage).toLocaleString('ja-JP')}
  • 稼働日数: ${days}日
  • 日平均メッセージ: ${Math.round(stats.totalMessages / days)}
`;
    
    SpreadsheetApp.getUi().alert('統計情報', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('統計情報表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '統計情報の取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ログをエクスポート
 * ログシートを新しいシートにコピーする
 */
function exportLogs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
    
    if (!logSheet || logSheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert('エクスポート', 'エクスポートするログがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // エクスポート用の新しいシート名を生成
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const exportSheetName = `ログ_エクスポート_${timestamp}`;
    
    // シートをコピー
    const exportSheet = logSheet.copyTo(ss);
    exportSheet.setName(exportSheetName);
    
    // 成功メッセージ
    SpreadsheetApp.getUi().alert(
      'エクスポート完了',
      `ログを「${exportSheetName}」シートにエクスポートしました。\n\n` +
      `エクスポート件数: ${logSheet.getLastRow() - 1}件`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log(`ログエクスポート完了: ${exportSheetName}`);
    
  } catch (error) {
    console.error('ログエクスポートエラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'ログのエクスポートに失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 古いログを削除
 * 指定日数より古いログを削除する
 */
function cleanOldLogs() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 削除する日数を入力
    const response = ui.prompt(
      '古いログの削除',
      '何日より古いログを削除しますか？（例: 30）',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const days = parseInt(response.getResponseText());
    if (isNaN(days) || days <= 0) {
      ui.alert('エラー', '有効な日数を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOG);
    if (!sheet || sheet.getLastRow() <= 1) {
      ui.alert('削除', '削除するログがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // 削除対象の日付を計算
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    // データを取得
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    const newData = [];
    let deletedCount = 0;
    
    // 削除対象をフィルタリング
    for (const row of data) {
      const logDate = new Date(row[0]);
      if (logDate >= cutoffDate) {
        newData.push(row);
      } else {
        deletedCount++;
      }
    }
    
    if (deletedCount === 0) {
      ui.alert('削除', `${days}日より古いログは見つかりませんでした。`, ui.ButtonSet.OK);
      return;
    }
    
    // 確認ダイアログ
    const confirmResult = ui.alert(
      '削除の確認',
      `${deletedCount}件のログを削除します。\n\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResult !== ui.Button.YES) {
      return;
    }
    
    // データをクリアして新しいデータを書き込み
    if (newData.length > 0) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
      sheet.getRange(2, 1, newData.length, 5).setValues(newData);
    } else {
      // すべて削除される場合
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
    }
    
    ui.alert(
      '削除完了',
      `${deletedCount}件の古いログを削除しました。\n残りのログ: ${newData.length}件`,
      ui.ButtonSet.OK
    );
    
    console.log(`古いログ削除完了: ${deletedCount}件削除、${newData.length}件保持`);
    
  } catch (error) {
    console.error('ログ削除エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'ログの削除に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * システムヘルスチェック
 * システムの状態を確認する
 */
function checkSystemHealth() {
  const ui = SpreadsheetApp.getUi();
  const checks = [];
  
  try {
    // 1. APIキーの確認
    let apiKeyStatus = '❌ 未設定';
    try {
      const apiKey = getApiKey();
      if (apiKey) {
        apiKeyStatus = '✅ 設定済み';
      }
    } catch (e) {
      apiKeyStatus = '❌ 未設定';
    }
    checks.push(`APIキー: ${apiKeyStatus}`);
    
    // 2. シートの確認
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = [CONFIG.SHEETS.PROMPT, CONFIG.SHEETS.LOG, CONFIG.SHEETS.SETTINGS];
    const sheetStatuses = [];
    
    for (const sheetName of requiredSheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const rows = sheet.getLastRow();
        sheetStatuses.push(`  • ${sheetName}: ✅ (${rows}行)`);
      } else {
        sheetStatuses.push(`  • ${sheetName}: ❌ 見つかりません`);
      }
    }
    checks.push('シート状態:\n' + sheetStatuses.join('\n'));
    
    // 3. プロンプトの確認
    let promptStatus = '❌ 未設定';
    try {
      const promptSheet = ss.getSheetByName(CONFIG.SHEETS.PROMPT);
      if (promptSheet) {
        const prompt = promptSheet.getRange(CONFIG.SHEETS.SYSTEM_PROMPT_CELL).getValue();
        if (prompt && prompt.length > 0) {
          promptStatus = `✅ 設定済み (${prompt.length}文字)`;
        }
      }
    } catch (e) {
      promptStatus = '❌ エラー';
    }
    checks.push(`システムプロンプト: ${promptStatus}`);
    
    // 4. キャッシュの確認
    let cacheStatus = '❌ 空';
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(CONFIG.CACHE.KEY);
      if (cached) {
        const history = JSON.parse(cached);
        const userCount = Object.keys(history).length;
        cacheStatus = `✅ アクティブ (${userCount}ユーザー)`;
      } else {
        cacheStatus = '⚪ 空（正常）';
      }
    } catch (e) {
      cacheStatus = '❌ エラー';
    }
    checks.push(`キャッシュ状態: ${cacheStatus}`);
    
    // 5. Web App URLの確認
    let webAppStatus = '❓ 確認できません';
    try {
      const url = ScriptApp.getService().getUrl();
      if (url) {
        webAppStatus = '✅ デプロイ済み';
      }
    } catch (e) {
      webAppStatus = '❓ 未デプロイの可能性';
    }
    checks.push(`Webアプリ: ${webAppStatus}`);
    
    // 結果を表示
    const message = `
🔍 システムヘルスチェック結果
━━━━━━━━━━━━━━━━━━━━━━━

${checks.join('\n\n')}

━━━━━━━━━━━━━━━━━━━━━━━
実行時刻: ${new Date().toLocaleString('ja-JP')}
`;
    
    ui.alert('システムヘルスチェック', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('ヘルスチェックエラー:', error);
    ui.alert('エラー', 'システムヘルスチェックに失敗しました: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * アクティブユーザー一覧を表示
 * 最近利用したユーザーの一覧を表示する
 */
function showActiveUsers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOG);
    if (!sheet || sheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert('ユーザー一覧', 'ログデータがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    
    // ユーザーごとの統計を集計
    const userStats = {};
    for (const row of data) {
      const userId = row[1];
      const timestamp = row[0];
      const role = row[2];
      const tokens = row[4] || 0;
      
      if (!userStats[userId]) {
        userStats[userId] = {
          messages: 0,
          tokens: 0,
          firstSeen: timestamp,
          lastSeen: timestamp
        };
      }
      
      userStats[userId].messages++;
      if (role === 'assistant') {
        userStats[userId].tokens += tokens;
      }
      if (timestamp > userStats[userId].lastSeen) {
        userStats[userId].lastSeen = timestamp;
      }
      if (timestamp < userStats[userId].firstSeen) {
        userStats[userId].firstSeen = timestamp;
      }
    }
    
    // ユーザーリストを作成（最終利用日でソート）
    const users = Object.entries(userStats)
      .sort((a, b) => new Date(b[1].lastSeen) - new Date(a[1].lastSeen))
      .slice(0, 20); // 上位20ユーザーまで表示
    
    let message = `
👥 アクティブユーザー一覧 (上位${Math.min(users.length, 20)}名)
━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
    
    users.forEach(([userId, stats], index) => {
      const lastSeenDate = new Date(stats.lastSeen);
      const daysAgo = Math.floor((new Date() - lastSeenDate) / (1000 * 60 * 60 * 24));
      const daysAgoText = daysAgo === 0 ? '今日' : daysAgo === 1 ? '昨日' : `${daysAgo}日前`;
      
      message += `${index + 1}. ${userId}\n`;
      message += `   📝 メッセージ: ${stats.messages}件\n`;
      message += `   🔢 トークン: ${stats.tokens.toLocaleString()}\n`;
      message += `   📅 最終利用: ${daysAgoText}\n`;
      message += `\n`;
    });
    
    message += `━━━━━━━━━━━━━━━━━━━━━━━\n`;
    message += `総ユーザー数: ${Object.keys(userStats).length}名`;
    
    SpreadsheetApp.getUi().alert('アクティブユーザー', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('ユーザー一覧表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'ユーザー一覧の取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * パフォーマンスメトリクスを表示
 * システムのパフォーマンス情報を表示する
 */
function showPerformanceMetrics() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // キャッシュヒット率を計算
    const totalCacheAccess = performanceMetrics.cacheHits + performanceMetrics.cacheMisses;
    const cacheHitRate = totalCacheAccess > 0 
      ? Math.round((performanceMetrics.cacheHits / totalCacheAccess) * 100) 
      : 0;
    
    // 平均応答時間を計算
    const avgResponseTime = performanceMetrics.apiCalls > 0
      ? Math.round(performanceMetrics.totalResponseTime / performanceMetrics.apiCalls)
      : 0;
    
    // バッファ状態
    const bufferStatus = logBuffer.length > 0 
      ? `${logBuffer.length}/${CONFIG.PERFORMANCE.BATCH_LOG_SIZE} 件` 
      : '空';
    
    const message = `
📊 パフォーマンスメトリクス
━━━━━━━━━━━━━━━━━━━━━━━

🔄 キャッシュ統計:
  • キャッシュヒット: ${performanceMetrics.cacheHits.toLocaleString()}回
  • キャッシュミス: ${performanceMetrics.cacheMisses.toLocaleString()}回
  • ヒット率: ${cacheHitRate}%
  
🌐 API統計:
  • APIコール数: ${performanceMetrics.apiCalls.toLocaleString()}回
  • 平均応答時間: ${avgResponseTime}ms
  
💾 バッチログ状態:
  • バッファ: ${bufferStatus}
  • バッチモード: ${CONFIG.PERFORMANCE.BATCH_LOG_ENABLED ? '有効' : '無効'}
  
⚙️ 最適化設定:
  • プロンプトキャッシュ: ${CONFIG.PERFORMANCE.CACHE_PROMPT ? '有効' : '無効'}
  • 設定キャッシュ: ${CONFIG.PERFORMANCE.CACHE_SETTINGS ? '有効' : '無効'}
  • パフォーマンスモニタリング: ${CONFIG.PERFORMANCE.ENABLE_PERFORMANCE_MONITORING ? '有効' : '無効'}

━━━━━━━━━━━━━━━━━━━━━━━
計測開始: サーバー起動時
現在時刻: ${new Date().toLocaleString('ja-JP')}
`;
    
    ui.alert('パフォーマンスメトリクス', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('パフォーマンスメトリクス表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'メトリクスの取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * セッション一覧を表示
 */
function showSessionList() {
  try {
    const ui = SpreadsheetApp.getUi();
    const allSessions = [];
    const keys = sessionManager.properties.getKeys();
    
    for (const key of keys) {
      if (key.startsWith('session_')) {
        try {
          const session = JSON.parse(sessionManager.properties.getProperty(key));
          allSessions.push(session);
        } catch (e) {
          console.error(`セッション読み込みエラー: ${key}`, e);
        }
      }
    }
    
    // 最終アクティビティでソート
    allSessions.sort((a, b) => new Date(b.lastActivity) - new Date(a.lastActivity));
    
    if (allSessions.length === 0) {
      ui.alert('セッション一覧', 'アクティブなセッションがありません。', ui.ButtonSet.OK);
      return;
    }
    
    let message = `
🔐 セッション一覧 (上位${Math.min(allSessions.length, 10)}件)
━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
    
    allSessions.slice(0, 10).forEach((session, index) => {
      const startTime = new Date(session.startTime).toLocaleString('ja-JP');
      const lastActivity = new Date(session.lastActivity).toLocaleString('ja-JP');
      const duration = Math.round((new Date(session.lastActivity) - new Date(session.startTime)) / 60000); // 分
      
      message += `${index + 1}. セッションID: ${session.id.substr(0, 8)}...\n`;
      message += `   👤 ユーザー: ${session.userId}\n`;
      message += `   💬 メッセージ数: ${session.messageCount || 0}\n`;
      message += `   🔢 トークン数: ${session.tokenCount || 0}\n`;
      message += `   ⌚ 期間: ${duration}分\n`;
      message += `   📅 開始: ${startTime}\n`;
      message += `   🔄 最終活動: ${lastActivity}\n`;
      message += `   🌐 状態: ${session.status || 'active'}\n`;
      message += `\n`;
    });
    
    message += `━━━━━━━━━━━━━━━━━━━━━━━\n`;
    message += `総セッション数: ${allSessions.length}`;
    
    ui.alert('セッション一覧', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('セッション一覧表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'セッション一覧の取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * セッションをエクスポートするダイアログ
 */
function exportSessionDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'セッションエクスポート',
      'エクスポートするセッションIDを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const sessionId = response.getResponseText();
    if (!sessionId) {
      ui.alert('エラー', 'セッションIDを入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const exportData = sessionManager.exportSession(sessionId);
    if (!exportData) {
      ui.alert('エラー', '指定されたセッションが見つかりません。', ui.ButtonSet.OK);
      return;
    }
    
    // エクスポートデータを新しいシートに保存
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const exportSheet = ss.insertSheet(`セッション_エクスポート_${new Date().getTime()}`);
    exportSheet.getRange(1, 1).setValue(exportData);
    
    ui.alert(
      'エクスポート完了',
      `セッションデータを新しいシートにエクスポートしました。\nシート名: ${exportSheet.getName()}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('セッションエクスポートエラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'セッションのエクスポートに失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * セッションをインポートするダイアログ
 */
function importSessionDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'セッションインポート',
      'インポートするセッションデータが含まれるシート名を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const sheetName = response.getResponseText();
    if (!sheetName) {
      ui.alert('エラー', 'シート名を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const importSheet = ss.getSheetByName(sheetName);
    if (!importSheet) {
      ui.alert('エラー', `シート「${sheetName}」が見つかりません。`, ui.ButtonSet.OK);
      return;
    }
    
    const jsonData = importSheet.getRange(1, 1).getValue();
    if (!jsonData) {
      ui.alert('エラー', 'シートにデータがありません。', ui.ButtonSet.OK);
      return;
    }
    
    const session = sessionManager.importSession(jsonData);
    
    ui.alert(
      'インポート完了',
      `セッションをインポートしました。\n新しいセッションID: ${session.id}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('セッションインポートエラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'セッションのインポートに失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 期限切れセッションをクリーンアップ
 */
function cleanupSessions() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'セッションクリーンアップ',
      '1時間以上アクティビティがないセッションを期限切れとしてマークします。\n続行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return;
    }
    
    const cleanedCount = sessionManager.cleanupTimedOutSessions();
    
    ui.alert(
      'クリーンアップ完了',
      `${cleanedCount}件のセッションを期限切れとしてマークしました。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('セッションクリーンアップエラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'セッションのクリーンアップに失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
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
 * システムプロンプトを取得（キャッシュ付き）
 * @returns {string} システムプロンプト
 */
function getSystemPrompt() {
  // キャッシュチェック
  if (CONFIG.PERFORMANCE.CACHE_PROMPT) {
    const now = Date.now();
    if (performanceCache.prompt && performanceCache.promptExpiry > now) {
      performanceMetrics.cacheHits++;
      return performanceCache.prompt;
    }
    performanceMetrics.cacheMisses++;
  }
  
  const sheet = getSpreadsheet().getSheetByName(CONFIG.SHEETS.PROMPT);
  if (!sheet) {
    throw new Error(`シート「${CONFIG.SHEETS.PROMPT}」が見つかりません`);
  }
  const result = sheet.getRange(CONFIG.SHEETS.SYSTEM_PROMPT_CELL).getValue();
  
  // キャッシュに保存
  if (CONFIG.PERFORMANCE.CACHE_PROMPT) {
    performanceCache.prompt = result;
    performanceCache.promptExpiry = Date.now() + (CONFIG.PERFORMANCE.PROMPT_CACHE_DURATION * 1000);
  }
  
  return result;
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
 * チャットログを記録（バッチ処理対応）
 * @param {string} userId - ユーザーID
 * @param {string} role - 役割（'user' または 'model'）
 * @param {string} message - メッセージ内容
 */
function logChat(userId, role, message) {
  try {
    const timestamp = new Date().toLocaleString('ja-JP', { timeZone: CONFIG.TIMEZONE });
    const tokenCount = estimateTokenCount(message);
    const rowData = [
      timestamp,
      userId,
      role,
      message,
      tokenCount
    ];
    
    if (CONFIG.PERFORMANCE.BATCH_LOG_ENABLED) {
      // バッチログバッファに追加
      logBuffer.push(rowData);
      
      // バッファがいっぱいになったら即座にフラッシュ
      if (logBuffer.length >= CONFIG.PERFORMANCE.BATCH_LOG_SIZE) {
        flushLogBuffer();
      }
      // タイマーは使用しない（GASでは利用不可）
      // 代わりにprocessMessageの最後で手動フラッシュ
    } else {
      // バッチ処理無効の場合は即座に書き込み
      const logSheet = getLogSheet();
      logSheet.appendRow(rowData);
    }
  } catch (error) {
    console.error('ログの記録に失敗しました:', error);
  }
}

/**
 * ログバッファをフラッシュ（バッチ書き込み）
 */
function flushLogBuffer() {
  if (logBuffer.length === 0) return;
  
  try {
    const logSheet = getLogSheet();
    const lastRow = logSheet.getLastRow();
    
    // バッチで書き込み（高速化）
    if (logBuffer.length > 0) {
      const range = logSheet.getRange(lastRow + 1, 1, logBuffer.length, 5);
      range.setValues(logBuffer);
      
      console.log(`バッチログ書き込み: ${logBuffer.length}件`);
    }
    
    // バッファをクリア
    logBuffer = [];
    
    // タイマー変数もクリア（互換性のため残す）
    logBufferTimer = null;
  } catch (error) {
    console.error('バッチログ書き込みエラー:', error);
    // エラー時は通常の書き込みにフォールバック
    const logSheet = getLogSheet();
    logBuffer.forEach(row => {
      try {
        logSheet.appendRow(row);
      } catch (e) {
        console.error('個別ログ書き込みエラー:', e);
      }
    });
    logBuffer = [];
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
    // 会話履歴を取得（userIdを渡してキャッシュ復元に対応）
    const conversationHistory = getConversationHistory(userId);
    
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
    
    // バッチログをフラッシュ（重要：GASではタイマーが使えないため手動フラッシュ）
    if (CONFIG.PERFORMANCE.BATCH_LOG_ENABLED && logBuffer.length > 0) {
      flushLogBuffer();
    }
    
    return response;
  } catch (error) {
    console.error('メッセージ処理エラー:', error);
    
    // ユーザーフレンドリーなエラーメッセージを返す
    let userMessage = CONFIG.ERRORS.USER_GENERIC_ERROR;
    
    // エラーの種類に応じて適切なメッセージを返す
    if (error.message && error.message.includes('429')) {
      userMessage = CONFIG.ERRORS.USER_RATE_LIMIT;
    } else if (error.message && error.message.includes('503')) {
      userMessage = CONFIG.ERRORS.USER_SERVICE_UNAVAILABLE;
    } else if (error.message && (error.message.includes('Network') || error.message.includes('UrlFetch'))) {
      userMessage = CONFIG.ERRORS.USER_NETWORK_ERROR;
    } else if (error.message && error.message.includes('Timeout')) {
      userMessage = CONFIG.ERRORS.USER_TIMEOUT;
    } else if (error.message && error.message.includes('API')) {
      userMessage = CONFIG.ERRORS.USER_API_ERROR;
    } else if (error.message && error.message.includes('API_KEY')) {
      userMessage = 'APIキーの設定に問題があります。管理者にお問い合わせください。';
    } else if (error.message && error.message.includes('INVALID')) {
      // 元のエラーメッセージをそのまま返す（入力検証エラー）
      throw error;
    }
    
    // エラーをログに記録（エラーメッセージとして）
    try {
      logChat(userId, 'system', `[エラー] ${error.message || 'Unknown error'}`);
    } catch (logError) {
      console.error('エラーログの記録に失敗:', logError);
    }
    
    // エラー時もバッチログをフラッシュ
    if (CONFIG.PERFORMANCE.BATCH_LOG_ENABLED && logBuffer.length > 0) {
      flushLogBuffer();
    }
    
    // ユーザーフレンドリーなメッセージを返す（エラーをthrowしない）
    return userMessage;
  }
}

/**
 * 会話履歴を取得（キャッシュ期限切れ時はログシートから復元）
 * @param {string} userId - ユーザーID（キャッシュ復元時に必要）
 * @returns {Object} 会話履歴オブジェクト
 */
function getConversationHistory(userId = null) {
  const cache = getCache();
  const cached = cache.get(CONFIG.CACHE.KEY);
  
  if (cached) {
    return JSON.parse(cached);
  }
  
  // キャッシュが期限切れの場合、ログシートから復元を試みる
  if (userId) {
    console.log('キャッシュ期限切れ: ログシートから履歴を復元します');
    try {
      const restoredHistory = restoreHistoryFromLogSheet(userId);
      if (restoredHistory && restoredHistory.length > 0) {
        const conversationHistory = { [userId]: restoredHistory };
        // 復元した履歴をキャッシュに保存
        saveConversationHistory(conversationHistory);
        console.log(`${restoredHistory.length}件の会話履歴を復元しました`);
        return conversationHistory;
      }
    } catch (error) {
      console.error('履歴の復元に失敗しました:', error);
    }
  }
  
  return {};
}

/**
 * ログシートから会話履歴を復元
 * @param {string} userId - ユーザーID
 * @param {number} limit - 取得する最大メッセージ数
 * @returns {Array} 会話履歴配列
 */
function restoreHistoryFromLogSheet(userId, limit = CONFIG.CONVERSATION.MAX_HISTORY_LENGTH) {
  try {
    const logSheet = getLogSheet();
    const lastRow = logSheet.getLastRow();
    
    if (lastRow <= 1) {
      return []; // ヘッダーのみの場合
    }
    
    // データ範囲を取得（ヘッダーを除く）
    const data = logSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const userMessages = [];
    
    // 最新のメッセージから逆順に取得
    for (let i = data.length - 1; i >= 0 && userMessages.length < limit; i--) {
      const row = data[i];
      const logUserId = row[1]; // ユーザーID列
      const role = row[2]; // 役割列
      const message = row[3]; // メッセージ列
      
      // 該当ユーザーのメッセージのみ取得
      if (logUserId === userId && message) {
        // systemロールのエラーメッセージは除外
        if (role !== 'system') {
          userMessages.unshift({
            role: role,
            parts: [{ text: message }]
          });
        }
      }
    }
    
    return userMessages;
  } catch (error) {
    console.error('ログシートからの復元エラー:', error);
    return [];
  }
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
 * 会話履歴を最大長に制限（スマートトリミング）
 * @param {Array} history - 会話履歴配列
 */
function trimConversationHistory(history) {
  if (history.length > CONFIG.CONVERSATION.MAX_HISTORY_LENGTH) {
    // システムメッセージを保持しつつ、古いユーザーメッセージから削除
    const systemMessages = history.filter(msg => msg.role === 'system');
    const conversationMessages = history.filter(msg => msg.role !== 'system');
    
    const trimmedConversation = conversationMessages.slice(
      conversationMessages.length - CONFIG.CONVERSATION.MAX_HISTORY_LENGTH
    );
    
    // システムメッセージとトリミングされた会話を結合
    history.length = 0;
    history.push(...systemMessages, ...trimmedConversation);
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
      
      // パフォーマンス計測
      const startTime = Date.now();
      const response = UrlFetchApp.fetch(getGeminiUrl(), options);
      const responseCode = response.getResponseCode();
      const responseTime = Date.now() - startTime;
      
      // メトリクス更新
      if (CONFIG.PERFORMANCE.ENABLE_PERFORMANCE_MONITORING) {
        performanceMetrics.apiCalls++;
        performanceMetrics.totalResponseTime += responseTime;
        performanceMetrics.averageResponseTime = performanceMetrics.totalResponseTime / performanceMetrics.apiCalls;
      }
      
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
    const history = getConversationHistory(userId);  // userIdを渡して復元対応
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