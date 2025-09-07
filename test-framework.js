/**
 * テスト基盤 - Geminiチャットボット
 * Google Apps Script用のテストフレームワーク
 */

// ===== テスト設定 =====
const TEST_CONFIG = {
  SHEETS: {
    PROMPT: 'プロンプト',
    LOG: 'ログ',
    TEST_RESULTS: 'テスト結果'  // テスト結果記録用
  },
  TEST_USER_ID: 'test_user_001',
  TEST_TIMEOUT: 10000,  // 10秒
  SYSTEM_PROMPT_CELL: 'A1' // メインコードと同じセル位置
};

// ===== スプレッドシート初期化 =====

/**
 * テスト環境のセットアップ
 * スプレッドシートと必要なシートを作成
 */
function setupTestEnvironment() {
  console.log('🔧 テスト環境をセットアップ中...');
  
  try {
    // スプレッドシートの取得または作成
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. プロンプトシートの作成
    let promptSheet = ss.getSheetByName(TEST_CONFIG.SHEETS.PROMPT);
    if (!promptSheet) {
      promptSheet = ss.insertSheet(TEST_CONFIG.SHEETS.PROMPT);
      // A1セルにプロンプトを設定（メインコードと同じ）
      promptSheet.getRange(TEST_CONFIG.SYSTEM_PROMPT_CELL).setValue('あなたは親切なAIアシスタントです。ユーザーの質問に丁寧に日本語で答えてください。');
      console.log('✅ プロンプトシートを作成しました');
    } else {
      console.log('✓ プロンプトシートは既に存在します');
    }
    
    // 2. ログシートの作成
    let logSheet = ss.getSheetByName(TEST_CONFIG.SHEETS.LOG);
    if (!logSheet) {
      logSheet = ss.insertSheet(TEST_CONFIG.SHEETS.LOG);
      logSheet.getRange('A1:E1').setValues([['タイムスタンプ', 'ユーザーID', '役割', 'メッセージ', 'トークン数']]);
      logSheet.setFrozenRows(1);
      
      // 列幅の設定
      logSheet.setColumnWidth(1, 180);
      logSheet.setColumnWidth(2, 150);
      logSheet.setColumnWidth(3, 100);
      logSheet.setColumnWidth(4, 400);
      logSheet.setColumnWidth(5, 100);
      
      console.log('✅ ログシートを作成しました');
    } else {
      console.log('✓ ログシートは既に存在します');
    }
    
    // 3. テスト結果シートの作成
    let testSheet = ss.getSheetByName(TEST_CONFIG.SHEETS.TEST_RESULTS);
    if (!testSheet) {
      testSheet = ss.insertSheet(TEST_CONFIG.SHEETS.TEST_RESULTS);
      testSheet.getRange('A1:F1').setValues([['実行日時', 'テスト名', '結果', '実行時間(ms)', 'エラー', '詳細']]);
      testSheet.setFrozenRows(1);
      console.log('✅ テスト結果シートを作成しました');
    } else {
      console.log('✓ テスト結果シートは既に存在します');
    }
    
    console.log('🎉 テスト環境のセットアップが完了しました！');
    return true;
    
  } catch (error) {
    console.error('❌ セットアップエラー:', error);
    return false;
  }
}

// ===== テストヘルパー関数 =====

/**
 * アサーション関数
 */
const assert = {
  /**
   * 値が真であることを確認
   */
  isTrue: function(value, message) {
    if (!value) {
      throw new Error(message || `Expected true but got ${value}`);
    }
  },
  
  /**
   * 値が等しいことを確認
   */
  equals: function(actual, expected, message) {
    if (actual !== expected) {
      throw new Error(message || `Expected ${expected} but got ${actual}`);
    }
  },
  
  /**
   * 値が存在することを確認
   */
  exists: function(value, message) {
    if (value === null || value === undefined) {
      throw new Error(message || `Expected value to exist but got ${value}`);
    }
  },
  
  /**
   * 配列に要素が含まれることを確認
   */
  contains: function(array, item, message) {
    if (!array.includes(item)) {
      throw new Error(message || `Expected array to contain ${item}`);
    }
  },
  
  /**
   * エラーがスローされることを確認
   */
  throws: function(func, message) {
    let thrown = false;
    try {
      func();
    } catch (e) {
      thrown = true;
    }
    if (!thrown) {
      throw new Error(message || 'Expected function to throw an error');
    }
  }
};

/**
 * テスト結果を記録
 */
function recordTestResult(testName, passed, duration, error = null, details = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEST_CONFIG.SHEETS.TEST_RESULTS);
  if (!sheet) return;
  
  const timestamp = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  const result = passed ? '✅ PASS' : '❌ FAIL';
  const errorMsg = error ? error.toString() : '';
  
  sheet.appendRow([timestamp, testName, result, duration, errorMsg, details]);
}

// ===== テストデータ生成 =====

/**
 * テスト用のサンプルメッセージを生成
 */
function generateTestMessages() {
  return [
    'こんにちは',
    'Geminiとは何ですか？',
    '天気はどうですか？',
    '1+1は？',
    'JavaScriptのコードを書いて',
    ''  // 空メッセージテスト
  ];
}

/**
 * テスト用のユーザーIDを生成
 */
function generateTestUserIds() {
  return [
    'test_user_001',
    'test_user_002',
    'user_' + Date.now(),
    'special_chars_!@#$%',
    ''  // 空ID テスト
  ];
}

// ===== 単体テスト =====

/**
 * APIキー取得のテスト
 */
function testGetApiKey() {
  const startTime = Date.now();
  const testName = 'APIキー取得テスト';
  
  try {
    console.log(`🧪 ${testName} 実行中...`);
    
    // APIキーが取得できることを確認
    const apiKey = getApiKey();
    assert.exists(apiKey, 'APIキーが存在しません');
    assert.isTrue(apiKey.length > 0, 'APIキーが空です');
    
    const duration = Date.now() - startTime;
    console.log(`✅ ${testName} 成功 (${duration}ms)`);
    recordTestResult(testName, true, duration);
    return true;
    
  } catch (error) {
    const duration = Date.now() - startTime;
    console.error(`❌ ${testName} 失敗:`, error);
    recordTestResult(testName, false, duration, error);
    return false;
  }
}

/**
 * プロンプト取得のテスト
 */
function testGetSystemPrompt() {
  const startTime = Date.now();
  const testName = 'システムプロンプト取得テスト';
  
  try {
    console.log(`🧪 ${testName} 実行中...`);
    
    const prompt = getSystemPrompt();
    assert.exists(prompt, 'プロンプトが存在しません');
    assert.isTrue(prompt.length > 0, 'プロンプトが空です');
    
    const duration = Date.now() - startTime;
    console.log(`✅ ${testName} 成功 (${duration}ms)`);
    recordTestResult(testName, true, duration);
    return true;
    
  } catch (error) {
    const duration = Date.now() - startTime;
    console.error(`❌ ${testName} 失敗:`, error);
    recordTestResult(testName, false, duration, error);
    return false;
  }
}

/**
 * ログ記録のテスト
 */
function testLogging() {
  const startTime = Date.now();
  const testName = 'ログ記録テスト';
  
  try {
    console.log(`🧪 ${testName} 実行中...`);
    
    // テストデータでログを記録
    const testUserId = 'test_log_' + Date.now();
    const testMessage = 'これはテストメッセージです';
    
    // メインコードのlogChat関数を使用
    logChat(testUserId, 'user', testMessage);
    
    // ログが記録されたか確認
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEST_CONFIG.SHEETS.LOG);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) { // ヘッダー行を考慮
      const lastLog = sheet.getRange(lastRow, 1, 1, 5).getValues()[0];
      
      assert.equals(lastLog[1], testUserId, 'ユーザーIDが一致しません');
      assert.equals(lastLog[2], 'user', '役割が一致しません');
      assert.equals(lastLog[3], testMessage, 'メッセージが一致しません');
    } else {
      throw new Error('ログが記録されていません');
    }
    
    const duration = Date.now() - startTime;
    console.log(`✅ ${testName} 成功 (${duration}ms)`);
    recordTestResult(testName, true, duration);
    return true;
    
  } catch (error) {
    const duration = Date.now() - startTime;
    console.error(`❌ ${testName} 失敗:`, error);
    recordTestResult(testName, false, duration, error);
    return false;
  }
}

/**
 * キャッシュ機能のテスト
 */
function testCaching() {
  const startTime = Date.now();
  const testName = 'キャッシュ機能テスト';
  
  try {
    console.log(`🧪 ${testName} 実行中...`);
    
    const testUserId = 'test_cache_' + Date.now();
    
    // メインコードの形式に合わせた履歴構造
    const testHistory = {};
    testHistory[testUserId] = [
      { role: 'user', parts: [{ text: 'テスト1' }] },
      { role: 'model', parts: [{ text: 'レスポンス1' }] }
    ];
    
    // キャッシュに保存
    saveConversationHistory(testHistory);
    
    // キャッシュから取得
    const retrieved = getConversationHistory();
    assert.exists(retrieved[testUserId], 'ユーザー履歴が存在しません');
    assert.equals(retrieved[testUserId].length, 2, '履歴の長さが一致しません');
    assert.equals(retrieved[testUserId][0].parts[0].text, 'テスト1', 'メッセージ内容が一致しません');
    
    // キャッシュクリアのテスト
    clearConversationHistory(testUserId);
    const clearedHistory = getConversationHistory();
    assert.isTrue(!clearedHistory[testUserId] || clearedHistory[testUserId].length === 0, 'キャッシュがクリアされていません');
    
    const duration = Date.now() - startTime;
    console.log(`✅ ${testName} 成功 (${duration}ms)`);
    recordTestResult(testName, true, duration);
    return true;
    
  } catch (error) {
    const duration = Date.now() - startTime;
    console.error(`❌ ${testName} 失敗:`, error);
    recordTestResult(testName, false, duration, error);
    return false;
  }
}

// ===== 統合テスト =====

/**
 * メッセージ処理の統合テスト
 */
function testProcessMessage() {
  const startTime = Date.now();
  const testName = 'メッセージ処理統合テスト';
  
  try {
    console.log(`🧪 ${testName} 実行中...`);
    
    const testUserId = 'test_process_' + Date.now();
    const testMessage = 'こんにちは、テストです';
    
    // メッセージを処理
    const response = processMessage(testUserId, testMessage);
    
    // レスポンスの検証
    assert.exists(response, 'レスポンスが存在しません');
    assert.isTrue(response.length > 0, 'レスポンスが空です');
    
    // ログが記録されているか確認
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEST_CONFIG.SHEETS.LOG);
    const lastRow = sheet.getLastRow();
    assert.isTrue(lastRow > 1, 'ログが記録されていません');
    
    const duration = Date.now() - startTime;
    console.log(`✅ ${testName} 成功 (${duration}ms)`);
    console.log(`📝 レスポンス: ${response.substring(0, 50)}...`);
    recordTestResult(testName, true, duration, null, response.substring(0, 100));
    return true;
    
  } catch (error) {
    const duration = Date.now() - startTime;
    console.error(`❌ ${testName} 失敗:`, error);
    recordTestResult(testName, false, duration, error);
    return false;
  }
}

// ===== テストランナー =====

/**
 * すべてのテストを実行
 */
function runAllTests() {
  console.log('🚀 テストスイート実行開始\n');
  console.log('=' .repeat(50));
  
  const tests = [
    { name: '環境セットアップ', func: setupTestEnvironment },
    { name: 'APIキー取得', func: testGetApiKey },
    { name: 'プロンプト取得', func: testGetSystemPrompt },
    { name: 'ログ記録', func: testLogging },
    { name: 'キャッシュ機能', func: testCaching },
    { name: 'メッセージ処理', func: testProcessMessage }
  ];
  
  let passed = 0;
  let failed = 0;
  const startTime = Date.now();
  
  tests.forEach((test, index) => {
    console.log(`\n[${index + 1}/${tests.length}] ${test.name}`);
    const result = test.func();
    if (result) {
      passed++;
    } else {
      failed++;
    }
  });
  
  const totalDuration = Date.now() - startTime;
  
  console.log('\n' + '=' .repeat(50));
  console.log('📊 テスト結果サマリー:');
  console.log(`✅ 成功: ${passed}`);
  console.log(`❌ 失敗: ${failed}`);
  console.log(`⏱️ 総実行時間: ${totalDuration}ms`);
  console.log('=' .repeat(50));
  
  // サマリーをシートに記録
  recordTestResult('テストスイート全体', failed === 0, totalDuration, null, 
    `成功: ${passed}, 失敗: ${failed}`);
  
  return failed === 0;
}

/**
 * 負荷テスト
 */
function loadTest() {
  console.log('🏋️ 負荷テスト開始...');
  
  const numberOfUsers = 5;
  const messagesPerUser = 3;
  const results = [];
  
  for (let i = 0; i < numberOfUsers; i++) {
    const userId = `load_test_user_${i}`;
    console.log(`👤 ユーザー ${i + 1}/${numberOfUsers} テスト中...`);
    
    for (let j = 0; j < messagesPerUser; j++) {
      const startTime = Date.now();
      try {
        const response = processMessage(userId, `テストメッセージ ${j + 1}`);
        const duration = Date.now() - startTime;
        results.push({ success: true, duration: duration });
        console.log(`  ✓ メッセージ ${j + 1}: ${duration}ms`);
      } catch (error) {
        const duration = Date.now() - startTime;
        results.push({ success: false, duration: duration });
        console.log(`  ✗ メッセージ ${j + 1}: エラー (${duration}ms)`);
      }
    }
  }
  
  // 統計を計算
  const successCount = results.filter(r => r.success).length;
  const totalDuration = results.reduce((sum, r) => sum + r.duration, 0);
  const avgDuration = Math.round(totalDuration / results.length);
  
  console.log('\n📈 負荷テスト結果:');
  console.log(`  成功率: ${(successCount / results.length * 100).toFixed(1)}%`);
  console.log(`  平均応答時間: ${avgDuration}ms`);
  console.log(`  総リクエスト数: ${results.length}`);
  
  recordTestResult('負荷テスト', successCount === results.length, avgDuration, null,
    `成功率: ${(successCount / results.length * 100).toFixed(1)}%, 平均: ${avgDuration}ms`);
}