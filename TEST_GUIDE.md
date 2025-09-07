# テスト基盤 - 実行ガイド

## 🎯 概要
Google Apps Script (GAS) とスプレッドシートを使用したテスト基盤です。
単体テスト、統合テスト、負荷テストをサポートします。

## 📋 必要なシート構成

| シート名 | 用途 | 自動作成 |
|---------|------|----------|
| プロンプト | システムプロンプト保存（A1セル） | ✅ |
| ログ | チャット履歴の記録 | ✅ |
| テスト結果 | テスト実行結果の記録 | ✅ |

## 🚀 セットアップ手順

### 1. ファイルのアップロード
```javascript
// GASエディタで以下のファイルを作成
1. test-framework.js の内容をコピー
2. minimal-コード.js の内容を同じプロジェクトに配置
```

### 2. APIキーの設定
```javascript
// プロジェクトの設定 → スクリプトプロパティ
GEMINI_API_KEY: [your-api-key]
```

### 3. 初期セットアップ実行
```javascript
// GASエディタで実行
setupTestEnvironment()
```

## 🧪 テスト実行方法

### 全テスト実行
```javascript
runAllTests()
```

実行されるテスト:
- ✅ 環境セットアップ
- ✅ APIキー取得テスト
- ✅ システムプロンプト取得テスト
- ✅ ログ記録テスト
- ✅ キャッシュ機能テスト
- ✅ メッセージ処理統合テスト

### 個別テスト実行
```javascript
// 各テストを個別に実行可能
testGetApiKey()        // APIキー取得
testGetSystemPrompt()  // プロンプト取得
testLogging()          // ログ記録
testCaching()          // キャッシュ
testProcessMessage()   // メッセージ処理
```

### 負荷テスト
```javascript
loadTest()
```
- 5ユーザー × 3メッセージ = 15リクエスト
- 成功率と平均応答時間を計測

## 📊 テスト結果の確認

### スプレッドシートで確認
「テスト結果」シートに自動記録:
- 実行日時
- テスト名
- 結果（✅ PASS / ❌ FAIL）
- 実行時間
- エラー詳細

### コンソールで確認
```
🚀 テストスイート実行開始
==================================================
[1/6] 環境セットアップ
✅ 環境セットアップ 成功 (523ms)

[2/6] APIキー取得
✅ APIキー取得テスト 成功 (145ms)

...

==================================================
📊 テスト結果サマリー:
✅ 成功: 6
❌ 失敗: 0
⏱️ 総実行時間: 3420ms
==================================================
```

## 🛠️ アサーション関数

```javascript
// 利用可能なアサーション
assert.isTrue(value, message)        // 真値チェック
assert.equals(actual, expected, msg) // 等価チェック
assert.exists(value, message)        // 存在チェック
assert.contains(array, item, msg)    // 配列包含チェック
assert.throws(func, message)         // 例外スローチェック
```

## 📝 テストデータ生成

```javascript
// サンプルメッセージ
generateTestMessages()
// => ['こんにちは', 'Geminiとは何ですか？', ...]

// テストユーザーID
generateTestUserIds()
// => ['test_user_001', 'test_user_002', ...]
```

## ⚠️ 注意事項

1. **APIレート制限**: Gemini APIのレート制限に注意
2. **実行時間制限**: GASの6分制限内で実行
3. **キャッシュ**: テスト間でキャッシュが共有される
4. **ログ蓄積**: ログシートが大きくなりすぎないよう定期的にクリア

## 🔍 トラブルシューティング

### APIキーエラー
```javascript
// スクリプトプロパティを確認
PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
```

### シートが見つからない
```javascript
// 手動で再セットアップ
setupTestEnvironment()
```

### タイムアウトエラー
```javascript
// タイムアウト設定を調整
TEST_CONFIG.TEST_TIMEOUT = 20000  // 20秒に延長
```

## 📈 パフォーマンス目標

| メトリクス | 目標値 | 現状 |
|-----------|--------|------|
| 単体テスト実行時間 | < 500ms | ✅ |
| API応答時間 | < 3秒 | ✅ |
| テスト成功率 | 100% | ✅ |
| 負荷テスト成功率 | > 95% | ✅ |