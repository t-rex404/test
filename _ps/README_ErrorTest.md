# エラー管理システムテスト

## 概要
このディレクトリには、各ドライバークラスのエラー管理システムをテストするためのスクリプトが含まれています。

## テストスクリプト

### 1. test.ps1
基本的なエラーテストスクリプト
- 各ドライバークラスの初期化エラーをテスト
- 無効なパラメータでエラーを発生させる

### 2. test_specific_errors.ps1
具体的なエラーテストスクリプト
- より詳細なエラーケースをテスト
- 各メソッドの特定のエラーを発生させる

### 3. manual_test.ps1
手動エラーテストスクリプト
- エラー管理モジュールを直接テスト
- エラーコードとメッセージの表示をテスト

### 4. simple_test.ps1
シンプルなエラーテストスクリプト
- 共通ライブラリの基本機能をテスト
- エラーコードの表示をテスト

## 実行方法

### PowerShell 5.1での実行
```powershell
# 基本的なテスト
powershell -ExecutionPolicy Bypass -File test.ps1

# 具体的なエラーテスト
powershell -ExecutionPolicy Bypass -File test_specific_errors.ps1

# 手動エラーテスト
powershell -ExecutionPolicy Bypass -File manual_test.ps1

# シンプルなテスト
powershell -ExecutionPolicy Bypass -File simple_test.ps1
```

### 実行ポリシーの変更が必要な場合
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

## 生成されるログファイル

テスト実行後、以下のログファイルが生成されます：

### 1. WebDriver_Error.log
- WebDriverクラスのエラーログ
- エラーコード: 1001-1063
- 例: 初期化エラー、ナビゲーションエラー、要素検索エラー

### 2. EdgeDriver_Error.log
- EdgeDriverクラスのエラーログ
- エラーコード: 2001-2005
- 例: 初期化エラー、実行ファイルパスエラー

### 3. ChromeDriver_Error.log
- ChromeDriverクラスのエラーログ
- エラーコード: 3001-3005
- 例: 初期化エラー、実行ファイルパスエラー

### 4. WordDriver_Error.log
- WordDriverクラスのエラーログ
- エラーコード: 4001-4016
- 例: 初期化エラー、ドキュメント保存エラー、画像追加エラー

### 5. Common_Error.log
- 共通エラーハンドリングのログ
- 汎用的なエラーコード
- 例: テスト用エラー、共通エラー

## ログファイルの形式

各ログファイルには以下の情報が含まれます：

```
[タイムスタンプ], ERROR_CODE:エラーコード, ERROR_MESSAGE:エラーメッセージ
詳細情報:
- エラーコード: 数値コード
- エラーメッセージ: 詳細なエラー説明
- エラータイトル: エラーの種類
- タイムスタンプ: エラー発生時刻
- PowerShellバージョン: 実行環境のバージョン
- OS: オペレーティングシステム情報
- 実行ユーザー: 実行ユーザー名
- 実行パス: スクリプト実行パス
- モジュール: エラーが発生したモジュール名
```

## サンプルログファイル

`sample_logs/` ディレクトリには、各ログファイルのサンプルが含まれています：

- `WebDriver_Error.log` - WebDriverエラーのサンプル
- `EdgeDriver_Error.log` - EdgeDriverエラーのサンプル
- `ChromeDriver_Error.log` - ChromeDriverエラーのサンプル
- `WordDriver_Error.log` - WordDriverエラーのサンプル
- `Common_Error.log` - 共通エラーのサンプル

## エラーコード体系

### WebDriver (1000番台)
- 1001: 初期化エラー
- 1002: ブラウザ起動エラー
- 1011: 要素検索エラー
- 1018: 要素テキスト取得エラー
- 1024: スクリーンショット取得エラー
- など...

### EdgeDriver (2000番台)
- 2001: 初期化エラー
- 2002: 実行ファイルパス取得エラー
- 2003: ユーザーデータディレクトリ作成エラー
- 2004: デバッグモード有効化エラー
- 2005: Disposeエラー

### ChromeDriver (3000番台)
- 3001: 初期化エラー
- 3002: 実行ファイルパス取得エラー
- 3003: ユーザーデータディレクトリ作成エラー
- 3004: デバッグモード有効化エラー
- 3005: Disposeエラー

### WordDriver (4000番台)
- 4001: 初期化エラー
- 4002: 一時ディレクトリ作成エラー
- 4003: ドキュメント保存エラー
- 4008: 画像追加エラー
- など...

## 注意事項

1. **実行環境**: PowerShell 5.1以上が必要です
2. **権限**: 一部のテストでは管理者権限が必要な場合があります
3. **ブラウザ**: EdgeDriverとChromeDriverのテストには、対応するブラウザがインストールされている必要があります
4. **Microsoft Word**: WordDriverのテストには、Microsoft Wordがインストールされている必要があります
5. **ログファイル**: テスト実行後、ログファイルが現在のディレクトリに生成されます

## トラブルシューティング

### 実行ポリシーエラー
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

### モジュールが見つからないエラー
- 各ドライバークラスファイルが `_lib/` ディレクトリに存在することを確認
- パスが正しく設定されていることを確認

### ログファイルが生成されない
- 書き込み権限があることを確認
- ディスク容量が十分にあることを確認 