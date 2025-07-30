# _libディレクトリ内のクラスメソッドテストプログラム

## 概要

このディレクトリには、`_lib`フォルダ内のすべてのクラスのメソッドをテストするためのPowerShellスクリプトが含まれています。

## テスト対象クラス

以下のクラスとそのメソッドがテスト対象です：

### 1. Commonクラス (`Common.ps1`)
- **WriteLog**: ログ出力機能
- **HandleError**: エラーハンドリング機能

### 2. WebDriverクラス (`WebDriver.ps1`)
- **StartBrowser**: ブラウザ起動
- **GetTabInfomation**: タブ情報取得
- **Navigate**: ページ遷移
- **GetUrl**: URL取得
- **GetTitle**: タイトル取得
- **Dispose**: リソース解放
- その他多数のメソッド（約50個以上）

### 3. ChromeDriverクラス (`ChromeDriver.ps1`)
- **GetChromeExecutablePath**: Chrome実行ファイルパス取得
- **GetUserDataDirectory**: ユーザーデータディレクトリ取得
- **EnableDebugMode**: デバッグモード有効化
- **CleanupOnInitializationFailure**: 初期化失敗時のクリーンアップ

### 4. EdgeDriverクラス (`EdgeDriver.ps1`)
- **GetEdgeExecutablePath**: Edge実行ファイルパス取得
- **GetUserDataDirectory**: ユーザーデータディレクトリ取得
- **EnableDebugMode**: デバッグモード有効化
- **CleanupOnInitializationFailure**: 初期化失敗時のクリーンアップ

### 5. WordDriverクラス (`WordDriver.ps1`)
- **CreateTempDirectory**: 一時ディレクトリ作成
- **AddText**: テキスト追加
- **AddHeading**: 見出し追加
- **AddParagraph**: 段落追加
- **AddTable**: テーブル追加
- **SetFont**: フォント設定
- **SaveDocument**: ドキュメント保存
- その他多数のメソッド

### 6. エラーモジュール
- **WebDriverErrors.psm1**
- **ChromeDriverErrors.psm1**
- **EdgeDriverErrors.psm1**
- **WordDriverErrors.psm1**

## テストプログラムファイル

### 1. `test_all_methods.ps1`
基本的なメソッドテストプログラム
- クラスのインポートテスト
- メソッドの存在確認
- 基本的な機能テスト
- エラーハンドリングテスト

### 2. `test_methods_detailed.ps1`
詳細なメソッドテストプログラム
- プロパティアクセステスト
- メソッドシグネチャテスト
- パラメータテスト
- 戻り値テスト
- 実際のメソッド呼び出しテスト

### 3. `run_all_tests.ps1`
すべてのテストプログラムを実行するラッパー
- 基本的なテストと詳細テストを順次実行
- 実行時間の計測
- 結果ファイルの確認

## 使用方法

### 1. すべてのテストを実行
```powershell
.\run_all_tests.ps1
```

### 2. 基本的なテストのみ実行
```powershell
.\test_all_methods.ps1
```

### 3. 詳細テストのみ実行
```powershell
.\test_methods_detailed.ps1
```

## 出力ファイル

テスト実行後、以下のCSVファイルが生成されます：

### 基本的なテスト結果
- `test_results_YYYYMMDD_HHMMSS.csv`
  - クラス名
  - メソッド名
  - ステータス（PASS/FAIL）
  - メッセージ
  - エラー詳細
  - タイムスタンプ

### 詳細テスト結果
- `detailed_test_results_YYYYMMDD_HHMMSS.csv`
  - クラス名
  - メソッド名
  - テストタイプ
  - ステータス（PASS/FAIL）
  - メッセージ
  - エラー詳細
  - パラメータ
  - 戻り値
  - タイムスタンプ

## テストの特徴

### 1. 安全性
- 実際のブラウザやWordアプリケーションは起動しません
- ダミーパスやモックデータを使用してテストを実行
- システムに影響を与えない安全なテスト

### 2. 包括性
- すべてのクラスのすべてのメソッドをテスト
- プロパティアクセステスト
- メソッドシグネチャテスト
- エラーハンドリングテスト

### 3. 詳細なレポート
- カラー付きのコンソール出力
- CSV形式での詳細な結果保存
- 成功/失敗の統計情報
- エラー詳細の記録

## 注意事項

### 1. 前提条件
- PowerShell 5.1以上が必要
- `_lib`ディレクトリが存在すること
- 必要なPowerShellモジュールが利用可能であること

### 2. 期待されるエラー
以下のエラーは正常な動作の一部です：
- Chrome/Edgeがインストールされていない場合のエラー
- Wordがインストールされていない場合のエラー
- ダミーパスを使用したテストでのエラー

### 3. 実行環境
- Windows環境での実行を想定
- 管理者権限は不要
- インターネット接続は不要

## トラブルシューティング

### 1. インポートエラーが発生する場合
- `_lib`ディレクトリのパスを確認
- PowerShellの実行ポリシーを確認
- 必要なモジュールが存在することを確認

### 2. テストが失敗する場合
- エラーメッセージを確認
- CSVファイルの詳細を確認
- 期待されるエラーかどうかを判断

### 3. ファイルが生成されない場合
- 書き込み権限を確認
- ディスク容量を確認
- アンチウイルスソフトの設定を確認

## カスタマイズ

### 1. テストケースの追加
各テストファイル内の配列に新しいテストケースを追加できます。

### 2. 出力形式の変更
CSVファイルの代わりにJSONやXML形式での出力に変更できます。

### 3. テスト対象の絞り込み
特定のクラスやメソッドのみをテスト対象にすることができます。

## 更新履歴

- 2024年: 初版作成
  - 基本的なテストプログラム
  - 詳細テストプログラム
  - 実行ラッパー
  - READMEファイル 