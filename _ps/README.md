# PowerShell WebDriver/WordDriver テストスイート

このディレクトリには、`_lib`フォルダ内のPowerShellライブラリの各メソッドをテストするためのスクリプトが含まれています。

## ファイル構成

### テストファイル
- `test_WebDriver.ps1` - WebDriverクラスの基本機能テスト
- `test_ChromeDriver.ps1` - ChromeDriverクラスの機能テスト
- `test_EdgeDriver.ps1` - EdgeDriverクラスの機能テスト
- `test_WordDriver.ps1` - WordDriverクラスの機能テスト
- `test_all.ps1` - すべてのテストを統合実行

### ドキュメント
- `bug_report.md` - 発見されたバグと問題点の詳細レポート
- `README.md` - このファイル

## 前提条件

### システム要件
- Windows 10/11
- PowerShell 5.1以上
- Microsoft Edge または Google Chrome（ブラウザテスト用）
- Microsoft Word（WordDriverテスト用）

### 必要なソフトウェア
1. **Microsoft Edge** - EdgeDriverテスト用
2. **Google Chrome** - ChromeDriverテスト用
3. **Microsoft Word** - WordDriverテスト用

## 使用方法

### 1. 個別テストの実行

#### WebDriver基本テスト
```powershell
.\test_WebDriver.ps1
```

#### ChromeDriverテスト
```powershell
.\test_ChromeDriver.ps1
```

#### EdgeDriverテスト
```powershell
.\test_EdgeDriver.ps1
```

#### WordDriverテスト
```powershell
.\test_WordDriver.ps1
```

### 2. 統合テストの実行

すべてのテストを一度に実行する場合：
```powershell
.\test_all.ps1
```

### 3. 実行権限の設定

初回実行時は、PowerShellの実行ポリシーを変更する必要があります：

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## テスト内容

### WebDriverテスト
- ブラウザの初期化
- ページナビゲーション
- 要素検索（CSS、XPath）
- 要素操作（クリック、テキスト入力）
- フォーム操作（セレクトボックス、チェックボックス、ラジオボタン）
- スクリーンショット取得
- JavaScript実行
- クッキー・ローカルストレージ操作
- ウィンドウ操作
- ナビゲーション履歴操作

### ChromeDriverテスト
- Chrome固有の初期化処理
- Chrome実行ファイルパス取得
- ユーザーデータディレクトリ管理
- WebDriverの全機能テスト

### EdgeDriverテスト
- Edge固有の初期化処理
- Edge実行ファイルパス取得
- ユーザーデータディレクトリ管理
- WebDriverの全機能テスト

### WordDriverテスト
- Wordアプリケーションの初期化
- ドキュメント作成・編集
- テキスト・見出し・段落の追加
- テーブル作成
- フォント設定
- ページ区切り
- 目次の作成・更新
- ドキュメントの保存・読み込み

## 出力ファイル

### テスト実行時に生成されるファイル
- `test_page.html` - WebDriverテスト用HTMLファイル
- `test_chrome_page.html` - ChromeDriverテスト用HTMLファイル
- `test_edge_page.html` - EdgeDriverテスト用HTMLファイル
- `test_screenshot.png` - スクリーンショット
- `element_screenshot.png` - 要素スクリーンショット
- `test_document.docx` - WordDriverテスト用ドキュメント
- `test_document_final.docx` - 最終Wordドキュメント

### ログファイル
- `WebDriver_Error.log` - WebDriverエラーログ
- `test_results_YYYYMMDD_HHMMSS.json` - テスト結果JSONファイル

## トラブルシューティング

### よくある問題

#### 1. ブラウザが見つからない
**症状**: "Chrome実行ファイルが見つかりません" または "Edge実行ファイルが見つかりません"
**解決策**: 
- ブラウザがインストールされているか確認
- レジストリパスを確認
- 手動でブラウザパスを指定

#### 2. Wordアプリケーションエラー
**症状**: "Microsoft.Office.Interop.Wordアセンブリが見つかりません"
**解決策**:
- Microsoft Wordがインストールされているか確認
- Office 365またはMicrosoft Office 2016以降を使用

#### 3. 実行ポリシーエラー
**症状**: "スクリプトの実行が無効になっています"
**解決策**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### 4. WebSocket接続エラー
**症状**: "WebSocket接続に失敗しました"
**解決策**:
- ファイアウォール設定を確認
- ポート9222が使用可能か確認
- ブラウザのバージョンを確認

### デバッグ方法

#### 詳細ログの有効化
```powershell
$VerbosePreference = "Continue"
.\test_WebDriver.ps1
```

#### エラーログの確認
```powershell
Get-Content .\WebDriver_Error.log
```

## 既知の問題

詳細は `bug_report.md` を参照してください。

### 主要な問題
1. **FindElementsメソッドの論理エラー** - 戻り値処理に問題があります
2. **未実装メソッド** - 一部の要素検索メソッドが実装されていません
3. **Disposeメソッドの呼び出しエラー** - 親クラスのメソッド呼び出しに問題があります
4. **エラーモジュールの依存関係** - 一部のエラーモジュールが存在しません

## 貢献

バグの報告や改善提案がある場合は、以下の手順でお願いします：

1. 問題を詳細に記述
2. 再現手順を提供
3. 期待される動作を説明
4. 環境情報（OS、PowerShellバージョン、ブラウザバージョン）を提供

## ライセンス

このテストスイートは、元のライブラリと同じライセンスに従います。

## 更新履歴

- 2024年 - 初回作成
  - 各ドライバークラスのテストファイル作成
  - 統合テストファイル作成
  - バグレポート作成
  - README作成 