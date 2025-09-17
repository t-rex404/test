# CLAUDE.md

このファイルは、Claude Code (claude.ai/code) がこのリポジトリでコードを扱う際のガイダンスを提供します。

## プロジェクト概要

これは、様々なアプリケーションとシステムの自動化を行うドライバークラスを提供するPowerShell自動化テストフレームワークです。PowerShellを使用したUI自動化、Web自動化、アプリケーションテストに焦点を当てています。

## プロジェクト構造

- `_ps/` - PowerShellスクリプトとテストファイル
  - `_lib/` - コアドライバークラスと共通ユーティリティ
  - `test_*.ps1` - 各ドライバー用の個別テストスクリプト
- `_docs/` - HTMLドキュメントと参考資料
- `_json/` - エラーコードを含む設定ファイル
- `run.cmd` - メイン実行スクリプト
- `sample.html` - Web自動化テスト用のテストHTMLファイル

## コアアーキテクチャ

### ドライバークラス（`_ps/_lib/`内）

フレームワークは以下の主要コンポーネントによるモジュラードライバーパターンに従います：

- **Common.ps1** - 全ドライバー間で共有されるユーティリティとエラーハンドリング
- **UIAutomationDriver.ps1** - デスクトップアプリケーション用のWindows UI自動化
- **WebDriver.ps1** - Chrome DevTools Protocolを使用したブラウザ自動化
- **EdgeDriver.ps1** - Microsoft Edgeブラウザ自動化
- **ExcelDriver.ps1** - Microsoft Excel自動化
- **WordDriver.ps1** - Microsoft Word自動化
- **PowerPointDriver.ps1** - Microsoft PowerPoint自動化
- **OracleDriver.ps1** - Oracleデータベース操作
- **TeraTermDriver.ps1** - ターミナル自動化
- **WinSCPDriver.ps1** - ファイル転送自動化
- **GUIDriver.ps1** - 汎用GUI自動化

### エラーハンドリングシステム

- `_json/ErrorCode.json`での集中化されたエラーコード
- Commonクラスを通じた構造化されたエラーハンドリング
- モジュール固有のエラーログ

## テスト実行

メインエントリーポイントは`run.cmd`で、PowerShellテストスクリプトを実行します：

```cmd
# 全テスト実行
.\run.cmd

# 特定のドライバーテスト実行（run.cmd内で希望する行のコメントアウトを外す）
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_UIAutomationDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_EdgeDriver_AllWebDriverMethods.ps1"
```

## 共通開発パターン

### ドライバー初期化
全てのドライバーは以下のパターンに従います：
1. 必要なアセンブリと依存関係の読み込み
2. エラーハンドリング用のCommonクラスでの初期化
3. 接続/起動メソッドの提供
4. クリーンアップ/破棄メソッドの実装

### テスト構造
テストスクリプトは以下のパターンに従います：
1. Common.ps1と特定のドライバーの読み込み
2. グローバルCommonインスタンスの初期化
3. エラー記録付きのテストメソッド実行
4. テストレポートの生成

### エラーハンドリング
一貫したエラーハンドリングにはCommonクラスを使用します：
```powershell
$global:Common.HandleError("ErrorCode", "Error message", "ModuleName", "LogFile")
```

## 主要な依存関係

- Windows PowerShell 5.0以上
- .NET Frameworkアセンブリ：
  - System.Windows.Forms
  - System.Drawing
  - UIAutomationClient
  - UIAutomationTypes
- Web自動化用のChrome/Edgeブラウザ
- Office自動化用のMicrosoft Officeアプリケーション