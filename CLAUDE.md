# CLAUDE.md

このファイルは、Claude Code (claude.ai/code) がこのリポジトリでコードを扱う際のガイダンスを提供します。

## プロジェクト概要

様々なアプリケーションとシステムの自動化を行うドライバークラスを提供するPowerShell自動化テストフレームワークです。PowerShellを使用したUI自動化、Web自動化、アプリケーションテストに焦点を当てています。

## コマンド

### テスト実行
```cmd
# デフォルトテスト実行（現在はtest_UIAutomationDriver_AllMethods.ps1）
.\run.cmd

# 特定のドライバーテスト実行（run.cmd内で希望する行のコメントアウトを外す）
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_EdgeDriver_AllWebDriverMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_WordDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_ExcelDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_PowerPointDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_UIAutomationDriver_Calculator_Notepad.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_GUIDriver_Calculator_Notepad.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_OracleDriver_AllMethods.ps1"
```

## アーキテクチャ

### ドライバーパターン
フレームワークは`_ps/_lib/`内のベースクラスを使用したモジュラードライバーパターンに従います：
- **Common.ps1** - 共有ユーティリティ、エラーハンドリング、ログ記録
- **UIAutomationDriver.ps1** - デスクトップアプリケーション用Windows UI自動化
- **EdgeDriver.ps1** / **ChromeDriver.ps1** - DevTools Protocol経由のブラウザ自動化
- **ExcelDriver.ps1**, **WordDriver.ps1**, **PowerPointDriver.ps1** - Office自動化
- **OracleDriver.ps1** - Oracleデータベース操作
- **GUIDriver.ps1** - 汎用GUI自動化
- **TeraTermDriver.ps1**, **WinSCPDriver.ps1** - ターミナルとファイル転送の自動化

### エラーハンドリングシステム
- `_json/ErrorCode.json`での集中化されたエラーコード
- Commonクラスを通じた構造化されたエラーハンドリング
- HandleErrorメソッドによるモジュール固有のエラーログ記録

### テスト構造パターン
全てのテストスクリプトは以下に従います：
1. アセンブリと依存関係の読み込み
2. Common.ps1と特定のドライバーの読み込み
3. `$global:Common = [Common]::new()`の初期化
4. Record-TestResultでテストメソッドを実行
5. テストレポートサマリーの生成

### ドライバー初期化パターン
```powershell
# 標準的な初期化
$global:Common = [Common]::new()
$driver = [DriverClass]::new()

# エラーハンドリング
$global:Common.HandleError("ERROR_CODE", "メッセージ", "モジュール名", "ログファイル.log")
```