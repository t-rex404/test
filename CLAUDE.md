# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

このファイルは、Claude Code (claude.ai/code) がこのリポジトリでコードを扱う際のガイダンスを提供します。
PowerShellのバージョンは5.1を想定しています。

## プロジェクト概要

様々なアプリケーションとシステムの自動化を行うドライバークラスを提供するPowerShell自動化テストフレームワークです。PowerShellを使用したUI自動化、Web自動化、アプリケーションテストに焦点を当てています。

## コマンド

### テスト実行
```cmd
# デフォルトテスト実行（run.cmdで現在アクティブな行を実行、デフォルトはaaa.ps1）
.\run.cmd

# 特定のドライバーテスト実行（run.cmd内で希望する行のコメントアウトを外す）
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_EdgeDriver_AllWebDriverMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_WordDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_ExcelDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_PowerPointDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_UIAutomationDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_UIAutomationDriver_Calculator_Notepad.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_GUIDriver_Calculator_Notepad.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_OracleDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_TeraTermDriver_AllMethods.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_WinSCPDriver_AllMethods.ps1"

# 単体テスト実行例
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_Edge_Browser_Simple.ps1"
powershell.exe -ExecutionPolicy Bypass -file "_ps\test_EdgeWordOracle_Sample.ps1"
```

## アーキテクチャ

### ドライバーパターン
フレームワークは`_ps/_lib/`内のベースクラスを使用したモジュラードライバーパターンに従います：
- **Common.ps1** - 共有ユーティリティ、エラーハンドリング、ログ記録
- **UIAutomationDriver.ps1** - デスクトップアプリケーション用Windows UI自動化
- **EdgeDriver.ps1** / **ChromeDriver.ps1** - DevTools Protocol経由のブラウザ自動化
- **WebDriver.ps1** - WebDriver基底クラス、CDP (Chrome DevTools Protocol) 実装
- **ExcelDriver.ps1**, **WordDriver.ps1**, **PowerPointDriver.ps1** - Office自動化
- **AccessDriver.ps1** - Microsoft Access データベース自動化
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

## ディレクトリ構造
```
.
├── _ps/                      # PowerShellスクリプト
│   ├── _lib/                 # ドライバークラスライブラリ
│   └── test_*.ps1            # テストスクリプト
├── _json/                    # 設定ファイル
│   └── ErrorCode.json        # エラーコード定義
├── _docs/                    # ドキュメント
└── run.cmd                   # メインテストランナー
```

## 重要な実装詳細

### WebDriver CDP実装
- Chrome DevTools Protocolを使用したブラウザ自動化
- WebSocketを使用した双方向通信
- タブ管理、DOM操作、JavaScript実行機能

### UI Automation実装
- Windows UI Automation APIを使用
- AutomationElementによる要素検索
- パターン（InvokePattern、ValuePattern等）を使用した操作

### テスト結果レポート
- Record-TestResult関数による統一的なテスト記録
- コンソール出力とログファイルへの記録
- テストサマリーの自動生成