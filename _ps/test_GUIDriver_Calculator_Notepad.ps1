# GUIDriverテストスクリプト
# 電卓とメモ帳を操作するテスト

# 必要なアセンブリを読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# スクリプトの基準パス
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$LibDir    = Join-Path $ScriptDir '_lib'

# 必要なライブラリをインポート
. (Join-Path $LibDir 'Common.ps1')

# GUIDriverクラスを確実に読み込む
Write-Host "GUIDriverクラスを読み込み中..." -ForegroundColor Yellow
. (Join-Path $LibDir 'GUIDriver.ps1')

# クラス定義の確認
if (-not ([System.Management.Automation.PSTypeName]'GUIDriver').Type)
{
    Write-Host "GUIDriverクラスの読み込みに失敗しました。" -ForegroundColor Red
    exit 1
}
else
{
    Write-Host "GUIDriverクラスが正常に読み込まれました。" -ForegroundColor Green
}

# テスト用の一時ディレクトリを作成
$testDir = ".\GUIDriver_Test_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $testDir -Force | Out-Null

Write-Host "=== GUIDriverテスト開始 ===" -ForegroundColor Cyan
Write-Host "テストディレクトリ: $testDir" -ForegroundColor Yellow

try
{
    # ========================================
    # 電卓テスト（Windows 10/11の電卓アプリはMicrosoft Storeアプリのため一時的に無効化）
    # ========================================
    Write-Host "`n--- 電卓テストは一時的に無効化されています ---" -ForegroundColor Yellow
    Write-Host "理由: Windows 10/11の電卓アプリはMicrosoft Storeアプリのため、従来の方法では制御が困難です" -ForegroundColor Yellow
    # 電卓テストは後で実装予定
    
    # ========================================
    # メモ帳テスト
    # ========================================
    Write-Host "`n--- メモ帳テスト開始 ---" -ForegroundColor Green
    
    $notepadDriver = [GUIDriver]::new()
    
    # メモ帳を起動
    $notepadPath = (Get-Command notepad).Source
    $notepadDriver.StartApplication($notepadPath, "")
    Start-Sleep -Seconds 5
    
    # メモ帳ウィンドウを検索
    $notepadDriver.FindWindow("タイトルなし - メモ帳")
    $notepadDriver.ActivateWindow()
    Start-Sleep -Seconds 1
    
    # メモ帳のスクリーンショットを取得
    $notepadDriver.TakeWindowScreenshot("$testDir\notepad_initial.png")
    
    # メモ帳にテキストを入力
    Write-Host "メモ帳にテキストを入力" -ForegroundColor Yellow
    
    $testText = @"
GUIDriverテスト
================

このテキストはGUIDriverクラスを使用して自動入力されました。

テスト項目:
1. アプリケーション起動
2. ウィンドウ検索
3. ウィンドウアクティブ化
4. テキスト入力
5. スクリーンショット取得

テスト日時: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')
"@
    
    $notepadDriver.TypeText($testText)
    Start-Sleep -Seconds 1
    
    # 入力後のスクリーンショットを取得
    $notepadDriver.TakeWindowScreenshot("$testDir\notepad_with_text.png")
    
    # キーボードショートカットテスト: Ctrl+A (全選択)
    Write-Host "キーボードショートカットテスト: Ctrl+A" -ForegroundColor Yellow
    $notepadDriver.SendKeyCombination(@("Ctrl", "A"))
    Start-Sleep -Milliseconds 500
    
    # 全選択後のスクリーンショットを取得
    $notepadDriver.TakeWindowScreenshot("$testDir\notepad_selected.png")
    
    # メモ帳を閉じる（保存しない）
    $notepadDriver.CloseApplication()
    Start-Sleep -Seconds 1
    
    Write-Host "メモ帳テスト完了" -ForegroundColor Green
    
    # ========================================
    # マウス操作テスト（メモ帳で実行）
    # ========================================
    Write-Host "`n--- マウス操作テスト開始 ---" -ForegroundColor Green
    
    $mouseDriver = [GUIDriver]::new()
    
    # メモ帳を再起動
    $notepadPath = (Get-Command notepad).Source
    $mouseDriver.StartApplication($notepadPath, "")
    Start-Sleep -Seconds 5
    
    # メモ帳ウィンドウを検索
    $mouseDriver.FindWindow("タイトルなし - メモ帳")
    $mouseDriver.ActivateWindow()
    Start-Sleep -Seconds 1
    
    # マウス操作テスト: メモ帳の中央付近をクリック
    Write-Host "マウス操作テスト: メモ帳の中央付近をクリック" -ForegroundColor Yellow
    
    # メモ帳の中央付近をクリック
    $mouseDriver.ClickMouse(400, 300, "Left")
    Start-Sleep -Milliseconds 500
    
    # 右クリックテスト
    $mouseDriver.RightClickMouse(500, 400)
    Start-Sleep -Milliseconds 500
    
    # マウス移動テスト
    $mouseDriver.MoveMouse(300, 200)
    Start-Sleep -Milliseconds 500
    
    # マウス操作後のスクリーンショットを取得
    $mouseDriver.TakeWindowScreenshot("$testDir\notepad_mouse_test.png")
    
    # メモ帳を閉じる
    $mouseDriver.CloseApplication()
    Start-Sleep -Seconds 1
    
    Write-Host "マウス操作テスト完了" -ForegroundColor Green
    
    # ========================================
    # テスト結果の表示
    # ========================================
    Write-Host "`n=== テスト結果 ===" -ForegroundColor Cyan
    
    # 生成されたファイル一覧を表示
    $screenshots = Get-ChildItem -Path $testDir -Filter "*.png" | Sort-Object Name
    Write-Host "生成されたスクリーンショット:" -ForegroundColor Yellow
    foreach ($screenshot in $screenshots)
    {
        Write-Host "  - $($screenshot.Name)" -ForegroundColor White
    }
    
    Write-Host "`nテストディレクトリ: $testDir" -ForegroundColor Yellow
    Write-Host "すべてのテストが正常に完了しました！" -ForegroundColor Green
}
catch
{
    Write-Host "`nテスト中にエラーが発生しました:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    if ($global:Common)
    {
        $global:Common.HandleError("CommonError_0001", "GUIDriverテストエラー: $($_.Exception.Message)", "Common", ".\AllDrivers_Error.log")
    }
}
finally
{
    # リソースのクリーンアップ
    if ($notepadDriver) { $notepadDriver.Dispose() }
    if ($mouseDriver) { $mouseDriver.Dispose() }
    
    Write-Host "`n=== GUIDriverテスト終了 ===" -ForegroundColor Cyan
}
