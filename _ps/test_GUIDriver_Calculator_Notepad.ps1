# GUIDriverテストスクリプト
# 電卓とメモ帳を操作するテスト

# 必要なライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"
. "$PSScriptRoot\_lib\GUIDriver.ps1"

# テスト用の一時ディレクトリを作成
$testDir = ".\GUIDriver_Test_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $testDir -Force | Out-Null

Write-Host "=== GUIDriverテスト開始 ===" -ForegroundColor Cyan
Write-Host "テストディレクトリ: $testDir" -ForegroundColor Yellow

try
{
    # ========================================
    # 電卓テスト
    # ========================================
    Write-Host "`n--- 電卓テスト開始 ---" -ForegroundColor Green
    
    $calcDriver = [GUIDriver]::new()
    
    # 電卓を起動
    $calcPath = "calc.exe"
    $calcDriver.StartApplication($calcPath)
    Start-Sleep -Seconds 2
    
    # 電卓ウィンドウを検索
    $calcDriver.FindWindow("電卓")
    $calcDriver.ActivateWindow()
    Start-Sleep -Seconds 1
    
    # 電卓のスクリーンショットを取得
    $calcDriver.TakeWindowScreenshot("$testDir\calculator_initial.png")
    
    # 電卓操作: 2 + 3 = 5
    Write-Host "電卓操作: 2 + 3 = 5" -ForegroundColor Yellow
    
    # 2を入力
    $calcDriver.TypeText("2")
    Start-Sleep -Milliseconds 500
    
    # +を入力
    $calcDriver.TypeText("+")
    Start-Sleep -Milliseconds 500
    
    # 3を入力
    $calcDriver.TypeText("3")
    Start-Sleep -Milliseconds 500
    
    # =を入力
    $calcDriver.TypeText("=")
    Start-Sleep -Milliseconds 500
    
    # 結果のスクリーンショットを取得
    $calcDriver.TakeWindowScreenshot("$testDir\calculator_result.png")
    
    # 電卓を閉じる
    $calcDriver.CloseApplication()
    Start-Sleep -Seconds 1
    
    Write-Host "電卓テスト完了" -ForegroundColor Green
    
    # ========================================
    # メモ帳テスト
    # ========================================
    Write-Host "`n--- メモ帳テスト開始 ---" -ForegroundColor Green
    
    $notepadDriver = [GUIDriver]::new()
    
    # メモ帳を起動
    $notepadPath = "notepad.exe"
    $notepadDriver.StartApplication($notepadPath)
    Start-Sleep -Seconds 2
    
    # メモ帳ウィンドウを検索
    $notepadDriver.FindWindow("メモ帳")
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
    # マウス操作テスト（電卓で実行）
    # ========================================
    Write-Host "`n--- マウス操作テスト開始 ---" -ForegroundColor Green
    
    $mouseDriver = [GUIDriver]::new()
    
    # 電卓を再起動
    $mouseDriver.StartApplication($calcPath)
    Start-Sleep -Seconds 2
    
    # 電卓ウィンドウを検索
    $mouseDriver.FindWindow("電卓")
    $mouseDriver.ActivateWindow()
    Start-Sleep -Seconds 1
    
    # マウス操作テスト: 電卓のボタンをクリック
    Write-Host "マウス操作テスト: 電卓のボタンをクリック" -ForegroundColor Yellow
    
    # 電卓の中央付近をクリック（数値ボタンエリア）
    $mouseDriver.ClickMouse(300, 300)
    Start-Sleep -Milliseconds 500
    
    # 右クリックテスト
    $mouseDriver.RightClickMouse(400, 400)
    Start-Sleep -Milliseconds 500
    
    # マウス移動テスト
    $mouseDriver.MoveMouse(200, 200)
    Start-Sleep -Milliseconds 500
    
    # マウス操作後のスクリーンショットを取得
    $mouseDriver.TakeWindowScreenshot("$testDir\calculator_mouse_test.png")
    
    # 電卓を閉じる
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
    if ($calcDriver) { $calcDriver.Dispose() }
    if ($notepadDriver) { $notepadDriver.Dispose() }
    if ($mouseDriver) { $mouseDriver.Dispose() }
    
    Write-Host "`n=== GUIDriverテスト終了 ===" -ForegroundColor Cyan
}
