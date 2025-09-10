# UIAutomationDriver全メソッドテストスクリプト
# UIAutomationDriverクラスの全メソッドの機能を網羅的にテストする

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# スクリプトの基準パス
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$LibDir    = Join-Path $ScriptDir '_lib'

# 必要なモジュールを読み込み
. (Join-Path $LibDir 'Common.ps1')
. (Join-Path $LibDir 'UIAutomationDriver.ps1')

# グローバル変数の初期化
$global:Common = [Common]::new()

# テスト結果を記録する変数
$TestResults = @{
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
    TestDetails = @()
}

# テスト結果を記録する関数
function Record-TestResult {
    param(
        [string]$TestName,
        [bool]$Passed,
        [string]$Message = ""
    )
    
    $TestResults.TotalTests++
    if ($Passed) {
        $TestResults.PassedTests++
        Write-Host "✓ $TestName" -ForegroundColor Green
    } else {
        $TestResults.FailedTests++
        Write-Host "✗ $TestName" -ForegroundColor Red
        if ($Message) {
            Write-Host "  エラー: $Message" -ForegroundColor Red
        }
    }
    
    $TestResults.TestDetails += @{
        TestName = $TestName
        Passed = $Passed
        Message = $Message
        Timestamp = Get-Date
    }
}

# テスト用の一時ディレクトリを作成
$TestDir = ".\UIAutomationDriver_Test_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $TestDir -Force | Out-Null
Write-Host "テスト用ディレクトリを作成しました: $TestDir" -ForegroundColor Cyan

Write-Host "=== UIAutomationDriver 全メソッドテスト開始 ===" -ForegroundColor Cyan
Write-Host "テスト開始時刻: $(Get-Date)" -ForegroundColor White

try
{
    # ========================================
    # 1. 初期化・接続関連メソッドのテスト
    # ========================================
    Write-Host "`n--- 1. 初期化・接続関連メソッドのテスト ---" -ForegroundColor Yellow

    # UIAutomationDriverインスタンス作成
    Write-Host "`n1.1 UIAutomationDriverインスタンス作成テスト" -ForegroundColor Cyan
    try {
        $guiDriver = [UIAutomationDriver]::new()
        Record-TestResult "UIAutomationDriverインスタンス作成" $true
    }
    catch {
        Record-TestResult "UIAutomationDriverインスタンス作成" $false $_.Exception.Message
        throw
    }

    # StartApplicationメソッドのテスト
    Write-Host "`n1.2 StartApplicationメソッドのテスト" -ForegroundColor Cyan
    try {
        $calcPath = "calc.exe"
        $guiDriver.StartApplication($calcPath, "")
        Start-Sleep -Seconds 2
        Record-TestResult "StartApplication - 電卓起動" $true
    }
    catch {
        Record-TestResult "StartApplication - 電卓起動" $false $_.Exception.Message
    }

    # FindWindowメソッドのテスト
    Write-Host "`n1.3 FindWindowメソッドのテスト" -ForegroundColor Cyan
    try {
        $calcWindow = $guiDriver.FindWindow("電卓")
        Record-TestResult "FindWindow - 電卓ウィンドウ検索" $true
    }
    catch {
        Record-TestResult "FindWindow - 電卓ウィンドウ検索" $false $_.Exception.Message
    }

    # FindWindowByPartialNameメソッドのテスト
    Write-Host "`n1.4 FindWindowByPartialNameメソッドのテスト" -ForegroundColor Cyan
    try {
        $partialWindow = $guiDriver.FindWindowByPartialName("電卓")
        Record-TestResult "FindWindowByPartialName - 部分一致検索" $true
    }
    catch {
        Record-TestResult "FindWindowByPartialName - 部分一致検索" $false $_.Exception.Message
    }

    # FindWindowByProcessNameメソッドのテスト
    Write-Host "`n1.5 FindWindowByProcessNameメソッドのテスト" -ForegroundColor Cyan
    try {
        $processWindow = $guiDriver.FindWindowByProcessName("CalculatorApp")
        Record-TestResult "FindWindowByProcessName - プロセス名検索" $true
    }
    catch {
        Record-TestResult "FindWindowByProcessName - プロセス名検索" $false $_.Exception.Message
    }

    # FindWindowByProcessAndTitleメソッドのテスト
    Write-Host "`n1.6 FindWindowByProcessAndTitleメソッドのテスト" -ForegroundColor Cyan
    try {
        $processAndTitleWindow = $guiDriver.FindWindowByProcessAndTitle("CalculatorApp", "電卓")
        Record-TestResult "FindWindowByProcessAndTitle - プロセス名とタイトル検索" $true
    }
    catch {
        Record-TestResult "FindWindowByProcessAndTitle - プロセス名とタイトル検索" $false $_.Exception.Message
    }

    # FindWindowFlexibleメソッドのテスト
    Write-Host "`n1.7 FindWindowFlexibleメソッドのテスト" -ForegroundColor Cyan
    try {
        $searchCriteria = @{
            ProcessName = "CalculatorApp"
            WindowTitle = "電卓"
            ExactMatch = $false
            Timeout = 5000
        }
        $flexibleWindow = $guiDriver.FindWindowFlexible($searchCriteria)
        Record-TestResult "FindWindowFlexible - 柔軟な検索" $true
    }
    catch {
        Record-TestResult "FindWindowFlexible - 柔軟な検索" $false $_.Exception.Message
    }

    # ActivateWindowメソッドのテスト
    Write-Host "`n1.8 ActivateWindowメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.ActivateWindow()
        Record-TestResult "ActivateWindow - ウィンドウアクティブ化" $true
    }
    catch {
        Record-TestResult "ActivateWindow - ウィンドウアクティブ化" $false $_.Exception.Message
    }

    # ========================================
    # 2. 要素検索・操作関連メソッドのテスト
    # ========================================
    Write-Host "`n--- 2. 要素検索・操作関連メソッドのテスト ---" -ForegroundColor Yellow

    # FindElementByNameメソッドのテスト
    Write-Host "`n2.1 FindElementByNameメソッドのテスト" -ForegroundColor Cyan
    try {
        $button7 = $guiDriver.FindElementByName("7")
        Record-TestResult "FindElementByName - 数字ボタン7検索" $true
    }
    catch {
        Record-TestResult "FindElementByName - 数字ボタン7検索" $false $_.Exception.Message
    }

    # FindElementByControlTypeメソッドのテスト
    Write-Host "`n2.2 FindElementByControlTypeメソッドのテスト" -ForegroundColor Cyan
    try {
        $buttonElement = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Button)
        Record-TestResult "FindElementByControlType - ボタン要素検索" $true
    }
    catch {
        Record-TestResult "FindElementByControlType - ボタン要素検索" $false $_.Exception.Message
    }

    # FindElementメソッドのテスト（複合条件）
    Write-Host "`n2.3 FindElementメソッドのテスト（複合条件）" -ForegroundColor Cyan
    try {
        $conditions = @{
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty = [System.Windows.Automation.ControlType]::Button
            [System.Windows.Automation.AutomationElement]::NameProperty = "7"
        }
        $element = $guiDriver.FindElement($conditions)
        Record-TestResult "FindElement - 複合条件検索" $true
    }
    catch {
        Record-TestResult "FindElement - 複合条件検索" $false $_.Exception.Message
    }

    # ClickElementメソッドのテスト
    Write-Host "`n2.4 ClickElementメソッドのテスト" -ForegroundColor Cyan
    try {
        $button7 = $guiDriver.FindElementByName("7")
        $guiDriver.ClickElement($button7)
        Start-Sleep -Milliseconds 500
        Record-TestResult "ClickElement - ボタンクリック" $true
    }
    catch {
        Record-TestResult "ClickElement - ボタンクリック" $false $_.Exception.Message
    }

    # SetElementTextメソッドのテスト（テキストボックスがある場合）
    Write-Host "`n2.5 SetElementTextメソッドのテスト" -ForegroundColor Cyan
    try {
        # メモ帳を起動してテキスト入力テスト
        $notepadPath = "notepad.exe"
        $guiDriver.StartApplication($notepadPath, "")
        Start-Sleep -Seconds 2
        
        $notepadWindow = $guiDriver.FindWindow("メモ帳")
        $guiDriver.ActivateWindow()
        
        $textArea = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Document)
        $guiDriver.SetElementText($textArea, "UIAutomationDriverテスト用テキスト")
        Record-TestResult "SetElementText - テキスト設定" $true
    }
    catch {
        Record-TestResult "SetElementText - テキスト設定" $false $_.Exception.Message
    }

    # GetElementTextメソッドのテスト
    Write-Host "`n2.6 GetElementTextメソッドのテスト" -ForegroundColor Cyan
    try {
        $textArea = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Document)
        $text = $guiDriver.GetElementText($textArea)
        Record-TestResult "GetElementText - テキスト取得" $true
    }
    catch {
        Record-TestResult "GetElementText - テキスト取得" $false $_.Exception.Message
    }

    # メモ帳を終了
    try {
        $guiDriver.CloseApplication()
        Start-Sleep -Seconds 1
    }
    catch {
        Write-Host "メモ帳の終了でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # ========================================
    # 3. マウス操作関連メソッドのテスト
    # ========================================
    Write-Host "`n--- 3. マウス操作関連メソッドのテスト ---" -ForegroundColor Yellow

    # ClickMouseメソッドのテスト
    Write-Host "`n3.1 ClickMouseメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.ClickMouse(100, 100, "Left")
        Record-TestResult "ClickMouse - 左クリック" $true
    }
    catch {
        Record-TestResult "ClickMouse - 左クリック" $false $_.Exception.Message
    }

    # RightClickMouseメソッドのテスト
    Write-Host "`n3.2 RightClickMouseメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.RightClickMouse(200, 200)
        Record-TestResult "RightClickMouse - 右クリック" $true
    }
    catch {
        Record-TestResult "RightClickMouse - 右クリック" $false $_.Exception.Message
    }

    # DoubleClickMouseメソッドのテスト
    Write-Host "`n3.3 DoubleClickMouseメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.DoubleClickMouse(300, 300)
        Record-TestResult "DoubleClickMouse - ダブルクリック" $true
    }
    catch {
        Record-TestResult "DoubleClickMouse - ダブルクリック" $false $_.Exception.Message
    }

    # MoveMouseメソッドのテスト
    Write-Host "`n3.4 MoveMouseメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.MoveMouse(400, 400)
        Record-TestResult "MoveMouse - マウス移動" $true
    }
    catch {
        Record-TestResult "MoveMouse - マウス移動" $false $_.Exception.Message
    }

    # ========================================
    # 4. キーボード操作関連メソッドのテスト
    # ========================================
    Write-Host "`n--- 4. キーボード操作関連メソッドのテスト ---" -ForegroundColor Yellow

    # SendKeysメソッドのテスト
    Write-Host "`n4.1 SendKeysメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.SendKeys("Hello World")
        Record-TestResult "SendKeys - キー送信" $true
    }
    catch {
        Record-TestResult "SendKeys - キー送信" $false $_.Exception.Message
    }

    # SendSpecialKeyメソッドのテスト
    Write-Host "`n4.2 SendSpecialKeyメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.SendSpecialKey("Enter")
        Record-TestResult "SendSpecialKey - Enterキー" $true
    }
    catch {
        Record-TestResult "SendSpecialKey - Enterキー" $false $_.Exception.Message
    }

    # SendKeyCombinationメソッドのテスト
    Write-Host "`n4.3 SendKeyCombinationメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.SendKeyCombination(@("Ctrl", "A"))
        Record-TestResult "SendKeyCombination - Ctrl+A" $true
    }
    catch {
        Record-TestResult "SendKeyCombination - Ctrl+A" $false $_.Exception.Message
    }

    # TypeTextメソッドのテスト
    Write-Host "`n4.4 TypeTextメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.TypeText("TypeTextテスト")
        Record-TestResult "TypeText - テキスト入力" $true
    }
    catch {
        Record-TestResult "TypeText - テキスト入力" $false $_.Exception.Message
    }

    # ========================================
    # 5. スクリーンショット関連メソッドのテスト
    # ========================================
    Write-Host "`n--- 5. スクリーンショット関連メソッドのテスト ---" -ForegroundColor Yellow

    # TakeScreenshotメソッドのテスト
    Write-Host "`n5.1 TakeScreenshotメソッドのテスト" -ForegroundColor Cyan
    try {
        $screenshotPath = Join-Path $TestDir "full_screenshot.png"
        $guiDriver.TakeScreenshot($screenshotPath)
        if (Test-Path $screenshotPath) {
            Record-TestResult "TakeScreenshot - 全画面スクリーンショット" $true
        } else {
            Record-TestResult "TakeScreenshot - 全画面スクリーンショット" $false "ファイルが作成されませんでした"
        }
    }
    catch {
        Record-TestResult "TakeScreenshot - 全画面スクリーンショット" $false $_.Exception.Message
    }

    # TakeWindowScreenshotメソッドのテスト
    Write-Host "`n5.2 TakeWindowScreenshotメソッドのテスト" -ForegroundColor Cyan
    try {
        $windowScreenshotPath = Join-Path $TestDir "window_screenshot.png"
        $guiDriver.TakeWindowScreenshot($windowScreenshotPath)
        if (Test-Path $windowScreenshotPath) {
            Record-TestResult "TakeWindowScreenshot - ウィンドウスクリーンショット" $true
        } else {
            Record-TestResult "TakeWindowScreenshot - ウィンドウスクリーンショット" $false "ファイルが作成されませんでした"
        }
    }
    catch {
        Record-TestResult "TakeWindowScreenshot - ウィンドウスクリーンショット" $false $_.Exception.Message
    }

    # ========================================
    # 6. プロセス管理関連メソッドのテスト
    # ========================================
    Write-Host "`n--- 6. プロセス管理関連メソッドのテスト ---" -ForegroundColor Yellow

    # IsProcessRunningメソッドのテスト
    Write-Host "`n6.1 IsProcessRunningメソッドのテスト" -ForegroundColor Cyan
    try {
        $isRunning = $guiDriver.IsProcessRunning()
        Record-TestResult "IsProcessRunning - プロセス状態確認" $true
    }
    catch {
        Record-TestResult "IsProcessRunning - プロセス状態確認" $false $_.Exception.Message
    }

    # Waitメソッドのテスト
    Write-Host "`n6.2 Waitメソッドのテスト" -ForegroundColor Cyan
    try {
        $startTime = Get-Date
        $guiDriver.Wait(1000)
        $endTime = Get-Date
        $elapsed = ($endTime - $startTime).TotalMilliseconds
        if ($elapsed -ge 900 -and $elapsed -le 1100) {
            Record-TestResult "Wait - 待機処理" $true
        } else {
            Record-TestResult "Wait - 待機処理" $false "待機時間が期待値と異なります: ${elapsed}ms"
        }
    }
    catch {
        Record-TestResult "Wait - 待機処理" $false $_.Exception.Message
    }

    # CloseApplicationメソッドのテスト
    Write-Host "`n6.3 CloseApplicationメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.CloseApplication()
        Start-Sleep -Seconds 2
        Record-TestResult "CloseApplication - アプリケーション終了" $true
    }
    catch {
        Record-TestResult "CloseApplication - アプリケーション終了" $false $_.Exception.Message
    }

    # KillApplicationメソッドのテスト（新しいプロセスで）
    Write-Host "`n6.4 KillApplicationメソッドのテスト" -ForegroundColor Cyan
    try {
        # 新しい電卓を起動
        $guiDriver.StartApplication("calc.exe", "")
        Start-Sleep -Seconds 2
        $guiDriver.FindWindow("電卓")
        $guiDriver.KillApplication()
        Start-Sleep -Seconds 1
        Record-TestResult "KillApplication - プロセス強制終了" $true
    }
    catch {
        Record-TestResult "KillApplication - プロセス強制終了" $false $_.Exception.Message
    }

    # ========================================
    # 7. ユーティリティメソッドのテスト
    # ========================================
    Write-Host "`n--- 7. ユーティリティメソッドのテスト ---" -ForegroundColor Yellow

    # Disposeメソッドのテスト
    Write-Host "`n7.1 Disposeメソッドのテスト" -ForegroundColor Cyan
    try {
        $guiDriver.Dispose()
        Record-TestResult "Dispose - リソース解放" $true
    }
    catch {
        Record-TestResult "Dispose - リソース解放" $false $_.Exception.Message
    }

    # ========================================
    # 8. エラーハンドリングのテスト
    # ========================================
    Write-Host "`n--- 8. エラーハンドリングのテスト ---" -ForegroundColor Yellow

    # 無効なアプリケーションパスでのテスト
    Write-Host "`n8.1 無効なアプリケーションパスでのテスト" -ForegroundColor Cyan
    try {
        $newDriver = [UIAutomationDriver]::new()
        $newDriver.StartApplication("invalid_app.exe", "")
        Record-TestResult "エラーハンドリング - 無効なアプリケーションパス" $false "エラーが発生すべきでした"
    }
    catch {
        Record-TestResult "エラーハンドリング - 無効なアプリケーションパス" $true
    }

    # 存在しないウィンドウでのテスト
    Write-Host "`n8.2 存在しないウィンドウでのテスト" -ForegroundColor Cyan
    try {
        $newDriver = [UIAutomationDriver]::new()
        $newDriver.FindWindow("存在しないウィンドウ")
        Record-TestResult "エラーハンドリング - 存在しないウィンドウ" $false "エラーが発生すべきでした"
    }
    catch {
        Record-TestResult "エラーハンドリング - 存在しないウィンドウ" $true
    }

    # 無効な要素でのテスト
    Write-Host "`n8.3 無効な要素でのテスト" -ForegroundColor Cyan
    try {
        $newDriver = [UIAutomationDriver]::new()
        $newDriver.ClickElement($null)
        Record-TestResult "エラーハンドリング - 無効な要素" $false "エラーが発生すべきでした"
    }
    catch {
        Record-TestResult "エラーハンドリング - 無効な要素" $true
    }

    # 新しいドライバーを破棄
    if ($newDriver) {
        $newDriver.Dispose()
    }

    # ========================================
    # 9. 統合テスト
    # ========================================
    Write-Host "`n--- 9. 統合テスト ---" -ForegroundColor Yellow

    # 電卓での完全な計算テスト
    Write-Host "`n9.1 電卓での完全な計算テスト" -ForegroundColor Cyan
    try {
        $integrationDriver = [UIAutomationDriver]::new()
        $integrationDriver.StartApplication("calc.exe", "")
        Start-Sleep -Seconds 2
        
        $integrationDriver.FindWindow("電卓")
        $integrationDriver.ActivateWindow()
        
        # 7をクリック
        $button7 = $integrationDriver.FindElementByName("7")
        $integrationDriver.ClickElement($button7)
        Start-Sleep -Milliseconds 500
        
        # +をクリック
        $buttonPlus = $integrationDriver.FindElementByName("+")
        $integrationDriver.ClickElement($buttonPlus)
        Start-Sleep -Milliseconds 500
        
        # 3をクリック
        $button3 = $integrationDriver.FindElementByName("3")
        $integrationDriver.ClickElement($button3)
        Start-Sleep -Milliseconds 500
        
        # =をクリック
        $buttonEquals = $integrationDriver.FindElementByName("=")
        $integrationDriver.ClickElement($buttonEquals)
        Start-Sleep -Milliseconds 1000
        
        # 結果を取得
        $display = $integrationDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Text)
        $result = $integrationDriver.GetElementText($display)
        
        if ($result -eq "10") {
            Record-TestResult "統合テスト - 電卓計算（7+3=10）" $true
        } else {
            Record-TestResult "統合テスト - 電卓計算（7+3=10）" $false "期待値: 10, 実際の値: $result"
        }
        
        # スクリーンショット取得
        $integrationScreenshotPath = Join-Path $TestDir "integration_test_screenshot.png"
        $integrationDriver.TakeScreenshot($integrationScreenshotPath)
        
        $integrationDriver.CloseApplication()
        $integrationDriver.Dispose()
    }
    catch {
        Record-TestResult "統合テスト - 電卓計算（7+3=10）" $false $_.Exception.Message
    }

    # ========================================
    # テスト結果の表示
    # ========================================
    Write-Host "`n=== テスト結果サマリー ===" -ForegroundColor Cyan
    Write-Host "総テスト数: $($TestResults.TotalTests)" -ForegroundColor White
    Write-Host "成功: $($TestResults.PassedTests)" -ForegroundColor Green
    Write-Host "失敗: $($TestResults.FailedTests)" -ForegroundColor Red
    
    $successRate = if ($TestResults.TotalTests -gt 0) { 
        [math]::Round(($TestResults.PassedTests / $TestResults.TotalTests) * 100, 2) 
    } else { 
        0 
    }
    Write-Host "成功率: $successRate%" -ForegroundColor $(if ($successRate -ge 80) { "Green" } elseif ($successRate -ge 60) { "Yellow" } else { "Red" })

    # 失敗したテストの詳細表示
    if ($TestResults.FailedTests -gt 0) {
        Write-Host "`n--- 失敗したテストの詳細 ---" -ForegroundColor Red
        $failedTests = $TestResults.TestDetails | Where-Object { -not $_.Passed }
        foreach ($test in $failedTests) {
            Write-Host "✗ $($test.TestName)" -ForegroundColor Red
            if ($test.Message) {
                Write-Host "  エラー: $($test.Message)" -ForegroundColor Red
            }
        }
    }

    # テスト結果をファイルに保存
    $testReportPath = Join-Path $TestDir "test_report.json"
    $TestResults | ConvertTo-Json -Depth 3 | Out-File -FilePath $testReportPath -Encoding UTF8
    Write-Host "`nテストレポートを保存しました: $testReportPath" -ForegroundColor Green

    # 作成されたファイルの一覧表示
    Write-Host "`n--- 作成されたファイル ---" -ForegroundColor Cyan
    Get-ChildItem -Path $TestDir | ForEach-Object {
        Write-Host "- $($_.Name)" -ForegroundColor White
    }

    Write-Host "`n=== UIAutomationDriver 全メソッドテスト完了 ===" -ForegroundColor Cyan
    Write-Host "テスト終了時刻: $(Get-Date)" -ForegroundColor White
}
catch
{
    Write-Host "`nテスト実行中に重大なエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
    Record-TestResult "テスト実行全体" $false $_.Exception.Message
}
finally
{
    # リソースのクリーンアップ
    if ($guiDriver) {
        try {
            $guiDriver.Dispose()
        }
        catch {
            Write-Host "リソースのクリーンアップでエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

Write-Host "`nテストスクリプトが完了しました。" -ForegroundColor Green
