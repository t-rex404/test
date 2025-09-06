# UIAutomationDriverテストスクリプト
# 電卓とメモ帳アプリケーションでのテスト

# 必要なモジュールを読み込み
. ".\_lib\Common.ps1"
. ".\_lib\UIAutomationDriver.ps1"

# グローバル変数の初期化
$global:Common = [Common]::new()

Write-Host "=== UIAutomationDriver テスト開始 ===" -ForegroundColor Cyan

try
{
    # UIAutomationDriverインスタンス作成
    $guiDriver = [UIAutomationDriver]::new()

    # ========================================
    # 電卓アプリケーションのテスト
    # ========================================
    Write-Host "`n--- 電卓アプリケーションのテスト ---" -ForegroundColor Yellow

    # 電卓起動
    $calcPath = "calc.exe"
    $guiDriver.StartApplication($calcPath)
    Start-Sleep -Seconds 2

    # 電卓ウィンドウ検索
    $calcWindow = $guiDriver.FindWindow("電卓")
    $guiDriver.ActivateWindow()

    # 電卓のボタンを検索してクリック
    try
    {
        # 数字ボタンを検索（例：7）
        $button7 = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Button)
        if ($button7 -ne $null)
        {
            $buttonText = $guiDriver.GetElementText($button7)
            Write-Host "発見したボタン: $buttonText" -ForegroundColor Green
        }

        # 電卓の表示エリアを検索
        $display = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Edit)
        if ($display -ne $null)
        {
            $displayText = $guiDriver.GetElementText($display)
            Write-Host "表示エリアの内容: $displayText" -ForegroundColor Green
        }
    }
    catch
    {
        Write-Host "電卓の要素検索でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

    # スクリーンショット取得
    $screenshotPath = ".\calc_screenshot.png"
    $guiDriver.TakeScreenshot($screenshotPath)

    # 電卓終了
    $guiDriver.CloseApplication()
    Start-Sleep -Seconds 1

    # ========================================
    # メモ帳アプリケーションのテスト
    # ========================================
    Write-Host "`n--- メモ帳アプリケーションのテスト ---" -ForegroundColor Yellow

    # メモ帳起動
    $notepadPath = "notepad.exe"
    $guiDriver.StartApplication($notepadPath)
    Start-Sleep -Seconds 2

    # メモ帳ウィンドウ検索
    $notepadWindow = $guiDriver.FindWindow("無題 - メモ帳")
    $guiDriver.ActivateWindow()

    # メモ帳のテキストエリアを検索
    try
    {
        $textArea = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Document)
        if ($textArea -ne $null)
        {
            Write-Host "メモ帳のテキストエリアを発見しました" -ForegroundColor Green
            
            # テキスト入力
            $guiDriver.SetElementText($textArea, "UIAutomationDriverテスト用のテキストです。`nこれは自動化テストのサンプルです。")
            Start-Sleep -Seconds 1

            # 入力されたテキストを取得
            $inputText = $guiDriver.GetElementText($textArea)
            Write-Host "入力されたテキスト: $inputText" -ForegroundColor Green
        }
    }
    catch
    {
        Write-Host "メモ帳の要素検索でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

    # ウィンドウスクリーンショット取得
    $windowScreenshotPath = ".\notepad_window_screenshot.png"
    $guiDriver.TakeWindowScreenshot($windowScreenshotPath)

    # メニューバーのテスト
    try
    {
        # ファイルメニューを検索
        $fileMenu = $guiDriver.FindElementByName("ファイル(&F)")
        if ($fileMenu -ne $null)
        {
            Write-Host "ファイルメニューを発見しました" -ForegroundColor Green
            $guiDriver.ClickElement($fileMenu)
            Start-Sleep -Seconds 1
        }
    }
    catch
    {
        Write-Host "メニューの検索でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

    # メモ帳終了（保存しない）
    $guiDriver.CloseApplication()
    Start-Sleep -Seconds 1

    # ========================================
    # キーボード操作のテスト
    # ========================================
    Write-Host "`n--- キーボード操作のテスト ---" -ForegroundColor Yellow

    # メモ帳を再起動
    $guiDriver.StartApplication($notepadPath)
    Start-Sleep -Seconds 2

    $notepadWindow = $guiDriver.FindWindow("無題 - メモ帳")
    $guiDriver.ActivateWindow()

    # キーボード入力テスト
    $guiDriver.TypeText("キーボード入力テスト: ")
    $guiDriver.SendSpecialKey("Enter")
    $guiDriver.TypeText("特殊キーテスト: ")
    $guiDriver.SendSpecialKey("Tab")
    $guiDriver.TypeText("Tabキーが押されました")
    $guiDriver.SendSpecialKey("Enter")
    $guiDriver.TypeText("Ctrl+Aテスト: ")
    $guiDriver.SendKeyCombination(@("Ctrl", "A"))
    $guiDriver.TypeText("全選択されました")

    Start-Sleep -Seconds 2

    # 最終スクリーンショット
    $finalScreenshotPath = ".\final_test_screenshot.png"
    $guiDriver.TakeScreenshot($finalScreenshotPath)

    # メモ帳終了
    $guiDriver.CloseApplication()

    Write-Host "`n=== UIAutomationDriver テスト完了 ===" -ForegroundColor Cyan
    Write-Host "スクリーンショットが保存されました:" -ForegroundColor Green
    Write-Host "- $screenshotPath" -ForegroundColor White
    Write-Host "- $windowScreenshotPath" -ForegroundColor White
    Write-Host "- $finalScreenshotPath" -ForegroundColor White
}
catch
{
    Write-Host "テスト実行中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally
{
    # リソースのクリーンアップ
    if ($guiDriver)
    {
        $guiDriver.Dispose()
    }
}

Write-Host "`nテストスクリプトが完了しました。" -ForegroundColor Green
