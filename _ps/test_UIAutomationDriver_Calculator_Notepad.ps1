# UIAutomationDriverテストスクリプト
# 電卓とメモ帳アプリケーションでのテスト

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
    Write-Host "電卓起動: $calcPath" -ForegroundColor Green
    $guiDriver.StartApplication($calcPath, "")
    Start-Sleep -Seconds 2

    # 電卓ウィンドウ検索
    $calcWindow = $guiDriver.FindWindow("電卓")
    Write-Host "電卓ウィンドウ検索: $calcWindow" -ForegroundColor Green
    $guiDriver.ActivateWindow()

    # 電卓のボタンを検索してクリック
    try
    {
        # 電卓の表示エリアを検索（Textコントロールタイプ）
        Write-Host "電卓の表示エリアを検索中..." -ForegroundColor Yellow
        $display = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Text)
        if ($display -ne $null)
        {
            $displayText = $guiDriver.GetElementText($display)
            Write-Host "表示エリアの内容: $displayText" -ForegroundColor Green
        }
    }
    catch
    {
        Write-Host "表示エリアの検索でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
        
        # 代替方法：名前で検索
        try
        {
            Write-Host "代替方法で要素を検索中..." -ForegroundColor Yellow
            $display = $guiDriver.FindElementByName("0")
            if ($display -ne $null)
            {
                $displayText = $guiDriver.GetElementText($display)
                Write-Host "代替方法で発見した要素: $displayText" -ForegroundColor Green
            }
        }
        catch
        {
            Write-Host "代替方法でも要素が見つかりませんでした: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # 電卓で簡単な計算を実行（7×8）
    try
    {
        Write-Host "電卓で計算を実行中（7×8）..." -ForegroundColor Yellow
        
        # 数字ボタン7を検索してクリック
        Write-Host "数字ボタン7を検索中..." -ForegroundColor Yellow
        $button7 = $guiDriver.FindElementByName("7")
        if ($button7 -ne $null)
        {
            $buttonText = $guiDriver.GetElementText($button7)
            Write-Host "発見したボタン: $buttonText" -ForegroundColor Green
            $guiDriver.ClickElement($button7)
            Start-Sleep -Milliseconds 500
        }
        
        # 乗算ボタンを検索してクリック（複数のパターンを試行）
        Write-Host "乗算ボタンを検索中..." -ForegroundColor Yellow
        $multiplyButton = $null
        $multiplyNames = @("×", "*", "乗算", "Multiply", "multiply")
        
        foreach ($name in $multiplyNames)
        {
            try
            {
                $multiplyButton = $guiDriver.FindElementByName($name)
                if ($multiplyButton -ne $null)
                {
                    Write-Host "乗算ボタン（$name）を発見しました" -ForegroundColor Green
                    $guiDriver.ClickElement($multiplyButton)
                    Start-Sleep -Milliseconds 500
                    break
                }
            }
            catch
            {
                Write-Host "乗算ボタン '$name' が見つかりませんでした" -ForegroundColor Yellow
            }
        }
        
        if ($multiplyButton -eq $null)
        {
            # 代替方法：キーボードで乗算記号を入力
            Write-Host "乗算ボタンが見つからないため、キーボードで入力します" -ForegroundColor Yellow
            $guiDriver.TypeText("*")
            Start-Sleep -Milliseconds 500
        }
        
        # 数字ボタン8を検索してクリック
        Write-Host "数字ボタン8を検索中..." -ForegroundColor Yellow
        $button8 = $guiDriver.FindElementByName("8")
        if ($button8 -ne $null)
        {
            $buttonText = $guiDriver.GetElementText($button8)
            Write-Host "発見したボタン: $buttonText" -ForegroundColor Green
            $guiDriver.ClickElement($button8)
            Start-Sleep -Milliseconds 500
        }
        
        # 等号ボタンを検索してクリック（複数のパターンを試行）
        Write-Host "等号ボタンを検索中..." -ForegroundColor Yellow
        $equalsButton = $null
        $equalsNames = @("=", "等号", "Equals", "equals", "計算", "Calculate")
        
        foreach ($name in $equalsNames)
        {
            try
            {
                $equalsButton = $guiDriver.FindElementByName($name)
                if ($equalsButton -ne $null)
                {
                    Write-Host "等号ボタン（$name）を発見しました" -ForegroundColor Green
                    $guiDriver.ClickElement($equalsButton)
                    Start-Sleep -Milliseconds 1000
                    break
                }
            }
            catch
            {
                Write-Host "等号ボタン '$name' が見つかりませんでした" -ForegroundColor Yellow
            }
        }
        
        if ($equalsButton -eq $null)
        {
            # 代替方法：キーボードで等号を入力
            Write-Host "等号ボタンが見つからないため、キーボードで入力します" -ForegroundColor Yellow
            $guiDriver.TypeText("=")
            Start-Sleep -Milliseconds 1000
        }
        
        # 計算結果を取得
        Write-Host "計算結果を取得中..." -ForegroundColor Yellow
        $display = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Text)
        if ($display -ne $null)
        {
            $result = $guiDriver.GetElementText($display)
            Write-Host "計算結果: 7×8 = $result" -ForegroundColor Green
        }
        
        Write-Host "電卓での計算が完了しました！" -ForegroundColor Green
    }
    catch
    {
        Write-Host "電卓の計算でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
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
    $guiDriver.StartApplication($notepadPath, "")
    Start-Sleep -Seconds 2

    # メモ帳ウィンドウ検索（デバッグ情報付き）
    $notepadWindow = $null
    
    # まず利用可能なウィンドウを確認
    Write-Host "利用可能なウィンドウを確認中..." -ForegroundColor Yellow
    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    $windows = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, 
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Window
        )))
    
    Write-Host "発見されたウィンドウ:" -ForegroundColor Cyan
    foreach ($window in $windows)
    {
        $windowName = $window.GetCurrentPropertyValue([System.Windows.Automation.AutomationElement]::NameProperty)
        Write-Host "- $windowName" -ForegroundColor White
    }
    
    # メモ帳関連のウィンドウを検索
    $notepadTitles = @("タイトルなし - メモ帳", "無題 - メモ帳", "メモ帳", "Notepad", "Untitled - Notepad")
    $notepadTitles = @("無題 - メモ帳", "メモ帳", "Notepad", "Untitled - Notepad", "Untitled - Notepad", "Notepad", "タイトルなし", "タイトルなし - メモ帳", "タイトルなし - Notepad")

    
    foreach ($title in $notepadTitles)
    {
        try
        {
            Write-Host "メモ帳ウィンドウを検索中: $title" -ForegroundColor Yellow
            $notepadWindow = $guiDriver.FindWindow($title)
            if ($notepadWindow -ne $null)
            {
                Write-Host "メモ帳ウィンドウを発見しました: $title" -ForegroundColor Green
                break
            }
        }
        catch
        {
            Write-Host "ウィンドウタイトル '$title' で見つかりませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    if ($notepadWindow -eq $null)
    {
        Write-Host "メモ帳ウィンドウが見つかりませんでした。利用可能なタイトル: $($notepadTitles -join ', ')" -ForegroundColor Red
        Write-Host "テストを続行しますが、メモ帳の操作はスキップされます。" -ForegroundColor Yellow
        return
    }
    
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
    $guiDriver.StartApplication($notepadPath, "")
    Start-Sleep -Seconds 2

    # メモ帳ウィンドウ検索（複数のタイトルパターンを試行）
    $notepadWindow = $null
    $notepadTitles = @("無題 - メモ帳", "メモ帳", "Notepad", "Untitled - Notepad", "Untitled - Notepad", "Notepad", "タイトルなし", "タイトルなし - メモ帳", "タイトルなし - Notepad")
    
    foreach ($title in $notepadTitles)
    {
        try
        {
            Write-Host "メモ帳ウィンドウを検索中: $title" -ForegroundColor Yellow
            $notepadWindow = $guiDriver.FindWindow($title)
            if ($notepadWindow -ne $null)
            {
                Write-Host "メモ帳ウィンドウを発見しました: $title" -ForegroundColor Green
                break
            }
        }
        catch
        {
            Write-Host "ウィンドウタイトル '$title' で見つかりませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    if ($notepadWindow -eq $null)
    {
        throw "メモ帳ウィンドウが見つかりませんでした。利用可能なタイトル: $($notepadTitles -join ', ')"
    }
    
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
