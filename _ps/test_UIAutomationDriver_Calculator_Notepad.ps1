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
        try
        {
            $display = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Text)
            if ($display -ne $null)
            {
                $result = $guiDriver.GetElementText($display)
                Write-Host "計算結果: 7×8 = $result" -ForegroundColor Green
            }
            else
            {
                Write-Host "表示エリアが見つかりませんでした" -ForegroundColor Yellow
            }
        }
        catch
        {
            Write-Host "計算結果の取得でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Write-Host "電卓での計算が完了しました！" -ForegroundColor Green
    }
    catch
    {
        Write-Host "電卓の計算でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

    # スクリーンショット取得
    try
    {
        $screenshotPath = ".\calc_screenshot.png"
        Write-Host "電卓のスクリーンショットを取得中..." -ForegroundColor Yellow
        $guiDriver.TakeScreenshot($screenshotPath)
        Write-Host "スクリーンショットを保存しました: $screenshotPath" -ForegroundColor Green
    }
    catch
    {
        Write-Host "スクリーンショット取得でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

    # 電卓終了
    try
    {
        Write-Host "電卓を終了中..." -ForegroundColor Yellow
        $guiDriver.CloseApplication()
        Start-Sleep -Seconds 1
        Write-Host "電卓を終了しました" -ForegroundColor Green
    }
    catch
    {
        Write-Host "電卓の終了でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

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

    # ========================================
    # EdgeブラウザでのGoogle検索テスト
    # ========================================
    Write-Host "`n--- EdgeブラウザでのGoogle検索テスト ---" -ForegroundColor Yellow

    # Edgeブラウザ起動
    $edgePath = $null
    
    # Edgeの実行ファイルパスを検索
    $edgePaths = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "msedge.exe"
    )
    
    foreach ($path in $edgePaths)
    {
        if (Test-Path $path)
        {
            $edgePath = $path
            Write-Host "Edgeブラウザを発見しました: $path" -ForegroundColor Green
            break
        }
    }
    
    if ($edgePath -eq $null)
    {
        Write-Host "Edgeブラウザが見つかりませんでした。利用可能なパス:" -ForegroundColor Red
        foreach ($path in $edgePaths)
        {
            Write-Host "- $path" -ForegroundColor Yellow
        }
        Write-Host "Edgeテストをスキップします。" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Edgeブラウザ起動: $edgePath" -ForegroundColor Green
    $guiDriver.StartApplication($edgePath, "")
    Write-Host "Edgeブラウザの起動を待機中..." -ForegroundColor Yellow
    Start-Sleep -Seconds 8

    # Edgeウィンドウ検索（複数のタイトルパターンを試行）
    $edgeWindow = $null
    
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
    
    # Edge関連のウィンドウを検索
    $edgeTitles = @("新しいタブ - プロファイル 1 - Microsoft​ Edge", "Microsoft Edge", "Edge", "新しいタブ", "New Tab", "Microsoft Edge - 新しいタブ", "Microsoft Edge - New Tab", "Bing", "Google")
    
    foreach ($title in $edgeTitles)
    {
        try
        {
            Write-Host "Edgeウィンドウを検索中: $title" -ForegroundColor Yellow
            $edgeWindow = $guiDriver.FindWindow($title)
            if ($edgeWindow -ne $null)
            {
                Write-Host "Edgeウィンドウを発見しました: $title" -ForegroundColor Green
                break
            }
        }
        catch
        {
            Write-Host "ウィンドウタイトル '$title' で見つかりませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    if ($edgeWindow -eq $null)
    {
        Write-Host "Edgeウィンドウが見つかりませんでした。利用可能なタイトル: $($edgeTitles -join ', ')" -ForegroundColor Red
        Write-Host "Edgeテストをスキップします。" -ForegroundColor Yellow
    }
    else
    {
        $guiDriver.ActivateWindow()

        # Googleに直接移動
        try
        {
            Write-Host "Googleに直接移動中..." -ForegroundColor Yellow
            
            # 方法1: Ctrl+LでアドレスバーにフォーカスしてURL入力
            Write-Host "Ctrl+Lでアドレスバーにフォーカスします" -ForegroundColor Yellow
            $guiDriver.SendKeyCombination(@("Ctrl", "L"))
            Start-Sleep -Milliseconds 1000
            
            # GoogleのURLを入力
            Write-Host "GoogleのURLを入力中..." -ForegroundColor Yellow
            $guiDriver.TypeText("https://www.google.com")
            $guiDriver.SendSpecialKey("Enter")
            Start-Sleep -Seconds 5
            
            # ページの読み込み完了を待つ
            Write-Host "ページの読み込み完了を待機中..." -ForegroundColor Yellow
            
            # 検索ボックスを検索（複数の方法を試行）
            Write-Host "検索ボックスを検索中..." -ForegroundColor Yellow
            $searchBox = $null
            
            # 方法1: 名前で検索
            $searchBoxNames = @("検索", "Search", "Google検索", "Google Search", "検索ボックス", "Search box", "q", "検索またはアドレスを入力")
            foreach ($name in $searchBoxNames)
            {
                try
                {
                    $searchBox = $guiDriver.FindElementByName($name)
                    if ($searchBox -ne $null)
                    {
                        Write-Host "検索ボックス（$name）を発見しました" -ForegroundColor Green
                        break
                    }
                }
                catch
                {
                    Write-Host "検索ボックス '$name' が見つかりませんでした" -ForegroundColor Yellow
                }
            }
            
            # 方法2: テキストボックスコントロールタイプで検索
            if ($searchBox -eq $null)
            {
                try
                {
                    Write-Host "テキストボックスコントロールタイプで検索中..." -ForegroundColor Yellow
                    $searchBox = $guiDriver.FindElementByControlType([System.Windows.Automation.ControlType]::Edit)
                    if ($searchBox -ne $null)
                    {
                        Write-Host "テキストボックスを発見しました" -ForegroundColor Green
                    }
                }
                catch
                {
                    Write-Host "テキストボックスも見つかりませんでした" -ForegroundColor Yellow
                }
            }
            
            # 方法3: 自動化IDで検索
            if ($searchBox -eq $null)
            {
                try
                {
                    Write-Host "自動化IDで検索中..." -ForegroundColor Yellow
                    $searchBox = $guiDriver.FindElementByAutomationId("q")
                    if ($searchBox -ne $null)
                    {
                        Write-Host "自動化ID 'q' で検索ボックスを発見しました" -ForegroundColor Green
                    }
                }
                catch
                {
                    Write-Host "自動化ID 'q' でも見つかりませんでした" -ForegroundColor Yellow
                }
            }
            
            # 検索を実行
            if ($searchBox -ne $null)
            {
                Write-Host "検索キーワードを入力中..." -ForegroundColor Yellow
                try
                {
                    $guiDriver.ClickElement($searchBox)
                    Start-Sleep -Milliseconds 500
                    $guiDriver.SetElementText($searchBox, "UIAutomationDriver テスト")
                    Start-Sleep -Milliseconds 500
                    
                    # Enterキーで検索実行
                    $guiDriver.SendSpecialKey("Enter")
                    Start-Sleep -Seconds 3
                    
                    Write-Host "Google検索が完了しました！" -ForegroundColor Green
                }
                catch
                {
                    Write-Host "要素操作でエラーが発生しました。キーボードで検索を実行します" -ForegroundColor Yellow
                    $guiDriver.TypeText("UIAutomationDriver テスト")
                    $guiDriver.SendSpecialKey("Enter")
                    Start-Sleep -Seconds 3
                    Write-Host "キーボードでの検索が完了しました！" -ForegroundColor Green
                }
            }
            else
            {
                Write-Host "検索ボックスが見つからないため、キーボードで検索を実行します" -ForegroundColor Yellow
                # Tabキーで検索ボックスにフォーカスを移動
                $guiDriver.SendSpecialKey("Tab")
                Start-Sleep -Milliseconds 500
                $guiDriver.TypeText("UIAutomationDriver テスト")
                $guiDriver.SendSpecialKey("Enter")
                Start-Sleep -Seconds 3
                Write-Host "キーボードでの検索が完了しました！" -ForegroundColor Green
            }
        }
        catch
        {
            Write-Host "Edgeでの検索でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Edgeスクリーンショット取得
        $edgeScreenshotPath = ".\edge_google_search_screenshot.png"
        $guiDriver.TakeScreenshot($edgeScreenshotPath)

        # Edgeブラウザ終了
        Write-Host "Edgeブラウザを終了中..." -ForegroundColor Yellow
        $guiDriver.CloseApplication()
        Start-Sleep -Seconds 2
    }

    Write-Host "`n=== UIAutomationDriver テスト完了 ===" -ForegroundColor Cyan
    Write-Host "スクリーンショットが保存されました:" -ForegroundColor Green
    Write-Host "- $screenshotPath" -ForegroundColor White
    Write-Host "- $windowScreenshotPath" -ForegroundColor White
    Write-Host "- $finalScreenshotPath" -ForegroundColor White
    if ($edgeScreenshotPath)
    {
        Write-Host "- $edgeScreenshotPath" -ForegroundColor White
    }
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
