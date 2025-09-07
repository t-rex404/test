# Edgeブラウザの簡単なテストスクリプト
# Google検索のテスト

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

Write-Host "=== Edgeブラウザテスト開始 ===" -ForegroundColor Cyan

try
{
    # UIAutomationDriverインスタンス作成
    $guiDriver = [UIAutomationDriver]::new()

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
        return
    }
    
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
    try
    {
        $edgeScreenshotPath = ".\edge_google_search_screenshot.png"
        Write-Host "Edgeのスクリーンショットを取得中..." -ForegroundColor Yellow
        $guiDriver.TakeScreenshot($edgeScreenshotPath)
        Write-Host "スクリーンショットを保存しました: $edgeScreenshotPath" -ForegroundColor Green
    }
    catch
    {
        Write-Host "スクリーンショット取得でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Edgeブラウザ終了
    try
    {
        Write-Host "Edgeブラウザを終了中..." -ForegroundColor Yellow
        $guiDriver.CloseApplication()
        Start-Sleep -Seconds 2
        Write-Host "Edgeブラウザを終了しました" -ForegroundColor Green
    }
    catch
    {
        Write-Host "Edgeブラウザの終了でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host "`n=== Edgeブラウザテスト完了 ===" -ForegroundColor Cyan
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

Write-Host "`nEdgeブラウザテストスクリプトが完了しました。" -ForegroundColor Green
