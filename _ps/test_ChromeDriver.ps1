# ChromeDriverクラステストファイル
# ChromeDriverクラスの各メソッドの機能をテストするためのスクリプト

# 必要なライブラリをインポート
. $PSScriptRoot\_lib\WebDriver.ps1
. $PSScriptRoot\_lib\ChromeDriver.ps1

# テスト用のHTMLファイルを作成
$testHtml = @"
<!DOCTYPE html>
<html>
<head>
    <title>ChromeDriver Test Page</title>
    <style>
        .test-class { color: blue; }
        .hidden { display: none; }
        .visible { display: block; }
    </style>
</head>
<body>
    <h1 id="main-title">ChromeDriver Test Page</h1>
    <div class="test-class">This is a test div for Chrome</div>
    <input type="text" id="test-input" name="test-input" value="initial value">
    <button id="test-button" onclick="alert('Chrome button clicked!')">Click Me</button>
    <select id="test-select">
        <option value="1">Option 1</option>
        <option value="2">Option 2</option>
        <option value="3">Option 3</option>
    </select>
    <input type="checkbox" id="test-checkbox">
    <input type="radio" name="test-radio" value="1" id="radio1">
    <input type="radio" name="test-radio" value="2" id="radio2">
    <a href="https://www.google.com" id="test-link">Google Link</a>
    <div id="hidden-element" class="hidden">Hidden Element</div>
    <div id="visible-element" class="visible">Visible Element</div>
    <table id="test-table">
        <tr><td>Cell 1</td><td>Cell 2</td></tr>
        <tr><td>Cell 3</td><td>Cell 4</td></tr>
    </table>
</body>
</html>
"@

# テスト用HTMLファイルを保存
$testHtmlPath = Join-Path $PSScriptRoot "test_chrome_page.html"
$testHtml | Out-File -FilePath $testHtmlPath -Encoding UTF8

Write-Host "ChromeDriverクラステストを開始します..." -ForegroundColor Green

try {
    # ChromeDriverインスタンスを作成
    Write-Host "ChromeDriverを初期化中..." -ForegroundColor Yellow
    $driver = [ChromeDriver]::new()
    
    # テストページに移動
    $testUrl = "file:///$($testHtmlPath -replace '\\', '/')"
    Write-Host "テストページに移動: $testUrl" -ForegroundColor Yellow
    $driver.Navigate($testUrl)
    
    # ChromeDriver固有の機能テスト
    Write-Host "`n=== ChromeDriver固有機能テスト ===" -ForegroundColor Cyan
    
    # Chrome実行ファイルパス取得テスト
    Write-Host "Chrome実行ファイルパス: $($driver.browser_exe_path)" -ForegroundColor Green
    
    # ユーザーデータディレクトリ取得テスト
    Write-Host "ユーザーデータディレクトリ: $($driver.browser_user_data_dir)" -ForegroundColor Green
    
    # Chrome初期化状態確認
    Write-Host "Chrome初期化状態: $($driver.is_chrome_initialized)" -ForegroundColor Green
    
    # 基本情報取得テスト
    Write-Host "`n=== 基本情報取得テスト ===" -ForegroundColor Cyan
    $title = $driver.GetTitle()
    Write-Host "ページタイトル: $title" -ForegroundColor Green
    
    $url = $driver.GetUrl()
    Write-Host "現在のURL: $url" -ForegroundColor Green
    
    # 要素検索テスト
    Write-Host "`n=== 要素検索テスト ===" -ForegroundColor Cyan
    
    # CSSセレクタで要素検索
    $element = $driver.FindElement("#main-title")
    Write-Host "CSSセレクタ検索成功: $($element.nodeId)" -ForegroundColor Green
    
    # XPathで要素存在確認
    $exists = $driver.IsExistsElementByXPath("//h1[@id='main-title']")
    Write-Host "XPath要素存在確認: $exists" -ForegroundColor Green
    
    # 要素テキスト取得テスト
    Write-Host "`n=== 要素テキスト取得テスト ===" -ForegroundColor Cyan
    $text = $driver.GetElementText($element.nodeId)
    Write-Host "要素テキスト: $text" -ForegroundColor Green
    
    # 要素操作テスト
    Write-Host "`n=== 要素操作テスト ===" -ForegroundColor Cyan
    
    # テキスト入力
    $inputElement = $driver.FindElement("#test-input")
    $driver.SetElementText($inputElement.nodeId, "Chromeテスト入力")
    Write-Host "テキスト入力完了" -ForegroundColor Green
    
    # 要素クリック
    $buttonElement = $driver.FindElement("#test-button")
    $driver.ClickElement($buttonElement.nodeId)
    Write-Host "ボタンクリック完了" -ForegroundColor Green
    
    # フォーム操作テスト
    Write-Host "`n=== フォーム操作テスト ===" -ForegroundColor Cyan
    
    # セレクトボックス操作
    $selectElement = $driver.FindElement("#test-select")
    $driver.SelectOptionByIndex($selectElement.nodeId, 1)
    Write-Host "セレクトボックス選択完了" -ForegroundColor Green
    
    # チェックボックス操作
    $checkboxElement = $driver.FindElement("#test-checkbox")
    $driver.SetCheckbox($checkboxElement.nodeId, $true)
    Write-Host "チェックボックス設定完了" -ForegroundColor Green
    
    # ラジオボタン操作
    $radioElement = $driver.FindElement("#radio1")
    $driver.SelectRadioButton($radioElement.nodeId)
    Write-Host "ラジオボタン選択完了" -ForegroundColor Green
    
    # 属性取得テスト
    Write-Host "`n=== 属性取得テスト ===" -ForegroundColor Cyan
    $linkElement = $driver.FindElement("#test-link")
    $href = $driver.GetHrefFromAnchor($linkElement.nodeId)
    Write-Host "リンクのhref: $href" -ForegroundColor Green
    
    # 要素存在確認テスト
    Write-Host "`n=== 要素存在確認テスト ===" -ForegroundColor Cyan
    $visibleExists = $driver.IsExistsElementByXPath("//div[@id='visible-element']")
    $hiddenExists = $driver.IsExistsElementByXPath("//div[@id='hidden-element']")
    Write-Host "表示要素存在: $visibleExists" -ForegroundColor Green
    Write-Host "非表示要素存在: $hiddenExists" -ForegroundColor Green
    
    # 待機機能テスト
    Write-Host "`n=== 待機機能テスト ===" -ForegroundColor Cyan
    $driver.WaitForElementVisible("#visible-element", 5)
    Write-Host "要素表示待機完了" -ForegroundColor Green
    
    $driver.WaitForElementClickable("#test-button", 5)
    Write-Host "要素クリック可能性待機完了" -ForegroundColor Green
    
    # スクリーンショットテスト
    Write-Host "`n=== スクリーンショットテスト ===" -ForegroundColor Cyan
    $screenshotPath = Join-Path $PSScriptRoot "chrome_test_screenshot.png"
    $driver.GetScreenshot("fullPage", $screenshotPath)
    Write-Host "フルページスクリーンショット保存: $screenshotPath" -ForegroundColor Green
    
    # 要素スクリーンショット
    $elementScreenshotPath = Join-Path $PSScriptRoot "chrome_element_screenshot.png"
    $driver.GetScreenshotObjectId($element.nodeId, $elementScreenshotPath)
    Write-Host "要素スクリーンショット保存: $elementScreenshotPath" -ForegroundColor Green
    
    # JavaScript実行テスト
    Write-Host "`n=== JavaScript実行テスト ===" -ForegroundColor Cyan
    $result = $driver.ExecuteScript("document.title")
    Write-Host "JavaScript実行結果: $result" -ForegroundColor Green
    
    # クッキー操作テスト
    Write-Host "`n=== クッキー操作テスト ===" -ForegroundColor Cyan
    $driver.SetCookie("chrome-test-cookie", "chrome-test-value")
    $cookieValue = $driver.GetCookie("chrome-test-cookie")
    Write-Host "クッキー値: $cookieValue" -ForegroundColor Green
    
    # ローカルストレージ操作テスト
    Write-Host "`n=== ローカルストレージ操作テスト ===" -ForegroundColor Cyan
    $driver.SetLocalStorage("chrome-test-key", "chrome-test-value")
    $storageValue = $driver.GetLocalStorage("chrome-test-key")
    Write-Host "ローカルストレージ値: $storageValue" -ForegroundColor Green
    
    # ウィンドウ操作テスト
    Write-Host "`n=== ウィンドウ操作テスト ===" -ForegroundColor Cyan
    $windowHandle = $driver.GetWindowHandle()
    Write-Host "ウィンドウハンドル: $windowHandle" -ForegroundColor Green
    
    $windowSize = $driver.GetWindowSize()
    Write-Host "ウィンドウサイズ: $($windowSize.width) x $($windowSize.height)" -ForegroundColor Green
    
    # ウィンドウ最大化
    $driver.MaximizeWindow($windowHandle)
    Write-Host "ウィンドウ最大化完了" -ForegroundColor Green
    
    # ナビゲーション履歴テスト
    Write-Host "`n=== ナビゲーション履歴テスト ===" -ForegroundColor Cyan
    $driver.Navigate("https://www.google.com")
    Write-Host "Googleに移動" -ForegroundColor Green
    
    $driver.GoBack()
    Write-Host "戻る操作完了" -ForegroundColor Green
    
    $driver.GoForward()
    Write-Host "進む操作完了" -ForegroundColor Green
    
    $driver.Refresh()
    Write-Host "ページ更新完了" -ForegroundColor Green
    
    Write-Host "`nすべてのテストが正常に完了しました！" -ForegroundColor Green
    
} catch {
    Write-Host "テスト中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    # リソースのクリーンアップ
    if ($driver) {
        Write-Host "`nリソースをクリーンアップ中..." -ForegroundColor Yellow
        $driver.Dispose()
    }
    
    # テスト用HTMLファイルを削除
    if (Test-Path $testHtmlPath) {
        Remove-Item $testHtmlPath -Force
        Write-Host "テスト用HTMLファイルを削除しました" -ForegroundColor Yellow
    }
    
    Write-Host "テスト終了" -ForegroundColor Green
} 