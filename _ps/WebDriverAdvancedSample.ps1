# WebDriver高度な機能サンプルスクリプト
# 新しく追加した機能の使用例

Write-Host "=== WebDriver高度な機能サンプル開始 ===" -ForegroundColor Green

# 各ドライバークラスをインポート
. "$PSScriptRoot\_lib\WebDriver.ps1"
. "$PSScriptRoot\_lib\EdgeDriver.ps1"
. "$PSScriptRoot\_lib\ChromeDriver.ps1"

# 共通ライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"

try {
    # EdgeDriverを使用してブラウザを起動
    Write-Host "`n--- EdgeDriverでブラウザを起動 ---" -ForegroundColor Yellow
    $edgeDriver = [EdgeDriver]::new()
    
    # テスト用のHTMLページに移動
    $testUrl = "https://httpbin.org/html"
    Write-Host "テストページに移動: $testUrl" -ForegroundColor Cyan
    $edgeDriver.Navigate($testUrl)
    
    # ページロード完了まで待機
    Write-Host "`n--- ページロード待機 ---" -ForegroundColor Yellow
    $edgeDriver.WaitForPageLoad(30)
    Write-Host "ページロードが完了しました" -ForegroundColor Green
    
    # 要素が表示されるまで待機
    Write-Host "`n--- 要素表示待機 ---" -ForegroundColor Yellow
    $edgeDriver.WaitForElementVisible("h1", 10)
    Write-Host "h1要素が表示されました" -ForegroundColor Green
    
    # 要素を検索してテキストを取得
    Write-Host "`n--- 要素検索とテキスト取得 ---" -ForegroundColor Yellow
    $h1Element = $edgeDriver.FindElement("h1")
    $h1Text = $edgeDriver.GetElementText($h1Element.nodeId)
    Write-Host "h1要素のテキスト: $h1Text" -ForegroundColor Green
    
    # 要素がクリック可能になるまで待機
    Write-Host "`n--- 要素クリック可能性待機 ---" -ForegroundColor Yellow
    $edgeDriver.WaitForElementClickable("h1", 10)
    Write-Host "h1要素がクリック可能になりました" -ForegroundColor Green
    
    # マウスホバー
    Write-Host "`n--- マウスホバー ---" -ForegroundColor Yellow
    $edgeDriver.MouseHover($h1Element.nodeId)
    Write-Host "h1要素にマウスホバーしました" -ForegroundColor Green
    
    # ダブルクリック
    Write-Host "`n--- ダブルクリック ---" -ForegroundColor Yellow
    $edgeDriver.DoubleClick($h1Element.nodeId)
    Write-Host "h1要素をダブルクリックしました" -ForegroundColor Green
    
    # JavaScript実行
    Write-Host "`n--- JavaScript実行 ---" -ForegroundColor Yellow
    $pageTitle = $edgeDriver.ExecuteScript("document.title")
    Write-Host "ページタイトル: $pageTitle" -ForegroundColor Green
    
    # カスタム条件で待機
    Write-Host "`n--- カスタム条件待機 ---" -ForegroundColor Yellow
    $edgeDriver.WaitForCondition("document.querySelector('h1') !== null", 10)
    Write-Host "h1要素の存在を確認しました" -ForegroundColor Green
    
    # クッキー操作
    Write-Host "`n--- クッキー操作 ---" -ForegroundColor Yellow
    $edgeDriver.SetCookie("test_cookie", "test_value", "", "/", 1)
    Write-Host "テストクッキーを設定しました" -ForegroundColor Green
    
    $cookieValue = $edgeDriver.GetCookie("test_cookie")
    Write-Host "取得したクッキー値: $cookieValue" -ForegroundColor Green
    
    # ローカルストレージ操作
    Write-Host "`n--- ローカルストレージ操作 ---" -ForegroundColor Yellow
    $edgeDriver.SetLocalStorage("test_key", "test_value")
    Write-Host "ローカルストレージに値を設定しました" -ForegroundColor Green
    
    $storageValue = $edgeDriver.GetLocalStorage("test_key")
    Write-Host "取得したローカルストレージ値: $storageValue" -ForegroundColor Green
    
    # 要素の属性を取得
    Write-Host "`n--- 要素属性取得 ---" -ForegroundColor Yellow
    $h1Class = $edgeDriver.GetElementAttribute($h1Element.nodeId, "class")
    Write-Host "h1要素のclass属性: $h1Class" -ForegroundColor Green
    
    # CSSプロパティを取得
    Write-Host "`n--- CSSプロパティ取得 ---" -ForegroundColor Yellow
    $h1Color = $edgeDriver.GetElementCssProperty($h1Element.nodeId, "color")
    Write-Host "h1要素の色: $h1Color" -ForegroundColor Green
    
    # スクリーンショット取得
    Write-Host "`n--- スクリーンショット取得 ---" -ForegroundColor Yellow
    $screenshotPath = ".\screenshot_test.png"
    $edgeDriver.GetScreenshot("viewPort", $screenshotPath)
    Write-Host "スクリーンショットを保存しました: $screenshotPath" -ForegroundColor Green
    
    # 要素スクリーンショット取得
    Write-Host "`n--- 要素スクリーンショット取得 ---" -ForegroundColor Yellow
    $elementScreenshotPath = ".\element_screenshot_test.png"
    $edgeDriver.GetScreenshotObjectId($h1Element.nodeId, $elementScreenshotPath)
    Write-Host "要素スクリーンショットを保存しました: $elementScreenshotPath" -ForegroundColor Green
    
    # ウィンドウサイズを取得
    Write-Host "`n--- ウィンドウサイズ取得 ---" -ForegroundColor Yellow
    $windowSize = $edgeDriver.GetWindowSize()
    Write-Host "ウィンドウサイズ: $($windowSize.width) x $($windowSize.height)" -ForegroundColor Green
    
    # ウィンドウを最大化
    Write-Host "`n--- ウィンドウ最大化 ---" -ForegroundColor Yellow
    $windowHandle = $edgeDriver.GetWindowHandle()
    $edgeDriver.MaximizeWindow($windowHandle)
    Write-Host "ウィンドウを最大化しました" -ForegroundColor Green
    
    # ウィンドウサイズを再取得
    $newWindowSize = $edgeDriver.GetWindowSize()
    Write-Host "最大化後のウィンドウサイズ: $($newWindowSize.width) x $($newWindowSize.height)" -ForegroundColor Green
    
    # ウィンドウを通常サイズに戻す
    $edgeDriver.NormalWindow($windowHandle)
    Write-Host "ウィンドウを通常サイズに戻しました" -ForegroundColor Green
    
    # URLとタイトルを取得
    Write-Host "`n--- URLとタイトル取得 ---" -ForegroundColor Yellow
    $currentUrl = $edgeDriver.GetUrl()
    $currentTitle = $edgeDriver.GetTitle()
    Write-Host "現在のURL: $currentUrl" -ForegroundColor Green
    Write-Host "現在のタイトル: $currentTitle" -ForegroundColor Green
    
    # ソースコードを取得（一部）
    Write-Host "`n--- ソースコード取得（一部） ---" -ForegroundColor Yellow
    $sourceCode = $edgeDriver.GetSourceCode()
    $sourcePreview = $sourceCode.Substring(0, [Math]::Min(200, $sourceCode.Length))
    Write-Host "ソースコード（最初の200文字）: $sourcePreview..." -ForegroundColor Green
    
    # 複数ウィンドウハンドルを取得
    Write-Host "`n--- 複数ウィンドウハンドル取得 ---" -ForegroundColor Yellow
    $windowHandles = $edgeDriver.GetWindowHandles()
    Write-Host "ウィンドウハンドル数: $($windowHandles.Count)" -ForegroundColor Green
    
    # 利用可能なタブ情報を取得
    Write-Host "`n--- 利用可能なタブ情報取得 ---" -ForegroundColor Yellow
    $availableTabs = $edgeDriver.GetAvailableTabs()
    Write-Host "利用可能なタブ数: $($availableTabs.targetInfos.Count)" -ForegroundColor Green
    
    # リソースを解放
    Write-Host "`n--- リソース解放 ---" -ForegroundColor Yellow
    $edgeDriver.Dispose()
    Write-Host "EdgeDriverのリソースを解放しました" -ForegroundColor Green
    
    Write-Host "`n=== WebDriver高度な機能サンプル完了 ===" -ForegroundColor Green
    Write-Host "すべての機能が正常に動作しました！" -ForegroundColor Green
}
catch {
    Write-Host "`nエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    # 確実にリソースを解放
    if ($edgeDriver) {
        try {
            $edgeDriver.Dispose()
        }
        catch {
            Write-Host "リソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

Write-Host "`n生成されたファイル:" -ForegroundColor Cyan
Write-Host "- screenshot_test.png" -ForegroundColor White
Write-Host "- element_screenshot_test.png" -ForegroundColor White
Write-Host "- WebDriver_Error.log" -ForegroundColor White
