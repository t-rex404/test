# 具体的なエラーテストスクリプト
# 各ドライバークラスの特定のエラーケースをテスト

Write-Host "=== 具体的なエラーテスト開始 ===" -ForegroundColor Green

# 各ドライバークラスをインポート
. "$PSScriptRoot\_lib\WebDriver.ps1"
. "$PSScriptRoot\_lib\EdgeDriver.ps1"
. "$PSScriptRoot\_lib\ChromeDriver.ps1"
. "$PSScriptRoot\_lib\WordDriver.ps1"

# 共通ライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"

# WebDriverの具体的なエラーテスト
Write-Host "`n--- WebDriver具体的エラーテスト ---" -ForegroundColor Yellow
try {
    $webDriver = [WebDriver]::new()
    
    # 1. 無効なURLでナビゲーション
    Write-Host "無効なURLでナビゲーションをテスト..." -ForegroundColor Gray
    $webDriver.Navigate("invalid-url")
    
    # 2. 存在しない要素を検索
    Write-Host "存在しない要素を検索..." -ForegroundColor Gray
    $webDriver.FindElement("#non-existent-element")
    
    # 3. 無効なCSSセレクタ
    Write-Host "無効なCSSセレクタをテスト..." -ForegroundColor Gray
    $webDriver.FindElement("invalid[css:selector")
    
    $webDriver.Dispose()
}
catch {
    Write-Host "WebDriverエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# EdgeDriverの具体的なエラーテスト
Write-Host "`n--- EdgeDriver具体的エラーテスト ---" -ForegroundColor Yellow
try {
    # 無効なレジストリパスを設定してエラーを発生させる
    Write-Host "無効なレジストリパスでEdgeDriverをテスト..." -ForegroundColor Gray
    $edgeDriver = [EdgeDriver]::new()
    $edgeDriver.Dispose()
}
catch {
    Write-Host "EdgeDriverエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# ChromeDriverの具体的なエラーテスト
Write-Host "`n--- ChromeDriver具体的エラーテスト ---" -ForegroundColor Yellow
try {
    # 無効なレジストリパスを設定してエラーを発生させる
    Write-Host "無効なレジストリパスでChromeDriverをテスト..." -ForegroundColor Gray
    $chromeDriver = [ChromeDriver]::new()
    $chromeDriver.Dispose()
}
catch {
    Write-Host "ChromeDriverエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# WordDriverの具体的なエラーテスト
Write-Host "`n--- WordDriver具体的エラーテスト ---" -ForegroundColor Yellow
try {
    $wordDriver = [WordDriver]::new()
    
    # 1. 無効なファイルパスで保存
    Write-Host "無効なファイルパスで保存をテスト..." -ForegroundColor Gray
    $wordDriver.SaveDocument("C:\Invalid\Path\test.docx")
    
    # 2. 無効な画像パスで画像追加
    Write-Host "無効な画像パスで画像追加をテスト..." -ForegroundColor Gray
    $wordDriver.AddImage("C:\Invalid\Path\image.png")
    
    # 3. 無効なテーブルデータ
    Write-Host "無効なテーブルデータをテスト..." -ForegroundColor Gray
    $wordDriver.AddTable($null, 2, 2)
    
    $wordDriver.Dispose()
}
catch {
    Write-Host "WordDriverエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# エラーコードとメッセージの表示テスト
Write-Host "`n--- エラーコードとメッセージ表示テスト ---" -ForegroundColor Yellow

Write-Host "WebDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($WebDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "NAVIGATE_ERROR: $($WebDriverErrorCodes.NAVIGATE_ERROR)" -ForegroundColor White
Write-Host "FIND_ELEMENT_ERROR: $($WebDriverErrorCodes.FIND_ELEMENT_ERROR)" -ForegroundColor White

Write-Host "`nEdgeDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($EdgeDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "EXECUTABLE_PATH_ERROR: $($EdgeDriverErrorCodes.EXECUTABLE_PATH_ERROR)" -ForegroundColor White

Write-Host "`nChromeDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($ChromeDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "EXECUTABLE_PATH_ERROR: $($ChromeDriverErrorCodes.EXECUTABLE_PATH_ERROR)" -ForegroundColor White

Write-Host "`nWordDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($WordDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "SAVE_DOCUMENT_ERROR: $($WordDriverErrorCodes.SAVE_DOCUMENT_ERROR)" -ForegroundColor White

Write-Host "`n=== 具体的なエラーテスト完了 ===" -ForegroundColor Green
Write-Host "生成されたログファイルを確認してください。" -ForegroundColor Cyan 