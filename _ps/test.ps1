# テストスクリプト
# 各ドライバークラスのエラー出力をテスト

Write-Host "=== ドライバークラスエラーテスト開始 ===" -ForegroundColor Green

# 各ドライバークラスをインポート
. "$PSScriptRoot\_lib\WebDriver.ps1"
. "$PSScriptRoot\_lib\EdgeDriver.ps1"
. "$PSScriptRoot\_lib\ChromeDriver.ps1"
. "$PSScriptRoot\_lib\WordDriver.ps1"

# 共通ライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"

# WebDriverエラーテスト
Write-Host "`n--- WebDriverエラーテスト ---" -ForegroundColor Yellow
try {
    # 存在しないURLでWebDriverを初期化してエラーを発生させる
    $webDriver = [WebDriver]::new()
    $webDriver.Navigate("https://invalid-url-that-does-not-exist-12345.com")
}
catch {
    Write-Host "WebDriverエラーが正常にキャッチされました: $($_.Exception.Message)" -ForegroundColor Red
}

# EdgeDriverエラーテスト
Write-Host "`n--- EdgeDriverエラーテスト ---" -ForegroundColor Yellow
try {
    # 無効なユーザーデータディレクトリでEdgeDriverを初期化してエラーを発生させる
    $env:EDGE_USER_DATA_DIR = "C:\Invalid\Path\That\Does\Not\Exist"
    $edgeDriver = [EdgeDriver]::new()
}
catch {
    Write-Host "EdgeDriverエラーが正常にキャッチされました: $($_.Exception.Message)" -ForegroundColor Red
}

# ChromeDriverエラーテスト
Write-Host "`n--- ChromeDriverエラーテスト ---" -ForegroundColor Yellow
try {
    # 無効なユーザーデータディレクトリでChromeDriverを初期化してエラーを発生させる
    $env:CHROME_USER_DATA_DIR = "C:\Invalid\Path\That\Does\Not\Exist"
    $chromeDriver = [ChromeDriver]::new()
}
catch {
    Write-Host "ChromeDriverエラーが正常にキャッチされました: $($_.Exception.Message)" -ForegroundColor Red
}

# WordDriverエラーテスト
Write-Host "`n--- WordDriverエラーテスト ---" -ForegroundColor Yellow
try {
    # 無効な一時ディレクトリでWordDriverを初期化してエラーを発生させる
    $env:TEMP = "C:\Invalid\Path\That\Does\Not\Exist"
    $wordDriver = [WordDriver]::new()
}
catch {
    Write-Host "WordDriverエラーが正常にキャッチされました: $($_.Exception.Message)" -ForegroundColor Red
}

# 共通エラーハンドリングテスト
Write-Host "`n--- 共通エラーハンドリングテスト ---" -ForegroundColor Yellow
try {
    $Common.HandleError("TEST001", "テスト用のエラーメッセージ", "TestModule")
    Write-Host "共通エラーハンドリングが正常に実行されました" -ForegroundColor Green
}
catch {
    Write-Host "共通エラーハンドリングエラー: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== エラーテスト完了 ===" -ForegroundColor Green
Write-Host "ログファイルを確認してください:" -ForegroundColor Cyan
Write-Host "- WebDriver_Error.log" -ForegroundColor White
Write-Host "- EdgeDriver_Error.log" -ForegroundColor White
Write-Host "- ChromeDriver_Error.log" -ForegroundColor White
Write-Host "- WordDriver_Error.log" -ForegroundColor White
Write-Host "- Common_Error.log" -ForegroundColor White 