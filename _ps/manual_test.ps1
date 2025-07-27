# 手動エラーテストスクリプト
# 各エラー管理モジュールを直接テスト

Write-Host "=== 手動エラーテスト開始 ===" -ForegroundColor Green

# エラー管理モジュールをインポート
. "$PSScriptRoot\_lib\WebDriverErrors.ps1"
. "$PSScriptRoot\_lib\EdgeDriverErrors.ps1"
. "$PSScriptRoot\_lib\ChromeDriverErrors.ps1"
. "$PSScriptRoot\_lib\WordDriverErrors.ps1"

Write-Host "エラー管理モジュールが正常にインポートされました" -ForegroundColor Green

# WebDriverエラーテスト
Write-Host "`n--- WebDriverエラーテスト ---" -ForegroundColor Yellow
try {
    LogWebDriverError $WebDriverErrorCodes.INIT_ERROR "手動テスト: WebDriver初期化エラー"
    LogWebDriverError $WebDriverErrorCodes.NAVIGATE_ERROR "手動テスト: WebDriverナビゲーションエラー"
    LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENT_ERROR "手動テスト: WebDriver要素検索エラー"
    Write-Host "WebDriverエラーログが正常に出力されました" -ForegroundColor Green
}
catch {
    Write-Host "WebDriverエラーテストエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# EdgeDriverエラーテスト
Write-Host "`n--- EdgeDriverエラーテスト ---" -ForegroundColor Yellow
try {
    LogEdgeDriverError $EdgeDriverErrorCodes.INIT_ERROR "手動テスト: EdgeDriver初期化エラー"
    LogEdgeDriverError $EdgeDriverErrorCodes.EXECUTABLE_PATH_ERROR "手動テスト: EdgeDriver実行ファイルパスエラー"
    Write-Host "EdgeDriverエラーログが正常に出力されました" -ForegroundColor Green
}
catch {
    Write-Host "EdgeDriverエラーテストエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# ChromeDriverエラーテスト
Write-Host "`n--- ChromeDriverエラーテスト ---" -ForegroundColor Yellow
try {
    LogChromeDriverError $ChromeDriverErrorCodes.INIT_ERROR "手動テスト: ChromeDriver初期化エラー"
    LogChromeDriverError $ChromeDriverErrorCodes.EXECUTABLE_PATH_ERROR "手動テスト: ChromeDriver実行ファイルパスエラー"
    Write-Host "ChromeDriverエラーログが正常に出力されました" -ForegroundColor Green
}
catch {
    Write-Host "ChromeDriverエラーテストエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# WordDriverエラーテスト
Write-Host "`n--- WordDriverエラーテスト ---" -ForegroundColor Yellow
try {
    LogWordDriverError $WordDriverErrorCodes.INIT_ERROR "手動テスト: WordDriver初期化エラー"
    LogWordDriverError $WordDriverErrorCodes.SAVE_DOCUMENT_ERROR "手動テスト: WordDriverドキュメント保存エラー"
    Write-Host "WordDriverエラーログが正常に出力されました" -ForegroundColor Green
}
catch {
    Write-Host "WordDriverエラーテストエラー: $($_.Exception.Message)" -ForegroundColor Red
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

Write-Host "`n=== 手動エラーテスト完了 ===" -ForegroundColor Green
Write-Host "以下のログファイルが生成されているはずです:" -ForegroundColor Cyan
Write-Host "- WebDriver_Error.log" -ForegroundColor White
Write-Host "- EdgeDriver_Error.log" -ForegroundColor White
Write-Host "- ChromeDriver_Error.log" -ForegroundColor White
Write-Host "- WordDriver_Error.log" -ForegroundColor White 