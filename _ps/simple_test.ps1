# シンプルなエラーテストスクリプト

Write-Host "=== シンプルなエラーテスト開始 ===" -ForegroundColor Green

# 共通ライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"

Write-Host "共通ライブラリが正常にインポートされました" -ForegroundColor Green

# エラーコードの表示テスト
Write-Host "`n--- エラーコード表示テスト ---" -ForegroundColor Yellow

Write-Host "WebDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($WebDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "NAVIGATE_ERROR: $($WebDriverErrorCodes.NAVIGATE_ERROR)" -ForegroundColor White

Write-Host "`nEdgeDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($EdgeDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "EXECUTABLE_PATH_ERROR: $($EdgeDriverErrorCodes.EXECUTABLE_PATH_ERROR)" -ForegroundColor White

Write-Host "`nChromeDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($ChromeDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "EXECUTABLE_PATH_ERROR: $($ChromeDriverErrorCodes.EXECUTABLE_PATH_ERROR)" -ForegroundColor White

Write-Host "`nWordDriverエラーコード例:" -ForegroundColor Cyan
Write-Host "INIT_ERROR: $($WordDriverErrorCodes.INIT_ERROR)" -ForegroundColor White
Write-Host "SAVE_DOCUMENT_ERROR: $($WordDriverErrorCodes.SAVE_DOCUMENT_ERROR)" -ForegroundColor White

# 共通エラーハンドリングテスト
Write-Host "`n--- 共通エラーハンドリングテスト ---" -ForegroundColor Yellow
try {
    $Common.HandleError("TEST001", "テスト用のエラーメッセージ", "TestModule")
    Write-Host "共通エラーハンドリングが正常に実行されました" -ForegroundColor Green
}
catch {
    Write-Host "共通エラーハンドリングエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# ログ出力テスト
Write-Host "`n--- ログ出力テスト ---" -ForegroundColor Yellow
try {
    $Common.WriteLog("テスト用のログメッセージ", "INFO")
    $Common.WriteLog("テスト用の警告メッセージ", "WARNING")
    $Common.WriteLog("テスト用のエラーメッセージ", "ERROR")
    Write-Host "ログ出力が正常に実行されました" -ForegroundColor Green
}
catch {
    Write-Host "ログ出力エラー: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== シンプルなエラーテスト完了 ===" -ForegroundColor Green
Write-Host "ログファイルを確認してください:" -ForegroundColor Cyan
Write-Host "- Common_Error.log" -ForegroundColor White
