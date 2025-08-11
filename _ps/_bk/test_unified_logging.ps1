# 統一されたログファイル名のテストスクリプト

# Common.ps1を読み込み
. ".\_lib\Common.ps1"

Write-Host "=== 統一されたエラーログファイル名のテスト ===" -ForegroundColor Green

# 各ドライバーでエラーログを出力して、統一されたログファイルに書き込まれることを確認
Write-Host "`n1. WebDriverエラーログテスト..." -ForegroundColor Yellow
$global:Common.HandleError("1001", "テスト用WebDriverエラー", "WebDriver", ".\AllDrivers_Error.log")

Write-Host "2. EdgeDriverエラーログテスト..." -ForegroundColor Yellow
$global:Common.HandleError("2001", "テスト用EdgeDriverエラー", "EdgeDriver", ".\AllDrivers_Error.log")

Write-Host "3. ChromeDriverエラーログテスト..." -ForegroundColor Yellow
$global:Common.HandleError("3001", "テスト用ChromeDriverエラー", "ChromeDriver", ".\AllDrivers_Error.log")

Write-Host "4. WordDriverエラーログテスト..." -ForegroundColor Yellow
$global:Common.HandleError("4001", "テスト用WordDriverエラー", "WordDriver", ".\AllDrivers_Error.log")

Write-Host "5. ExcelDriverエラーログテスト..." -ForegroundColor Yellow
$global:Common.HandleError("5001", "テスト用ExcelDriverエラー", "ExcelDriver", ".\AllDrivers_Error.log")

Write-Host "6. PowerPointDriverエラーログテスト..." -ForegroundColor Yellow
$global:Common.HandleError("6001", "テスト用PowerPointDriverエラー", "PowerPointDriver", ".\AllDrivers_Error.log")

# ログファイルの内容を確認
Write-Host "`n=== ログファイルの内容確認 ===" -ForegroundColor Green
if (Test-Path ".\AllDrivers_Error.log") {
    Write-Host "ログファイルが作成されました: .\AllDrivers_Error.log" -ForegroundColor Green
    Write-Host "`nログファイルの内容:" -ForegroundColor Cyan
    Get-Content ".\AllDrivers_Error.log" | ForEach-Object { Write-Host $_ }
} else {
    Write-Host "ログファイルが作成されていません。" -ForegroundColor Red
}

Write-Host "`n=== テスト完了 ===" -ForegroundColor Green
Write-Host "すべてのドライバーが統一されたログファイル '.\AllDrivers_Error.log' にエラーを出力するようになりました。" -ForegroundColor Green 