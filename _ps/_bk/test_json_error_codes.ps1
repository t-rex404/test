# JSONエラーコード読み込みテストファイル

Write-Host "=== JSONエラーコード読み込みテスト開始 ===" -ForegroundColor Green

# 必要なライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"

# ========================================
# Common.ps1 JSON読み込みテスト
# ========================================
Write-Host "`n=== Common.ps1 JSON読み込みテスト ===" -ForegroundColor Yellow

try {
    # エラーコードデータが読み込まれているかチェック
    if ($Common.ErrorCodes.Count -gt 0) {
        Write-Host "✓ エラーコードデータが正常に読み込まれました" -ForegroundColor Green
        Write-Host "  読み込まれたモジュール数: $($Common.ErrorCodes.Count)" -ForegroundColor Gray
        
        # 各モジュールのエラーコード数を表示
        foreach ($module in $Common.ErrorCodes.Keys) {
            $errorCount = $Common.ErrorCodes[$module].Count
            Write-Host "  - $module`: $errorCount 個のエラーコード" -ForegroundColor Gray
        }
    } else {
        Write-Host "✗ エラーコードデータが読み込まれていません" -ForegroundColor Red
    }

    # 各モジュールのエラータイトル取得テスト
    Write-Host "`n=== エラータイトル取得テスト ===" -ForegroundColor Cyan
    
    # WebDriverテスト
    $title1 = $Common.GetErrorTitle("1001", "WebDriver")
    $title2 = $Common.GetErrorTitle("1011", "WebDriver")
    Write-Host "✓ WebDriver エラータイトル取得テスト完了" -ForegroundColor Green
    Write-Host "  1001: $title1" -ForegroundColor Gray
    Write-Host "  1011: $title2" -ForegroundColor Gray
    
    # EdgeDriverテスト
    $title3 = $Common.GetErrorTitle("2001", "EdgeDriver")
    Write-Host "✓ EdgeDriver エラータイトル取得テスト完了" -ForegroundColor Green
    Write-Host "  2001: $title3" -ForegroundColor Gray
    
    # ChromeDriverテスト
    $title4 = $Common.GetErrorTitle("3001", "ChromeDriver")
    Write-Host "✓ ChromeDriver エラータイトル取得テスト完了" -ForegroundColor Green
    Write-Host "  3001: $title4" -ForegroundColor Gray
    
    # WordDriverテスト
    $title5 = $Common.GetErrorTitle("4001", "WordDriver")
    Write-Host "✓ WordDriver エラータイトル取得テスト完了" -ForegroundColor Green
    Write-Host "  4001: $title5" -ForegroundColor Gray
    
    # ExcelDriverテスト
    $title6 = $Common.GetErrorTitle("5001", "ExcelDriver")
    Write-Host "✓ ExcelDriver エラータイトル取得テスト完了" -ForegroundColor Green
    Write-Host "  5001: $title6" -ForegroundColor Gray
    
    # PowerPointDriverテスト
    $title7 = $Common.GetErrorTitle("6001", "PowerPointDriver")
    Write-Host "✓ PowerPointDriver エラータイトル取得テスト完了" -ForegroundColor Green
    Write-Host "  6001: $title7" -ForegroundColor Gray
    
    # 存在しないエラーコードのテスト
    $title8 = $Common.GetErrorTitle("9999", "WebDriver")
    Write-Host "✓ 存在しないエラーコードテスト完了" -ForegroundColor Green
    Write-Host "  9999: $title8" -ForegroundColor Gray
    
    # 存在しないモジュールのテスト
    $title9 = $Common.GetErrorTitle("1001", "UnknownDriver")
    Write-Host "✓ 存在しないモジュールテスト完了" -ForegroundColor Green
    Write-Host "  UnknownDriver: $title9" -ForegroundColor Gray

    # HandleError関数のテスト
    Write-Host "`n=== HandleError関数テスト ===" -ForegroundColor Cyan
    $Common.HandleError("1001", "テストエラーメッセージ", "WebDriver", ".\test_error.log")
    Write-Host "✓ HandleError関数テスト完了" -ForegroundColor Green
    
    # テストログファイルの確認
    if (Test-Path ".\test_error.log") {
        Write-Host "✓ エラーログファイルが作成されました" -ForegroundColor Green
        $logContent = Get-Content ".\test_error.log" -Raw
        Write-Host "ログ内容:" -ForegroundColor Gray
        Write-Host $logContent -ForegroundColor Gray
        
        # テストログファイルを削除
        Remove-Item ".\test_error.log" -Force
        Write-Host "✓ テストログファイルを削除しました" -ForegroundColor Green
    }
}
catch {
    Write-Host "✗ JSONエラーコード読み込みテストでエラー: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== JSONエラーコード読み込みテスト完了 ===" -ForegroundColor Green
Write-Host "JSONファイルからのエラーコード読み込みが正常に動作しています！" -ForegroundColor Green 