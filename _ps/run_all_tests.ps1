# _libディレクトリ内のすべてのクラスのメソッドをテストするプログラムの実行ラッパー
# 作成日: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')

Write-Host "="*80 -ForegroundColor Green
Write-Host "_libディレクトリ内のクラスメソッドテストプログラム" -ForegroundColor Green
Write-Host "="*80 -ForegroundColor Green
Write-Host ""

# 実行開始時刻を記録
$startTime = Get-Date
Write-Host "テスト開始時刻: $($startTime.ToString('yyyy/MM/dd HH:mm:ss'))" -ForegroundColor Yellow
Write-Host ""

# 基本的なテストプログラムを実行
Write-Host "1. 基本的なテストプログラムを実行中..." -ForegroundColor Cyan
Write-Host "="*60 -ForegroundColor Cyan

try {
    & "$PSScriptRoot\test_all_methods.ps1"
    Write-Host "基本的なテストプログラムが正常に完了しました。" -ForegroundColor Green
} catch {
    Write-Host "基本的なテストプログラムでエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. 詳細テストプログラムを実行中..." -ForegroundColor Cyan
Write-Host "="*60 -ForegroundColor Cyan

try {
    & "$PSScriptRoot\test_methods_detailed.ps1"
    Write-Host "詳細テストプログラムが正常に完了しました。" -ForegroundColor Green
} catch {
    Write-Host "詳細テストプログラムでエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
}

# 実行終了時刻を記録
$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host ""
Write-Host "="*80 -ForegroundColor Green
Write-Host "テスト実行完了" -ForegroundColor Green
Write-Host "="*80 -ForegroundColor Green
Write-Host "テスト終了時刻: $($endTime.ToString('yyyy/MM/dd HH:mm:ss'))" -ForegroundColor Yellow
Write-Host "実行時間: $($duration.TotalMinutes.ToString('F2')) 分" -ForegroundColor Yellow
Write-Host ""

# 生成されたファイルの確認
Write-Host "生成されたファイル:" -ForegroundColor Cyan
$testFiles = Get-ChildItem -Path "." -Filter "test_results_*.csv" | Sort-Object LastWriteTime -Descending
$detailedFiles = Get-ChildItem -Path "." -Filter "detailed_test_results_*.csv" | Sort-Object LastWriteTime -Descending

if ($testFiles) {
    Write-Host "基本的なテスト結果ファイル:" -ForegroundColor White
    foreach ($file in $testFiles) {
        Write-Host "  - $($file.Name) (作成日時: $($file.LastWriteTime.ToString('yyyy/MM/dd HH:mm:ss')))" -ForegroundColor White
    }
}

if ($detailedFiles) {
    Write-Host "詳細テスト結果ファイル:" -ForegroundColor White
    foreach ($file in $detailedFiles) {
        Write-Host "  - $($file.Name) (作成日時: $($file.LastWriteTime.ToString('yyyy/MM/dd HH:mm:ss')))" -ForegroundColor White
    }
}

Write-Host ""
Write-Host "テストプログラムの実行が完了しました。" -ForegroundColor Green
Write-Host "結果は上記のCSVファイルで確認できます。" -ForegroundColor Cyan 