# 統合テストファイル
# すべてのドライバークラスのテストを実行するためのスクリプト

Write-Host "統合テストを開始します..." -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

# テスト開始時刻を記録
$startTime = Get-Date
Write-Host "開始時刻: $startTime" -ForegroundColor Yellow

# テスト結果を格納する配列
$testResults = @()

# テスト関数
function Run-Test {
    param(
        [string]$TestName,
        [string]$TestScript,
        [string]$Description
    )
    
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "テスト: $TestName" -ForegroundColor White
    Write-Host "説明: $Description" -ForegroundColor Gray
    Write-Host "==========================================" -ForegroundColor Cyan
    
    $testStartTime = Get-Date
    
    try {
        # テストスクリプトを実行
        & $TestScript
        
        $testEndTime = Get-Date
        $duration = $testEndTime - $testStartTime
        
        $result = @{
            Name = $TestName
            Status = "SUCCESS"
            Duration = $duration
            Error = $null
        }
        
        Write-Host "`n✅ テスト成功: $TestName" -ForegroundColor Green
        Write-Host "実行時間: $($duration.TotalSeconds.ToString('F2'))秒" -ForegroundColor Green
        
    } catch {
        $testEndTime = Get-Date
        $duration = $testEndTime - $testStartTime
        
        $result = @{
            Name = $TestName
            Status = "FAILED"
            Duration = $duration
            Error = $_.Exception.Message
        }
        
        Write-Host "`n❌ テスト失敗: $TestName" -ForegroundColor Red
        Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "実行時間: $($duration.TotalSeconds.ToString('F2'))秒" -ForegroundColor Red
    }
    
    $testResults += $result
    return $result
}

# 各テストを実行
$tests = @(
    @{
        Name = "WebDriver基本テスト"
        Script = "test_WebDriver.ps1"
        Description = "WebDriverクラスの基本機能をテスト"
    },
    @{
        Name = "ChromeDriverテスト"
        Script = "test_ChromeDriver.ps1"
        Description = "ChromeDriverクラスの機能をテスト"
    },
    @{
        Name = "EdgeDriverテスト"
        Script = "test_EdgeDriver.ps1"
        Description = "EdgeDriverクラスの機能をテスト"
    },
    @{
        Name = "WordDriverテスト"
        Script = "test_WordDriver.ps1"
        Description = "WordDriverクラスの機能をテスト"
    }
)

# 各テストを実行
foreach ($test in $tests) {
    $scriptPath = Join-Path $PSScriptRoot $test.Script
    
    if (Test-Path $scriptPath) {
        Run-Test -TestName $test.Name -TestScript $scriptPath -Description $test.Description
    } else {
        Write-Host "`n⚠️  テストスクリプトが見つかりません: $scriptPath" -ForegroundColor Yellow
        
        $result = @{
            Name = $test.Name
            Status = "SKIPPED"
            Duration = [TimeSpan]::Zero
            Error = "テストスクリプトが見つかりません"
        }
        $testResults += $result
    }
}

# テスト結果サマリー
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "テスト結果サマリー" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Cyan

$endTime = Get-Date
$totalDuration = $endTime - $startTime

$successCount = ($testResults | Where-Object { $_.Status -eq "SUCCESS" }).Count
$failedCount = ($testResults | Where-Object { $_.Status -eq "FAILED" }).Count
$skippedCount = ($testResults | Where-Object { $_.Status -eq "SKIPPED" }).Count
$totalCount = $testResults.Count

Write-Host "総実行時間: $($totalDuration.TotalSeconds.ToString('F2'))秒" -ForegroundColor Yellow
Write-Host "総テスト数: $totalCount" -ForegroundColor White
Write-Host "成功: $successCount" -ForegroundColor Green
Write-Host "失敗: $failedCount" -ForegroundColor Red
Write-Host "スキップ: $skippedCount" -ForegroundColor Yellow

# 詳細結果
Write-Host "`n詳細結果:" -ForegroundColor White
foreach ($result in $testResults) {
    $statusIcon = switch ($result.Status) {
        "SUCCESS" { "✅" }
        "FAILED" { "❌" }
        "SKIPPED" { "⚠️" }
        default { "❓" }
    }
    
    $statusColor = switch ($result.Status) {
        "SUCCESS" { "Green" }
        "FAILED" { "Red" }
        "SKIPPED" { "Yellow" }
        default { "Gray" }
    }
    
    Write-Host "$statusIcon $($result.Name)" -ForegroundColor $statusColor
    Write-Host "  実行時間: $($result.Duration.TotalSeconds.ToString('F2'))秒" -ForegroundColor Gray
    
    if ($result.Error) {
        Write-Host "  エラー: $($result.Error)" -ForegroundColor Red
    }
}

# 最終結果
Write-Host "`n==========================================" -ForegroundColor Cyan
if ($failedCount -eq 0 -and $skippedCount -eq 0) {
    Write-Host "🎉 すべてのテストが成功しました！" -ForegroundColor Green
} elseif ($failedCount -eq 0) {
    Write-Host "✅ 実行されたテストはすべて成功しました（一部スキップあり）" -ForegroundColor Green
} else {
    Write-Host "⚠️  一部のテストが失敗しました" -ForegroundColor Yellow
}
Write-Host "==========================================" -ForegroundColor Cyan

# テスト結果をファイルに保存
$resultsPath = Join-Path $PSScriptRoot "test_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
$testResults | ConvertTo-Json -Depth 3 | Out-File -FilePath $resultsPath -Encoding UTF8
Write-Host "`nテスト結果を保存しました: $resultsPath" -ForegroundColor Gray 