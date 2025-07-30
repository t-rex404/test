# 変数初期化エラー修正スクリプト
# "Variable is not assigned in the method." エラーを修正するためのスクリプト

Write-Host "変数初期化エラー修正を開始します..." -ForegroundColor Green

try {
    # バックアップを作成
    $backupDir = Join-Path $PSScriptRoot "backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    
    # 元のファイルをバックアップ
    $filesToBackup = @(
        "_lib\ChromeDriver.ps1",
        "_lib\EdgeDriver.ps1"
    )
    
    foreach ($file in $filesToBackup) {
        $sourcePath = Join-Path $PSScriptRoot $file
        $backupPath = Join-Path $backupDir (Split-Path $file -Leaf)
        
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath $backupPath -Force
            Write-Host "バックアップ作成: $file" -ForegroundColor Yellow
        }
    }
    
    # 修正されたファイルを元のファイルにコピー
    $fixedChromePath = Join-Path $PSScriptRoot "_lib\ChromeDriver_fixed.ps1"
    $originalChromePath = Join-Path $PSScriptRoot "_lib\ChromeDriver.ps1"
    
    if (Test-Path $fixedChromePath) {
        Copy-Item $fixedChromePath $originalChromePath -Force
        Write-Host "ChromeDriver.ps1 を修正版で置き換えました" -ForegroundColor Green
    }
    
    $fixedEdgePath = Join-Path $PSScriptRoot "_lib\EdgeDriver_fixed.ps1"
    $originalEdgePath = Join-Path $PSScriptRoot "_lib\EdgeDriver.ps1"
    
    if (Test-Path $fixedEdgePath) {
        Copy-Item $fixedEdgePath $originalEdgePath -Force
        Write-Host "EdgeDriver.ps1 を修正版で置き換えました" -ForegroundColor Green
    }
    
    # エラーモジュールが存在しない場合の対処
    $errorModules = @(
        "ChromeDriverErrors.psm1",
        "EdgeDriverErrors.psm1",
        "WordDriverErrors.psm1"
    )
    
    foreach ($module in $errorModules) {
        $modulePath = Join-Path $PSScriptRoot "_lib\$module"
        if (-not (Test-Path $modulePath)) {
            Write-Host "警告: エラーモジュールが見つかりません: $module" -ForegroundColor Yellow
            Write-Host "このモジュールは後で作成する必要があります" -ForegroundColor Yellow
        }
    }
    
    Write-Host "`n修正が完了しました！" -ForegroundColor Green
    Write-Host "バックアップは以下のディレクトリに保存されています: $backupDir" -ForegroundColor Cyan
    
    # 修正内容の説明
    Write-Host "`n修正内容:" -ForegroundColor White
    Write-Host "1. ChromeDriverとEdgeDriverクラスのコンストラクタで親クラスのコンストラクタを正しく呼び出すように修正" -ForegroundColor White
    Write-Host "2. クラスプロパティの初期化を明示的に行うように修正" -ForegroundColor White
    Write-Host "3. Disposeメソッドの呼び出し方を修正" -ForegroundColor White
    
    Write-Host "`nこれで 'Variable is not assigned in the method.' エラーが解決されるはずです。" -ForegroundColor Green
    
} catch {
    Write-Host "修正中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
} 