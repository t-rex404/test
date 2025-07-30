# 統合エラー修正スクリプト
# すべての "Variable is not assigned in the method." エラーを修正するためのスクリプト

Write-Host "統合エラー修正を開始します..." -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan

try {
    # バックアップを作成
    $backupDir = Join-Path $PSScriptRoot "backup_all_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    Write-Host "バックアップディレクトリを作成しました: $backupDir" -ForegroundColor Yellow
    
    # 修正対象ファイルのリスト
    $filesToFix = @(
        @{
            Original = "_lib\Common.ps1"
            Fixed = "_lib\Common_fixed.ps1"
            Description = "Common.ps1 - ヒアドキュメント変数スコープ問題"
        },
        @{
            Original = "_lib\ChromeDriver.ps1"
            Fixed = "_lib\ChromeDriver_fixed.ps1"
            Description = "ChromeDriver.ps1 - 親クラスコンストラクタ呼び出し問題"
        },
        @{
            Original = "_lib\EdgeDriver.ps1"
            Fixed = "_lib\EdgeDriver_fixed.ps1"
            Description = "EdgeDriver.ps1 - 親クラスコンストラクタ呼び出し問題"
        }
    )
    
    $successCount = 0
    $totalCount = $filesToFix.Count
    
    foreach ($file in $filesToFix) {
        $originalPath = Join-Path $PSScriptRoot $file.Original
        $fixedPath = Join-Path $PSScriptRoot $file.Fixed
        $backupPath = Join-Path $backupDir (Split-Path $file.Original -Leaf)
        
        Write-Host "`n処理中: $($file.Description)" -ForegroundColor White
        
        # 元のファイルが存在するかチェック
        if (-not (Test-Path $originalPath)) {
            Write-Host "警告: 元のファイルが見つかりません: $($file.Original)" -ForegroundColor Yellow
            continue
        }
        
        # 修正版ファイルが存在するかチェック
        if (-not (Test-Path $fixedPath)) {
            Write-Host "警告: 修正版ファイルが見つかりません: $($file.Fixed)" -ForegroundColor Yellow
            continue
        }
        
        try {
            # バックアップを作成
            Copy-Item $originalPath $backupPath -Force
            Write-Host "  バックアップ作成: $($file.Original)" -ForegroundColor Gray
            
            # 修正版を適用
            Copy-Item $fixedPath $originalPath -Force
            Write-Host "  修正版を適用: $($file.Original)" -ForegroundColor Green
            
            $successCount++
            
        } catch {
            Write-Host "  エラー: $($file.Original) の修正に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # エラーモジュールの確認
    Write-Host "`nエラーモジュールの確認..." -ForegroundColor White
    $errorModules = @(
        "ChromeDriverErrors.psm1",
        "EdgeDriverErrors.psm1",
        "WordDriverErrors.psm1"
    )
    
    foreach ($module in $errorModules) {
        $modulePath = Join-Path $PSScriptRoot "_lib\$module"
        if (Test-Path $modulePath) {
            Write-Host "  ✓ $module が見つかりました" -ForegroundColor Green
        } else {
            Write-Host "  ✗ $module が見つかりません" -ForegroundColor Red
        }
    }
    
    # 結果サマリー
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "修正完了サマリー" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "成功: $successCount / $totalCount ファイル" -ForegroundColor Green
    Write-Host "バックアップ: $backupDir" -ForegroundColor Cyan
    
    if ($successCount -eq $totalCount) {
        Write-Host "`nすべての修正が完了しました！" -ForegroundColor Green
    } else {
        Write-Host "`n一部の修正が失敗しました。手動で確認してください。" -ForegroundColor Yellow
    }
    
    # 修正内容の説明
    Write-Host "`n修正内容:" -ForegroundColor White
    Write-Host "1. Common.ps1: ヒアドキュメント（@\"...\"@）を文字列連結に変更" -ForegroundColor White
    Write-Host "2. ChromeDriver.ps1: 親クラスコンストラクタの正しい呼び出し" -ForegroundColor White
    Write-Host "3. EdgeDriver.ps1: 親クラスコンストラクタの正しい呼び出し" -ForegroundColor White
    Write-Host "4. プロパティの明示的初期化" -ForegroundColor White
    Write-Host "5. Disposeメソッドの呼び出し修正" -ForegroundColor White
    
    # 修正後のテスト
    Write-Host "`n修正後のテストを実行します..." -ForegroundColor Yellow
    try {
        # Common.ps1のテスト
        . "$PSScriptRoot\_lib\Common.ps1"
        Write-Host "✓ Common.ps1 の読み込みテスト成功" -ForegroundColor Green
        
        # エラーモジュールのテスト
        Import-Module "$PSScriptRoot\_lib\ChromeDriverErrors.psm1" -Force
        Import-Module "$PSScriptRoot\_lib\EdgeDriverErrors.psm1" -Force
        Import-Module "$PSScriptRoot\_lib\WordDriverErrors.psm1" -Force
        Write-Host "✓ エラーモジュールの読み込みテスト成功" -ForegroundColor Green
        
        Write-Host "`nすべてのテストが成功しました！" -ForegroundColor Green
        Write-Host "これで 'Variable is not assigned in the method.' エラーが解決されるはずです。" -ForegroundColor Green
        
    } catch {
        Write-Host "修正後のテストでエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "手動でテストを実行してください。" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "修正中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
}

Write-Host "`n修正スクリプトが完了しました。" -ForegroundColor Cyan 