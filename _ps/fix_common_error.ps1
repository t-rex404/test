# Common.ps1 変数初期化エラー修正スクリプト
# "Variable is not assigned in the method." エラーを修正するためのスクリプト

Write-Host "Common.ps1 変数初期化エラー修正を開始します..." -ForegroundColor Green

try {
    # バックアップを作成
    $backupDir = Join-Path $PSScriptRoot "backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    
    # 元のファイルをバックアップ
    $filesToBackup = @(
        "_lib\Common.ps1"
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
    $fixedCommonPath = Join-Path $PSScriptRoot "_lib\Common_fixed.ps1"
    $originalCommonPath = Join-Path $PSScriptRoot "_lib\Common.ps1"
    
    if (Test-Path $fixedCommonPath) {
        Copy-Item $fixedCommonPath $originalCommonPath -Force
        Write-Host "Common.ps1 を修正版で置き換えました" -ForegroundColor Green
    } else {
        Write-Host "修正版ファイルが見つかりません: $fixedCommonPath" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "`n修正が完了しました！" -ForegroundColor Green
    Write-Host "バックアップは以下のディレクトリに保存されています: $backupDir" -ForegroundColor Cyan
    
    # 修正内容の説明
    Write-Host "`n修正内容:" -ForegroundColor White
    Write-Host "1. HandleErrorメソッド内のヒアドキュメント（@\"...\"@）を文字列連結に変更" -ForegroundColor White
    Write-Host "2. 変数を明示的に宣言してから使用するように修正" -ForegroundColor White
    Write-Host "3. PowerShellクラス内での変数スコープ問題を解決" -ForegroundColor White
    
    Write-Host "`nこれで 'Variable is not assigned in the method.' エラーが解決されるはずです。" -ForegroundColor Green
    
    # 修正後のテスト
    Write-Host "`n修正後のテストを実行します..." -ForegroundColor Yellow
    try {
        # Common.ps1をテスト読み込み
        . "$PSScriptRoot\_lib\Common.ps1"
        Write-Host "Common.ps1 の読み込みテストが成功しました！" -ForegroundColor Green
        
        # Commonクラスのテスト
        $testCommon = [Common]::new()
        $testCommon.WriteLog("テストメッセージ", "INFO")
        Write-Host "Commonクラスのテストが成功しました！" -ForegroundColor Green
        
    } catch {
        Write-Host "修正後のテストでエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }
    
} catch {
    Write-Host "修正中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
} 