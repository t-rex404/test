# WordDriverクラステストファイル
# WordDriverクラスの各メソッドの機能をテストするためのスクリプト

# 必要なライブラリをインポート
. $PSScriptRoot\_lib\WordDriver.ps1

Write-Host "WordDriverクラステストを開始します..." -ForegroundColor Green

try {
    # WordDriverインスタンスを作成
    Write-Host "WordDriverを初期化中..." -ForegroundColor Yellow
    $wordDriver = [WordDriver]::new()
    
    # 基本テキスト追加テスト
    Write-Host "`n=== 基本テキスト追加テスト ===" -ForegroundColor Cyan
    $wordDriver.AddText("これはテスト用のテキストです。")
    Write-Host "基本テキスト追加完了" -ForegroundColor Green
    
    # 見出し追加テスト
    Write-Host "`n=== 見出し追加テスト ===" -ForegroundColor Cyan
    $wordDriver.AddHeading("テスト見出し1", 1)
    $wordDriver.AddHeading("テスト見出し2", 2)
    $wordDriver.AddHeading("テスト見出し3", 3)
    Write-Host "見出し追加完了" -ForegroundColor Green
    
    # 段落追加テスト
    Write-Host "`n=== 段落追加テスト ===" -ForegroundColor Cyan
    $wordDriver.AddParagraph("これはテスト用の段落です。PowerShellを使用してWordドキュメントを操作しています。")
    $wordDriver.AddParagraph("2番目の段落です。様々な機能をテストしています。")
    Write-Host "段落追加完了" -ForegroundColor Green
    
    # フォント設定テスト
    Write-Host "`n=== フォント設定テスト ===" -ForegroundColor Cyan
    $wordDriver.SetFont("MS Gothic", 12)
    Write-Host "フォント設定完了" -ForegroundColor Green
    
    # ページ区切りテスト
    Write-Host "`n=== ページ区切りテスト ===" -ForegroundColor Cyan
    $wordDriver.AddPageBreak()
    Write-Host "ページ区切り追加完了" -ForegroundColor Green
    
    # 2ページ目のコンテンツ
    $wordDriver.AddHeading("2ページ目の見出し", 1)
    $wordDriver.AddParagraph("これは2ページ目のコンテンツです。")
    
    # テーブル追加テスト
    Write-Host "`n=== テーブル追加テスト ===" -ForegroundColor Cyan
    $tableData = @(
        @("項目", "値", "説明"),
        @("テスト1", "100", "最初のテスト項目"),
        @("テスト2", "200", "2番目のテスト項目"),
        @("テスト3", "300", "3番目のテスト項目")
    )
    $wordDriver.AddTable($tableData, "テストテーブル")
    Write-Host "テーブル追加完了" -ForegroundColor Green
    
    # 目次追加テスト
    Write-Host "`n=== 目次追加テスト ===" -ForegroundColor Cyan
    $wordDriver.AddTableOfContents()
    Write-Host "目次追加完了" -ForegroundColor Green
    
    # 目次更新テスト
    Write-Host "`n=== 目次更新テスト ===" -ForegroundColor Cyan
    $wordDriver.UpdateTableOfContents()
    Write-Host "目次更新完了" -ForegroundColor Green
    
    # ドキュメント保存テスト
    Write-Host "`n=== ドキュメント保存テスト ===" -ForegroundColor Cyan
    $savePath = Join-Path $PSScriptRoot "test_document.docx"
    $wordDriver.SaveDocument($savePath)
    Write-Host "ドキュメント保存完了: $savePath" -ForegroundColor Green
    
    # ドキュメント開くテスト
    Write-Host "`n=== ドキュメント開くテスト ===" -ForegroundColor Cyan
    $wordDriver.OpenDocument($savePath)
    Write-Host "ドキュメント開く完了" -ForegroundColor Green
    
    # 追加コンテンツ
    $wordDriver.AddHeading("開いたドキュメントに追加", 1)
    $wordDriver.AddParagraph("既存のドキュメントを開いて、新しいコンテンツを追加しました。")
    
    # 最終保存
    $finalSavePath = Join-Path $PSScriptRoot "test_document_final.docx"
    $wordDriver.SaveDocument($finalSavePath)
    Write-Host "最終ドキュメント保存完了: $finalSavePath" -ForegroundColor Green
    
    Write-Host "`nすべてのテストが正常に完了しました！" -ForegroundColor Green
    Write-Host "生成されたファイル:" -ForegroundColor Yellow
    Write-Host "- $savePath" -ForegroundColor White
    Write-Host "- $finalSavePath" -ForegroundColor White
    
} catch {
    Write-Host "テスト中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    # リソースのクリーンアップ
    if ($wordDriver) {
        Write-Host "`nリソースをクリーンアップ中..." -ForegroundColor Yellow
        $wordDriver.Dispose()
    }
    
    Write-Host "テスト終了" -ForegroundColor Green
} 