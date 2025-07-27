# WordDriver使用例
# 必要なライブラリを読み込み
. "$PSScriptRoot\_lib\Common.ps1"
. "$PSScriptRoot\_lib\WordDriver.ps1"

# WordDriverの使用例
try
{
    Write-Host "WordDriverの使用例を開始します..." -ForegroundColor Green
    
    # WordDriverのインスタンスを作成
    $wordDriver = [WordDriver]::new()
    
    # フォントを設定
    $wordDriver.SetFont("MS Gothic", 12, $false, $false)
    
    # タイトルを追加
    $wordDriver.AddHeading("PowerShell自動化レポート", 1)
    $wordDriver.AddParagraph("このドキュメントはPowerShellを使用して自動生成されました。", "Center")
    
    # ページ区切りを追加
    $wordDriver.AddPageBreak()
    
    # 目次を追加
    $wordDriver.AddTableOfContents(3)
    
    # 第1章
    $wordDriver.AddHeading("第1章 概要", 1)
    $wordDriver.AddParagraph("本章では、PowerShellを使用した自動化の概要について説明します。")
    
    $wordDriver.AddHeading("1.1 自動化の目的", 2)
    $wordDriver.AddParagraph("PowerShell自動化の主な目的は以下の通りです：")
    
    # 箇条書きを追加
    $wordDriver.AddText("• 作業の効率化", "Normal")
    $wordDriver.AddText("• ヒューマンエラーの削減", "Normal")
    $wordDriver.AddText("• 一貫性のある処理の実行", "Normal")
    $wordDriver.AddText("• 時間の節約", "Normal")
    
    $wordDriver.AddHeading("1.2 対象システム", 2)
    $wordDriver.AddParagraph("自動化の対象となるシステムは以下の通りです：")
    
    # 表を追加
    $tableData = @(
        "システム名", "説明", "自動化レベル",
        "WebDriver", "Webブラウザ自動化", "高",
        "WordDriver", "Word文書自動化", "中",
        "ExcelDriver", "Excel自動化", "中"
    )
    $wordDriver.AddTable($tableData, 4, 3, "自動化対象システム一覧")
    
    # 第2章
    $wordDriver.AddHeading("第2章 実装詳細", 1)
    $wordDriver.AddParagraph("本章では、各ドライバークラスの実装詳細について説明します。")
    
    $wordDriver.AddHeading("2.1 WebDriverクラス", 2)
    $wordDriver.AddParagraph("WebDriverクラスは、Chrome DevTools Protocolを使用してWebブラウザを制御します。")
    
    $wordDriver.AddHeading("2.1.1 主要機能", 3)
    $wordDriver.AddParagraph("• ブラウザの起動と制御")
    $wordDriver.AddParagraph("• ページのナビゲーション")
    $wordDriver.AddParagraph("• 要素の検索と操作")
    $wordDriver.AddParagraph("• スクリーンショットの取得")
    
    $wordDriver.AddHeading("2.1.2 エラーハンドリング", 3)
    $wordDriver.AddParagraph("包括的なエラーハンドリング機能を実装しており、以下の機能を提供します：")
    $wordDriver.AddText("• try-catchブロックによる例外処理", "Normal")
    $wordDriver.AddText("• リトライ機能", "Normal")
    $wordDriver.AddText("• タイムアウト設定", "Normal")
    $wordDriver.AddText("• 詳細なログ出力", "Normal")
    
    $wordDriver.AddHeading("2.2 WordDriverクラス", 2)
    $wordDriver.AddParagraph("WordDriverクラスは、Microsoft WordのCOMオブジェクトを使用して文書を操作します。")
    
    $wordDriver.AddHeading("2.2.1 主要機能", 3)
    $wordDriver.AddParagraph("• 新規文書の作成")
    $wordDriver.AddParagraph("• テキスト、見出し、段落の追加")
    $wordDriver.AddParagraph("• 表と画像の挿入")
    $wordDriver.AddParagraph("• 目次の自動生成")
    
    # 第3章
    $wordDriver.AddHeading("第3章 使用例", 1)
    $wordDriver.AddParagraph("本章では、実際の使用例を示します。")
    
    $wordDriver.AddHeading("3.1 WebDriver使用例", 2)
    $wordDriver.AddParagraph("WebDriverを使用したWebサイトの自動操作例：")
    
    # コード例を追加（PowerShell 5.1互換）
    $codeExample = @'
# WebDriverの使用例
try {
    $edgeDriver = [EdgeDriver]::new()
    $edgeDriver.Navigate("https://www.example.com")
    
    # 要素を検索して操作
    $element = $edgeDriver.FindElementById("search-box")
    $edgeDriver.SetElementText($element, "検索キーワード")
    
    # スクリーンショットを取得
    $edgeDriver.GetScreenshot("viewPort", "screenshot.png")
}
catch {
    Write-Error "エラー: $($_.Exception.Message)"
}
finally {
    if ($edgeDriver) {
        $edgeDriver.Dispose()
    }
}
'@
    
    $wordDriver.AddParagraph($codeExample, "Left")
    
    $wordDriver.AddHeading("3.2 WordDriver使用例", 2)
    $wordDriver.AddParagraph("WordDriverを使用した文書作成例：")
    
    # コード例を追加（PowerShell 5.1互換）
    $wordCodeExample = @'
# WordDriverの使用例
try {
    $wordDriver = [WordDriver]::new()
    
    # タイトルを追加
    $wordDriver.AddHeading("自動生成レポート", 1)
    
    # 内容を追加
    $wordDriver.AddParagraph("このレポートは自動生成されました。")
    
    # 目次を追加
    $wordDriver.AddTableOfContents(3)
    
    # 保存
    $wordDriver.Save("C:\Reports\report.docx")
}
catch {
    Write-Error "エラー: $($_.Exception.Message)"
}
finally {
    if ($wordDriver) {
        $wordDriver.Dispose()
    }
}
'@
    
    $wordDriver.AddParagraph($wordCodeExample, "Left")
    
    # 第4章
    $wordDriver.AddHeading("第4章 まとめ", 1)
    $wordDriver.AddParagraph("PowerShellを使用した自動化により、効率的なシステム運用が可能になりました。")
    
    $wordDriver.AddHeading("4.1 今後の展望", 2)
    $wordDriver.AddParagraph("今後の開発予定：")
    $wordDriver.AddText("• ExcelDriverクラスの実装", "Normal")
    $wordDriver.AddText("• PDF出力機能の追加", "Normal")
    $wordDriver.AddText("• より高度なエラーハンドリング", "Normal")
    $wordDriver.AddText("• パフォーマンスの最適化", "Normal")
    
    $wordDriver.AddHeading("4.2 注意事項", 2)
    $wordDriver.AddParagraph("使用時の注意事項：")
    $wordDriver.AddText("• 適切なリソース管理（Disposeメソッドの呼び出し）", "Normal")
    $wordDriver.AddText("• エラーハンドリングの実装", "Normal")
    $wordDriver.AddText("• セキュリティの考慮", "Normal")
    $wordDriver.AddText("• パフォーマンスの監視", "Normal")
    
    # ドキュメントを保存
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outputPath = Join-Path $PSScriptRoot "PowerShell自動化レポート_$timestamp.docx"
    $wordDriver.Save($outputPath)
    
    Write-Host "Wordドキュメントが正常に作成されました: $outputPath" -ForegroundColor Green
}
catch
{
    Write-Error "WordDriver使用例でエラーが発生しました: $($_.Exception.Message)"
}
finally
{
    # リソースを解放
    if ($wordDriver)
    {
        $wordDriver.Dispose()
    }
    
    Write-Host "WordDriver使用例が完了しました。" -ForegroundColor Green
} 