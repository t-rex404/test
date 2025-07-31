# HTMLファイルを外部CSS・JavaScriptファイルを使用するように更新するスクリプト（改良版）

$docsPath = "..\docs\pages"
$htmlFiles = @(
    "WebDriver.html",
    "ChromeDriver.html", 
    "EdgeDriver.html",
    "WordDriver.html",
    "ExcelDriver.html",
    "PowerPointDriver.html",
    "Common.html"
)

Write-Host "HTMLファイルの更新を開始します..." -ForegroundColor Green

foreach ($file in $htmlFiles) {
    $filePath = Join-Path $docsPath $file
    if (Test-Path $filePath) {
        Write-Host "更新中: $file" -ForegroundColor Yellow
        
        # ファイルの内容を読み込み
        $content = Get-Content $filePath -Raw -Encoding UTF8
        
        # <style>タグとその内容を<link>タグに置換
        $content = $content -replace '(?s)<style>.*?</style>', '<link rel="stylesheet" href="./css/styles.css">'
        
        # 既存のscriptタグを削除（script.js以外のもの）
        $content = $content -replace '(?s)<script>.*?</script>', ''
        
        # </body>タグの前に正しいscriptタグを追加（まだ存在しない場合）
        if ($content -notmatch '<script src="script\.js"></script>') {
            $content = $content -replace '</body>', '    <script src="./js/script.js"></script>`n</body>'
        }
        
        # 修正された内容をファイルに書き戻し
        Set-Content -Path $filePath -Value $content -Encoding UTF8
        
        Write-Host "完了: $file" -ForegroundColor Green
    } else {
        Write-Host "ファイルが見つかりません: $file" -ForegroundColor Red
    }
}

Write-Host "`nすべてのHTMLファイルの更新が完了しました。" -ForegroundColor Green
Write-Host "外部CSS・JavaScriptファイルを使用するように変更されました。" -ForegroundColor Green 