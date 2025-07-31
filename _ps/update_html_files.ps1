# HTMLファイルを外部CSS・JavaScriptファイルを使用するように更新するスクリプト

$docsPath = "..\docs"
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
        
        # <style>タグを<link>タグに置換
        $content = $content -replace '<style>.*?</style>', '<link rel="stylesheet" href="./css/styles.css">'
        
        # </body>タグの前に<script>タグを追加
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