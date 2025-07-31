# HTMLファイルの重複スクリプトタグを削除し、パスを修正するスクリプト

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

Write-Host "HTMLファイルのクリーンアップを開始します..." -ForegroundColor Green

foreach ($file in $htmlFiles) {
    $filePath = Join-Path $docsPath $file
    if (Test-Path $filePath) {
        Write-Host "クリーンアップ中: $file" -ForegroundColor Yellow
        
        # ファイルの内容を読み込み
        $content = Get-Content $filePath -Raw -Encoding UTF8
        
        # 重複するscriptタグを削除
        $content = $content -replace '(?s)<script src="\./js/script\.js"></script>.*?<script src="\./js/script\.js"></script>', '<script src="../js/script.js"></script>'
        
        # 単一のscriptタグのパスを修正
        $content = $content -replace '<script src="\./js/script\.js"></script>', '<script src="../js/script.js"></script>'
        
        # CSSのパスを修正
        $content = $content -replace 'href="\./css/styles\.css"', 'href="../css/styles.css"'
        
        # 修正された内容をファイルに書き戻し
        Set-Content -Path $filePath -Value $content -Encoding UTF8
        
        Write-Host "完了: $file" -ForegroundColor Green
    } else {
        Write-Host "ファイルが見つかりません: $file" -ForegroundColor Red
    }
}

Write-Host "`nすべてのHTMLファイルのクリーンアップが完了しました。" -ForegroundColor Green 