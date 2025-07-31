# エラーログファイル名を統一するスクリプト
# すべてのドライバーで共通のログファイル名を使用するように変更

$files = @(
    "_lib/WebDriver.ps1",
    "_lib/EdgeDriver.ps1", 
    "_lib/ChromeDriver.ps1",
    "_lib/WordDriver.ps1",
    "_lib/ExcelDriver.ps1",
    "_lib/PowerPointDriver.ps1"
)

$unifiedLogFile = ".\AllDrivers_Error.log"

foreach ($file in $files) {
    if (Test-Path $file) {
        Write-Host "修正中: $file"
        
        # ファイルの内容を読み込み
        $content = Get-Content $file -Raw
        
        # 各ドライバー固有のログファイル名を統一されたログファイル名に置換
        $content = $content -replace '\.\\WebDriver_Error\.log', $unifiedLogFile
        $content = $content -replace '\.\\EdgeDriver_Error\.log', $unifiedLogFile
        $content = $content -replace '\.\\ChromeDriver_Error\.log', $unifiedLogFile
        $content = $content -replace '\.\\WordDriver_Error\.log', $unifiedLogFile
        $content = $content -replace '\.\\ExcelDriver_Error\.log', $unifiedLogFile
        $content = $content -replace '\.\\PowerPointDriver_Error\.log', $unifiedLogFile
        
        # 修正された内容をファイルに書き込み
        Set-Content $file $content -Encoding UTF8
        
        Write-Host "完了: $file"
    } else {
        Write-Host "ファイルが見つかりません: $file"
    }
}

Write-Host "`nエラーログファイル名の統一が完了しました。"
Write-Host "統一されたログファイル名: $unifiedLogFile" 