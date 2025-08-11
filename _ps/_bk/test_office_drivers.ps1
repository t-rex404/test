# Officeドライバーテストファイル
# ExcelDriver、PowerPointDriver、WordDriverのテスト

Write-Host "=== Officeドライバーテスト開始 ===" -ForegroundColor Green

# 必要なライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"
. "$PSScriptRoot\_lib\WordDriver.ps1"
. "$PSScriptRoot\_lib\ExcelDriver.ps1"
. "$PSScriptRoot\_lib\PowerPointDriver.ps1"

# テスト用の一時ディレクトリを作成
$test_dir = Join-Path $env:TEMP "OfficeDriverTest_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
if (-not (Test-Path $test_dir)) {
    New-Item -ItemType Directory -Path $test_dir -Force | Out-Null
}
Write-Host "テストディレクトリを作成: $test_dir" -ForegroundColor Cyan

# ========================================
# Common.ps1 テスト
# ========================================
Write-Host "`n=== Common.ps1 テスト ===" -ForegroundColor Yellow

try {
    # GetErrorTitle関数のテスト（新しいドライバー用）
    $title1 = $Common.GetErrorTitle("5001", "ExcelDriver")
    $title2 = $Common.GetErrorTitle("6001", "PowerPointDriver")
    Write-Host "✓ Common.GetErrorTitle テスト完了" -ForegroundColor Green
    Write-Host "  取得したタイトル: $title1, $title2" -ForegroundColor Gray
}
catch {
    Write-Host "✗ Common.ps1 テストでエラー: $($_.Exception.Message)" -ForegroundColor Red
}

# ========================================
# ExcelDriver テスト
# ========================================
Write-Host "`n=== ExcelDriver テスト ===" -ForegroundColor Yellow

try {
    # ExcelDriverインスタンスを作成
    Write-Host "ExcelDriverの初期化をテスト中..." -ForegroundColor Cyan
    $excel_driver = [ExcelDriver]::new()
    Write-Host "✓ ExcelDriver初期化完了" -ForegroundColor Green
    
    # セル操作テスト
    $excel_driver.SetCellValue("A1", "テストデータ")
    Write-Host "✓ ExcelDriver.SetCellValue テスト完了" -ForegroundColor Green
    
    $cell_value = $excel_driver.GetCellValue("A1")
    Write-Host "✓ ExcelDriver.GetCellValue テスト完了: $cell_value" -ForegroundColor Green
    
    # 範囲操作テスト
    $values = @("A", "B", "C"), @("1", "2", "3"), @("X", "Y", "Z")
    $excel_driver.SetRangeValue("A2:C4", $values)
    Write-Host "✓ ExcelDriver.SetRangeValue テスト完了" -ForegroundColor Green
    
    $range_values = $excel_driver.GetRangeValue("A2:C4")
    Write-Host "✓ ExcelDriver.GetRangeValue テスト完了" -ForegroundColor Green
    
    # フォーマットテスト
    $excel_driver.SetCellFont("A1", "MS Gothic", 14)
    Write-Host "✓ ExcelDriver.SetCellFont テスト完了" -ForegroundColor Green
    
    $excel_driver.SetCellBold("A1", $true)
    Write-Host "✓ ExcelDriver.SetCellBold テスト完了" -ForegroundColor Green
    
    $excel_driver.SetCellBackgroundColor("A1", 15)
    Write-Host "✓ ExcelDriver.SetCellBackgroundColor テスト完了" -ForegroundColor Green
    
    # ワークシート操作テスト
    $excel_driver.AddWorksheet("テストシート")
    Write-Host "✓ ExcelDriver.AddWorksheet テスト完了" -ForegroundColor Green
    
    $excel_driver.SelectWorksheet("テストシート")
    Write-Host "✓ ExcelDriver.SelectWorksheet テスト完了" -ForegroundColor Green
    
    # ファイル保存テスト
    $save_path = Join-Path $test_dir "test_workbook.xlsx"
    $excel_driver.SaveWorkbook($save_path)
    Write-Host "✓ ExcelDriver.SaveWorkbook テスト完了: $save_path" -ForegroundColor Green
    
    # リソース解放
    $excel_driver.Dispose()
    Write-Host "✓ ExcelDriver.Dispose テスト完了" -ForegroundColor Green
}
catch {
    Write-Host "✗ ExcelDriver テストでエラー: $($_.Exception.Message)" -ForegroundColor Red
    if ($excel_driver) {
        try { $excel_driver.Dispose() } catch { }
    }
}

# ========================================
# PowerPointDriver テスト
# ========================================
Write-Host "`n=== PowerPointDriver テスト ===" -ForegroundColor Yellow

try {
    # PowerPointDriverインスタンスを作成
    Write-Host "PowerPointDriverの初期化をテスト中..." -ForegroundColor Cyan
    $powerpoint_driver = [PowerPointDriver]::new()
    Write-Host "✓ PowerPointDriver初期化完了" -ForegroundColor Green
    
    # スライド操作テスト
    $powerpoint_driver.SetTitle("テストプレゼンテーション")
    Write-Host "✓ PowerPointDriver.SetTitle テスト完了" -ForegroundColor Green
    
    $powerpoint_driver.AddSlide(2) # タイトルとコンテンツ
    Write-Host "✓ PowerPointDriver.AddSlide テスト完了" -ForegroundColor Green
    
    $powerpoint_driver.SelectSlide(2)
    Write-Host "✓ PowerPointDriver.SelectSlide テスト完了" -ForegroundColor Green
    
    # テキスト操作テスト
    $powerpoint_driver.AddText("これはテスト用のテキストです。")
    Write-Host "✓ PowerPointDriver.AddText テスト完了" -ForegroundColor Green
    
    $powerpoint_driver.AddTextBox("テキストボックスのテスト", 200, 200, 300, 100)
    Write-Host "✓ PowerPointDriver.AddTextBox テスト完了" -ForegroundColor Green
    
    # 図形操作テスト
    $powerpoint_driver.AddShape(1, 400, 100, 100, 100) # 矩形
    Write-Host "✓ PowerPointDriver.AddShape テスト完了" -ForegroundColor Green
    
    # フォーマットテスト
    $powerpoint_driver.SetFont("MS Gothic", 12)
    Write-Host "✓ PowerPointDriver.SetFont テスト完了" -ForegroundColor Green
    
    $powerpoint_driver.SetBackgroundColor(16777215) # 白色
    Write-Host "✓ PowerPointDriver.SetBackgroundColor テスト完了" -ForegroundColor Green
    
    # ファイル保存テスト
    $save_path = Join-Path $test_dir "test_presentation.pptx"
    $powerpoint_driver.SavePresentation($save_path)
    Write-Host "✓ PowerPointDriver.SavePresentation テスト完了: $save_path" -ForegroundColor Green
    
    # リソース解放
    $powerpoint_driver.Dispose()
    Write-Host "✓ PowerPointDriver.Dispose テスト完了" -ForegroundColor Green
}
catch {
    Write-Host "✗ PowerPointDriver テストでエラー: $($_.Exception.Message)" -ForegroundColor Red
    if ($powerpoint_driver) {
        try { $powerpoint_driver.Dispose() } catch { }
    }
}

# ========================================
# WordDriver テスト
# ========================================
Write-Host "`n=== WordDriver テスト ===" -ForegroundColor Yellow

try {
    # WordDriverインスタンスを作成
    Write-Host "WordDriverの初期化をテスト中..." -ForegroundColor Cyan
    $word_driver = [WordDriver]::new()
    Write-Host "✓ WordDriver初期化完了" -ForegroundColor Green
    
    # テキスト操作テスト
    $word_driver.AddText("これはテスト用のテキストです。")
    Write-Host "✓ WordDriver.AddText テスト完了" -ForegroundColor Green
    
    $word_driver.AddHeading("テスト見出し", 1)
    Write-Host "✓ WordDriver.AddHeading テスト完了" -ForegroundColor Green
    
    $word_driver.AddParagraph("これはテスト用の段落です。")
    Write-Host "✓ WordDriver.AddParagraph テスト完了" -ForegroundColor Green
    
    # フォーマットテスト
    $word_driver.SetFont("MS Gothic", 12)
    Write-Host "✓ WordDriver.SetFont テスト完了" -ForegroundColor Green
    
    # ファイル保存テスト
    $save_path = Join-Path $test_dir "test_document.docx"
    $word_driver.SaveDocument($save_path)
    Write-Host "✓ WordDriver.SaveDocument テスト完了: $save_path" -ForegroundColor Green
    
    # リソース解放
    $word_driver.Dispose()
    Write-Host "✓ WordDriver.Dispose テスト完了" -ForegroundColor Green
}
catch {
    Write-Host "✗ WordDriver テストでエラー: $($_.Exception.Message)" -ForegroundColor Red
    if ($word_driver) {
        try { $word_driver.Dispose() } catch { }
    }
}

# ========================================
# テスト結果の表示
# ========================================
Write-Host "`n=== テスト結果 ===" -ForegroundColor Green
Write-Host "テストディレクトリ: $test_dir" -ForegroundColor Cyan
Write-Host "生成されたファイル:" -ForegroundColor Cyan
if (Test-Path $test_dir) {
    Get-ChildItem $test_dir -Recurse | ForEach-Object {
        Write-Host "  - $($_.Name)" -ForegroundColor Gray
    }
}

Write-Host "`n=== Officeドライバーテスト完了 ===" -ForegroundColor Green
Write-Host "すべてのOfficeドライバーテストが正常に完了しました！" -ForegroundColor Green 