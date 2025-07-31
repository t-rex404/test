# Excelファイル操作クラス
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# 共通ライブラリをインポート
#. "$PSScriptRoot\Common.ps1"
#$Common = New-Object -TypeName 'Common'

class ExcelDriver
{
    [Microsoft.Office.Interop.Excel.Application]$excel_app
    [Microsoft.Office.Interop.Excel.Workbook]$excel_workbook
    [Microsoft.Office.Interop.Excel.Worksheet]$excel_worksheet
    [string]$file_path
    [bool]$is_initialized
    [bool]$is_saved
    [string]$temp_directory

    # ========================================
    # 初期化・接続関連
    # ========================================

    ExcelDriver()
    {
        try
        {
            $this.is_initialized = $false
            $this.is_saved = $false
            
            # 一時ディレクトリの作成
            $this.temp_directory = $this.CreateTempDirectory()
            
            # Excelアプリケーションの初期化
            $this.InitializeExcelApplication()
            
            # 新規ワークブックの作成
            $this.CreateNewWorkbook()
            
            $this.is_initialized = $true
            Write-Host "ExcelDriverの初期化が完了しました。"
        }
        catch
        {
            $this.CleanupOnInitializationFailure()
            $global:Common.HandleError("5001", "ExcelDriver初期化エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "ExcelDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 一時ディレクトリを作成
    [string] CreateTempDirectory()
    {
        try
        {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $temp_dir = Join-Path $env:TEMP "ExcelDriver_$timestamp"
            if (-not (Test-Path $temp_dir))
            {
                New-Item -ItemType Directory -Path $temp_dir -Force | Out-Null
            }
            
            Write-Host "一時ディレクトリを作成しました: $temp_dir"
            return $temp_dir
        }
        catch
        {
            $global:Common.HandleError("5002", "一時ディレクトリ作成エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "一時ディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # Excelアプリケーションを初期化
    [void] InitializeExcelApplication()
    {
        try
        {
            $this.excel_app = New-Object -ComObject Excel.Application
            $this.excel_app.Visible = $false
            $this.excel_app.DisplayAlerts = $false
            
            Write-Host "Excelアプリケーションを初期化しました。"
        }
        catch
        {
            $global:Common.HandleError("5003", "Excelアプリケーション初期化エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "Excelアプリケーションの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 新規ワークブックを作成
    [void] CreateNewWorkbook()
    {
        try
        {
            $this.excel_workbook = $this.excel_app.Workbooks.Add()
            $this.excel_worksheet = $this.excel_workbook.Worksheets.Item(1)
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $this.file_path = Join-Path $this.temp_directory "Workbook_$timestamp.xlsx"
            
            Write-Host "新規ワークブックを作成しました。"
        }
        catch
        {
            $global:Common.HandleError("5004", "新規ワークブック作成エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "新規ワークブックの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # セル操作関連
    # ========================================

    # セルに値を設定
    [void] SetCellValue([string]$cell, [object]$value)
    {
        try
        {
            if ([string]::IsNullOrEmpty($cell))
            {
                throw "セル参照が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $this.excel_worksheet.Range($cell).Value = $value
            Write-Host "セル $cell に値を設定しました: $value"
        }
        catch
        {
            $global:Common.HandleError("5005", "セル値設定エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "セル値の設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # セルの値を取得
    [object] GetCellValue([string]$cell)
    {
        try
        {
            if ([string]::IsNullOrEmpty($cell))
            {
                throw "セル参照が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $value = $this.excel_worksheet.Range($cell).Value
            Write-Host "セル $cell の値を取得しました: $value"
            return $value
        }
        catch
        {
            $global:Common.HandleError("5006", "セル値取得エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "セル値の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # セル範囲に値を設定
    [void] SetRangeValue([string]$range, [object[,]]$values)
    {
        try
        {
            if ([string]::IsNullOrEmpty($range))
            {
                throw "範囲参照が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $this.excel_worksheet.Range($range).Value = $values
            Write-Host "範囲 $range に値を設定しました。"
        }
        catch
        {
            $global:Common.HandleError("5007", "範囲値設定エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "範囲値の設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # セル範囲の値を取得
    [object[,]] GetRangeValue([string]$range)
    {
        try
        {
            if ([string]::IsNullOrEmpty($range))
            {
                throw "範囲参照が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $values = $this.excel_worksheet.Range($range).Value
            Write-Host "範囲 $range の値を取得しました。"
            return $values
        }
        catch
        {
            $global:Common.HandleError("5008", "範囲値取得エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "範囲値の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # フォーマット関連
    # ========================================

    # セルのフォントを設定
    [void] SetCellFont([string]$cell, [string]$fontName, [int]$fontSize)
    {
        try
        {
            if ([string]::IsNullOrEmpty($cell))
            {
                throw "セル参照が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $range = $this.excel_worksheet.Range($cell)
            $range.Font.Name = $fontName
            $range.Font.Size = $fontSize
            
            Write-Host "セル $cell のフォントを設定しました: $fontName, $fontSize"
        }
        catch
        {
            $global:Common.HandleError("5009", "フォント設定エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "フォントの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # セルを太字にする
    [void] SetCellBold([string]$cell, [bool]$isBold = $true)
    {
        try
        {
            if ([string]::IsNullOrEmpty($cell))
            {
                throw "セル参照が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $this.excel_worksheet.Range($cell).Font.Bold = $isBold
            Write-Host "セル $cell の太字設定を変更しました: $isBold"
        }
        catch
        {
            $global:Common.HandleError("5010", "太字設定エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "太字設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # セルの背景色を設定
    [void] SetCellBackgroundColor([string]$cell, [int]$colorIndex)
    {
        try
        {
            if ([string]::IsNullOrEmpty($cell))
            {
                throw "セル参照が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $this.excel_worksheet.Range($cell).Interior.ColorIndex = $colorIndex
            Write-Host "セル $cell の背景色を設定しました: $colorIndex"
        }
        catch
        {
            $global:Common.HandleError("5011", "背景色設定エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "背景色の設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ワークシート操作関連
    # ========================================

    # 新しいワークシートを追加
    [void] AddWorksheet([string]$sheetName)
    {
        try
        {
            if ([string]::IsNullOrEmpty($sheetName))
            {
                throw "ワークシート名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $newSheet = $this.excel_workbook.Worksheets.Add()
            $newSheet.Name = $sheetName
            Write-Host "新しいワークシートを追加しました: $sheetName"
        }
        catch
        {
            $global:Common.HandleError("5012", "ワークシート追加エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "ワークシートの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # ワークシートを選択
    [void] SelectWorksheet([string]$sheetName)
    {
        try
        {
            if ([string]::IsNullOrEmpty($sheetName))
            {
                throw "ワークシート名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $this.excel_worksheet = $this.excel_workbook.Worksheets.Item($sheetName)
            Write-Host "ワークシートを選択しました: $sheetName"
        }
        catch
        {
            $global:Common.HandleError("5013", "ワークシート選択エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "ワークシートの選択に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ファイル操作関連
    # ========================================

    # ワークブックを保存
    [void] SaveWorkbook([string]$filePath = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            if ([string]::IsNullOrEmpty($filePath))
            {
                $filePath = $this.file_path
            }

            $this.excel_workbook.SaveAs($filePath)
            $this.is_saved = $true
            Write-Host "ワークブックを保存しました: $filePath"
        }
        catch
        {
            $global:Common.HandleError("5014", "ワークブック保存エラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "ワークブックの保存に失敗しました: $($_.Exception.Message)"
        }
    }

    # 既存のワークブックを開く
    [void] OpenWorkbook([string]$filePath)
    {
        try
        {
            if ([string]::IsNullOrEmpty($filePath))
            {
                throw "ファイルパスが指定されていません。"
            }

            if (-not (Test-Path $filePath))
            {
                throw "指定されたファイルが存在しません: $filePath"
            }

            if (-not $this.is_initialized)
            {
                throw "ExcelDriverが初期化されていません。"
            }

            $this.excel_workbook = $this.excel_app.Workbooks.Open($filePath)
            $this.excel_worksheet = $this.excel_workbook.Worksheets.Item(1)
            $this.file_path = $filePath
            Write-Host "ワークブックを開きました: $filePath"
        }
        catch
        {
            $global:Common.HandleError("5015", "ワークブック開くエラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            throw "ワークブックを開くのに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # クリーンアップ関連
    # ========================================

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            if ($this.excel_workbook -ne $null)
            {
                $this.excel_workbook.Close($false)
                $this.excel_workbook = $null
            }

            if ($this.excel_app -ne $null)
            {
                $this.excel_app.Quit()
                $this.excel_app = $null
            }

            if (-not [string]::IsNullOrEmpty($this.temp_directory) -and (Test-Path $this.temp_directory))
            {
                Remove-Item $this.temp_directory -Recurse -Force -ErrorAction SilentlyContinue
            }

            Write-Host "初期化失敗時のクリーンアップを開始します。"
            Write-Host "一時ディレクトリを削除しました: $($this.temp_directory)"
            Write-Host "クリーンアップが完了しました。"
        }
        catch
        {
            Write-Host "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # リソースを解放
    [void] Dispose()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                return
            }

            if ($this.excel_workbook -ne $null)
            {
                if (-not $this.is_saved)
                {
                    $this.excel_workbook.Close($false)
                }
                else
                {
                    $this.excel_workbook.Close($true)
                }
                $this.excel_workbook = $null
            }

            if ($this.excel_app -ne $null)
            {
                $this.excel_app.Quit()
                $this.excel_app = $null
            }

            # COMオブジェクトの解放
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.excel_worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.excel_workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.excel_app) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Write-Host "ExcelDriverのリソースを解放しました。"
        }
        catch
        {
            $global:Common.HandleError("5016", "ExcelDriver Disposeエラー: $($_.Exception.Message)", "ExcelDriver", ".\AllDrivers_Error.log")
            Write-Host "リソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
} 
