# Excelファイル操作クラス
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

class ExcelDriver
{
    [Microsoft.Office.Interop.Excel.Application]$excel_app
    [object]$excel_workbook
    [object]$excel_worksheet
    [string]$file_path
    [bool]$is_initialized
    [bool]$is_saved
    [string]$temp_directory

    # ログファイルパス（共有可能）
    static [string]$NormalLogFile = ".\ExcelDriver_$($env:USERNAME)_Normal.log"
    static [string]$ErrorLogFile = ".\ExcelDriver_$($env:USERNAME)_Error.log"

    # ========================================
    # ログユーティリティ
    # ========================================

    # UTF-8 (BOMなし) で追記するヘルパー
    [void] AppendTextNoBom([string]$filePath, [string]$text)
    {
        try
        {
            $encoding = New-Object System.Text.UTF8Encoding($false)
            $directory = Split-Path -Parent $filePath
            if (-not [string]::IsNullOrEmpty($directory) -and -not (Test-Path -LiteralPath $directory))
            {
                New-Item -ItemType Directory -Path $directory -Force -ErrorAction SilentlyContinue | Out-Null
            }
            $streamWriter = New-Object System.IO.StreamWriter($filePath, $true, $encoding)
            try
            {
                $streamWriter.Write($text)
            }
            finally
            {
                $streamWriter.Dispose()
            }
        }
        catch
        {
            Write-Host "ログ書き込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # 情報ログを出力
    [void] LogInfo([string]$message)
    {
        if ($global:Common)
        {
            try
            {
                $global:Common.WriteLog($message, "INFO", "ExcelDriver")
            }
            catch
            {
                Write-Host "正常ログ出力に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([ExcelDriver]::NormalLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Green
    }

    # エラーログを出力
    [void] LogError([string]$errorCode, [string]$message)
    {
        if ($global:Common)
        {
            try
            {
                $global:Common.HandleError($errorCode, $message, "ExcelDriver")
            }
            catch
            {
                Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([ExcelDriver]::ErrorLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Red
    }

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

            # 正常ログ出力
            $this.LogInfo("ExcelDriverの初期化が完了しました")
        }
        catch
        {
            # 初期化失敗時のクリーンアップ
            Write-Host "ExcelDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()

            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0001", "ExcelDriver初期化エラー: $($_.Exception.Message)")
            
            throw "ExcelDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 一時ディレクトリを作成
    [string] CreateTempDirectory()
    {
        try
        {
            $base_dirs = @(
                "C:\temp",
                "$($env:TEMP)"
            )
            $base_dir = $base_dirs[0]
            if (-not (Test-Path $base_dir))
            {
                $base_dir = $base_dirs[1]
            }
            $temp_dir = Join-Path $base_dir "ExcelDriver"
            if (-not (Test-Path $temp_dir))
            {
                New-Item -ItemType Directory -Path $temp_dir -Force | Out-Null
            }
            
            Write-Host "一時ディレクトリを作成しました: $temp_dir"

            # 正常ログ出力
            $this.LogInfo("一時ディレクトリを作成しました: $temp_dir")

            return $temp_dir
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0002", "一時ディレクトリ作成エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("Excelアプリケーションを初期化しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0010", "Excelアプリケーション初期化エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("新規ワークブックを作成しました")
        }
        catch
        {
            $this.LogError("ExcelDriverError_0011", "新規ワークブック作成エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("セル $cell に値を設定しました: $value")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0020", "セル値設定エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("セル $cell の値を取得しました: $value")

            return $value
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0021", "セル値取得エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("範囲 $range に値を設定しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0022", "範囲値設定エラー: $($_.Exception.Message)")
            
            throw "範囲値の設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # セル範囲の値を取得（戻りは可変: 単一値 / 配列 / 2次元配列）
    [object] GetRangeValue([string]$range)
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

            # 正常ログ出力
            $this.LogInfo("範囲 $range の値を取得しました")

            return $values
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0023", "範囲値取得エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("セル $cell のフォントを設定しました: $fontName, $fontSize")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0040", "フォント設定エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("セル $cell の太字設定を変更しました: $isBold")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0041", "太字設定エラー: $($_.Exception.Message)")
            
            throw "太字の設定に失敗しました: $($_.Exception.Message)"
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

            # 正常ログ出力
            $this.LogInfo("セル $cell の背景色を設定しました: $colorIndex")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0042", "背景色設定エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("新しいワークシートを追加しました: $sheetName")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0024", "ワークシート追加エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("ワークシートを選択しました: $sheetName")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0025", "ワークシート選択エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("ワークブックを保存しました: $filePath")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0060", "ワークブック保存エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("ワークブックを開きました: $filePath")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0061", "ワークブック開くエラー: $($_.Exception.Message)")
            
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
            if ($null -ne $this.excel_workbook)
            {
                $this.excel_workbook.Close($false)
                $this.excel_workbook = $null
            }

            if ($null -ne $this.excel_app)
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

            # 正常ログ出力
            $this.LogInfo("初期化失敗時のクリーンアップが完了しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0090", "初期化失敗時のクリーンアップエラー: $($_.Exception.Message)")
            
            throw "初期化失敗時のクリーンアップ中にエラーが発生しました: $($_.Exception.Message)"
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

            if ($null -ne $this.excel_workbook)
            {
                if (-not $this.is_saved)
                {
                    #$this.excel_workbook.Close($false)
                    $this.SaveWorkbook($this.temp_directory + "\ExcelDriver_AutoSave.xlsx")
                }
                else
                {
                    #$this.excel_workbook.Close($true)
                    $this.SaveWorkbook($this.temp_directory + "\ExcelDriver_AutoSave.xlsx")
                }
                $this.excel_workbook = $null
            }

            if ($null -ne $this.excel_app)
            {
                $this.excel_app.Quit()
                $this.excel_app = $null
            }

            # COMオブジェクトの解放（null チェック）
            if ($this.excel_worksheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.excel_worksheet) | Out-Null }
            if ($this.excel_workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.excel_workbook) | Out-Null }
            if ($this.excel_app) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.excel_app) | Out-Null }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Write-Host "ExcelDriverのリソースを解放しました。"

            # 正常ログ出力
            $this.LogInfo("ExcelDriverのリソースを解放しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("ExcelDriverError_0091", "ExcelDriver Disposeエラー: $($_.Exception.Message)")
            
            throw "ExcelDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)"
        }
    }
} 
