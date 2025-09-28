# PowerPointファイル操作クラス
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

class PowerPointDriver
{
    [Microsoft.Office.Interop.PowerPoint.Application]$powerpoint_app
    [object]$powerpoint_presentation
    [object]$powerpoint_slide
    [string]$file_path
    [bool]$is_initialized
    [bool]$is_saved
    [string]$temp_directory

    # ========================================
    # 初期化・接続関連
    # ========================================

    PowerPointDriver()
    {
        try
        {
            $this.is_initialized = $false
            $this.is_saved = $false
            
            # 一時ディレクトリの作成
            $this.temp_directory = $this.CreateTempDirectory()
            
            # PowerPointアプリケーションの初期化
            $this.InitializePowerPointApplication()
            
            # 新規プレゼンテーションの作成
            $this.CreateNewPresentation()
            
            $this.is_initialized = $true
            Write-Host "PowerPointDriverの初期化が完了しました。"
        }
        catch
        {
            # 初期化失敗時のクリーンアップ
            Write-Host "PowerPointDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()

            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0001", "PowerPointDriver初期化エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "PowerPointDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "PowerPointDriverの初期化に失敗しました: $($_.Exception.Message)"
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
            $temp_dir = Join-Path $base_dir "PowerPointDriver"
            if (-not (Test-Path $temp_dir))
            {
                New-Item -ItemType Directory -Path $temp_dir -Force | Out-Null
            }
            
            Write-Host "一時ディレクトリを作成しました: $temp_dir"
            return $temp_dir
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0002", "一時ディレクトリ作成エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "一時ディレクトリの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "一時ディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # PowerPointアプリケーションを初期化
    [void] InitializePowerPointApplication()
    {
        try
        {
            $this.powerpoint_app = New-Object -ComObject PowerPoint.Application
            $this.powerpoint_app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $this.powerpoint_app.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
            
            Write-Host "PowerPointアプリケーションを初期化しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0010", "PowerPointアプリケーション初期化エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "PowerPointアプリケーションの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "PowerPointアプリケーションの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 新規プレゼンテーションを作成
    [void] CreateNewPresentation()
    {
        try
        {
            $this.powerpoint_presentation = $this.powerpoint_app.Presentations.Add()
            $this.powerpoint_slide = $this.powerpoint_presentation.Slides.Add(1, 1) # タイトルスライド
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $this.file_path = Join-Path $this.temp_directory "Presentation_$timestamp.pptx"
            
            Write-Host "新規プレゼンテーションを作成しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0011", "新規プレゼンテーション作成エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "新規プレゼンテーションの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "新規プレゼンテーションの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # スライド操作関連
    # ========================================

    # 新しいスライドを追加
    [void] AddSlide([int]$layoutType = 1)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            $slideCount = $this.powerpoint_presentation.Slides.Count
            $this.powerpoint_slide = $this.powerpoint_presentation.Slides.Add($slideCount + 1, $layoutType)
            Write-Host "新しいスライドを追加しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0020", "スライド追加エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "スライドの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "スライドの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # スライドを選択
    [void] SelectSlide([int]$slideIndex)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            if ($slideIndex -lt 1 -or $slideIndex -gt $this.powerpoint_presentation.Slides.Count)
            {
                throw "無効なスライドインデックスです: $slideIndex"
            }

            $this.powerpoint_slide = $this.powerpoint_presentation.Slides.Item($slideIndex)
            Write-Host "スライドを選択しました: $slideIndex"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0021", "スライド選択エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "スライドの選択に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "スライドの選択に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # テキスト操作関連
    # ========================================

    # タイトルを設定
    [void] SetTitle([string]$title)
    {
        try
        {
            if ([string]::IsNullOrEmpty($title))
            {
                throw "タイトルが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            $this.powerpoint_slide.Shapes.Title.TextFrame.TextRange.Text = $title
            Write-Host "タイトルを設定しました: $title"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0022", "タイトル設定エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "タイトルの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "タイトルの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # テキストボックスを追加
    [void] AddTextBox([string]$text, [int]$left = 100, [int]$top = 100, [int]$width = 400, [int]$height = 100)
    {
        try
        {
            if ([string]::IsNullOrEmpty($text))
            {
                throw "テキストが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            $textBox = $this.powerpoint_slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
            $textBox.TextFrame.TextRange.Text = $text
            Write-Host "テキストボックスを追加しました: $text"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0023", "テキストボックス追加エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テキストボックスの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "テキストボックスの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # テキストを追加
    [void] AddText([string]$text, [int]$left = 100, [int]$top = 100)
    {
        try
        {
            if ([string]::IsNullOrEmpty($text))
            {
                throw "テキストが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            $textShape = $this.powerpoint_slide.Shapes.AddTextbox(1, $left, $top, 400, 100)
            $textShape.TextFrame.TextRange.Text = $text
            Write-Host "テキストを追加しました: $text"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0024", "テキスト追加エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テキストの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "テキストの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 図形操作関連
    # ========================================

    # 図形を追加
    [void] AddShape([int]$shapeType, [int]$left = 100, [int]$top = 100, [int]$width = 100, [int]$height = 100)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            $shape = $this.powerpoint_slide.Shapes.AddShape($shapeType, $left, $top, $width, $height)
            Write-Host "図形を追加しました: タイプ $shapeType"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0025", "図形追加エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "図形の追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "図形の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 画像を追加
    [void] AddPicture([string]$imagePath, [int]$left = 100, [int]$top = 100, [int]$width = 200, [int]$height = 150)
    {
        try
        {
            if ([string]::IsNullOrEmpty($imagePath))
            {
                throw "画像パスが指定されていません。"
            }

            if (-not (Test-Path $imagePath))
            {
                throw "指定された画像ファイルが存在しません: $imagePath"
            }

            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            $picture = $this.powerpoint_slide.Shapes.AddPicture($imagePath, $false, $true, $left, $top, $width, $height)
            Write-Host "画像を追加しました: $imagePath"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0026", "画像追加エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "画像の追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "画像の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # フォーマット関連
    # ========================================

    # フォントを設定
    [void] SetFont([string]$fontName, [int]$fontSize)
    {
        try
        {
            if ([string]::IsNullOrEmpty($fontName))
            {
                throw "フォント名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            # 現在のスライドのすべてのテキストにフォントを適用
            foreach ($shape in $this.powerpoint_slide.Shapes)
            {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText)
                {
                    $shape.TextFrame.TextRange.Font.Name = $fontName
                    $shape.TextFrame.TextRange.Font.Size = $fontSize
                }
            }
            
            Write-Host "フォントを設定しました: $fontName, $fontSize"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0040", "フォント設定エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "フォントの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "フォントの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 背景色を設定
    [void] SetBackgroundColor([int]$colorIndex)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            $this.powerpoint_slide.Background.Fill.ForeColor.RGB = $colorIndex
            Write-Host "背景色を設定しました: $colorIndex"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0041", "背景色設定エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "背景色の設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "背景色の設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ファイル操作関連
    # ========================================

    # プレゼンテーションを保存
    [void] SavePresentation([string]$filePath = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "PowerPointDriverが初期化されていません。"
            }

            if ([string]::IsNullOrEmpty($filePath))
            {
                $filePath = $this.file_path
            }

            $this.powerpoint_presentation.SaveAs($filePath)
            $this.is_saved = $true
            Write-Host "プレゼンテーションを保存しました: $filePath"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0060", "プレゼンテーション保存エラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "プレゼンテーションの保存に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "プレゼンテーションの保存に失敗しました: $($_.Exception.Message)"
        }
    }

    # 既存のプレゼンテーションを開く
    [void] OpenPresentation([string]$filePath)
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
                throw "PowerPointDriverが初期化されていません。"
            }

            $this.powerpoint_presentation = $this.powerpoint_app.Presentations.Open($filePath)
            $this.powerpoint_slide = $this.powerpoint_presentation.Slides.Item(1)
            $this.file_path = $filePath
            Write-Host "プレゼンテーションを開きました: $filePath"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0061", "プレゼンテーション開くエラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "プレゼンテーションを開くのに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "プレゼンテーションを開くのに失敗しました: $($_.Exception.Message)"
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
            if ($null -ne $this.powerpoint_presentation)
            {
                $this.powerpoint_presentation.Close()
                $this.powerpoint_presentation = $null
            }

            if ($null -ne $this.powerpoint_app)
            {
                $this.powerpoint_app.Quit()
                $this.powerpoint_app = $null
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0090", "初期化失敗時のクリーンアップエラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "初期化失敗時のクリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            }

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

            if ($null -ne $this.powerpoint_presentation)
            {
                # 既に保存済みでも再保存は行わない（読み取り専用エラー回避）
                #$this.powerpoint_presentation.Close()
                $this.SavePresentation($this.temp_directory + "\PowerPointDriver_AutoSave.pptx")
                $this.powerpoint_presentation = $null
            }

            if ($null -ne $this.powerpoint_app)
            {
                $this.powerpoint_app.Quit()
                $this.powerpoint_app = $null
            }

            # COMオブジェクトの解放（null チェック）
            if ($this.powerpoint_slide) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.powerpoint_slide) | Out-Null }
            if ($this.powerpoint_presentation) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.powerpoint_presentation) | Out-Null }
            if ($this.powerpoint_app) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.powerpoint_app) | Out-Null }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Write-Host "PowerPointDriverのリソースを解放しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("PowerPointDriverError_0091", "PowerPointDriver Disposeエラー: $($_.Exception.Message)", "PowerPointDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "リソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "リソース解放中にエラーが発生しました: $($_.Exception.Message)"
        }
    }
} 
