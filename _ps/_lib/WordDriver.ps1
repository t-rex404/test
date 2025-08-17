# Wordファイル操作クラス
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Word

# WordDriverエラー管理モジュールをインポート
#import-module "$PSScriptRoot\WordDriverErrors.psm1"

# 共通ライブラリをインポート
#. "$PSScriptRoot\Common.ps1"
#$Common = New-Object -TypeName 'Common'

class WordDriver
{
    [Microsoft.Office.Interop.Word.Application]$word_app
    [Microsoft.Office.Interop.Word.Document]$word_document
    [string]$file_path
    [bool]$is_initialized
    [bool]$is_saved
    [string]$temp_directory

    # ========================================
    # 初期化・接続関連
    # ========================================

    WordDriver()
    {
        try
        {
            $this.is_initialized = $false
            $this.is_saved = $false
            
            # 一時ディレクトリの作成
            $this.temp_directory = $this.CreateTempDirectory()
            
            # Wordアプリケーションの初期化
            $this.InitializeWordApplication()
            
            # 新規ドキュメントの作成
            $this.CreateNewDocument()
            
            # フッターに総ページと現在のページを設定
            $this.SetupFooterWithPageNumbers()
            
            $this.is_initialized = $true
            Write-Host "WordDriverの初期化が完了しました。"
        }
        catch
        {
            $this.CleanupOnInitializationFailure()
            $global:Common.HandleError("4001", "WordDriver初期化エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "WordDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 一時ディレクトリを作成
    [string] CreateTempDirectory()
    {
        try
        {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            #$temp_dir = Join-Path $env:TEMP "WordDriver_$timestamp"
            $temp_dir = Join-Path "C:\temp" "WordDriver"
            if (-not (Test-Path $temp_dir))
            {
                New-Item -ItemType Directory -Path $temp_dir -Force | Out-Null
            }
            
            Write-Host "一時ディレクトリを作成しました: $temp_dir"
            return $temp_dir
        }
        catch
        {
            $global:Common.HandleError("4002", "一時ディレクトリ作成エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "一時ディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # Wordアプリケーションを初期化
    [void] InitializeWordApplication()
    {
        try
        {
            $this.word_app = New-Object -ComObject Word.Application
            $this.word_app.Visible = $false
            $this.word_app.DisplayAlerts = $false
            
            Write-Host "Wordアプリケーションを初期化しました。"
        }
        catch
        {
            $global:Common.HandleError("4003", "Wordアプリケーション初期化エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "Wordアプリケーションの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 新規ドキュメントを作成
    [void] CreateNewDocument()
    {
        try
        {
            $this.word_document = $this.word_app.Documents.Add()
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $this.file_path = Join-Path $this.temp_directory "Document_$timestamp.docx"
            
            Write-Host "新規ドキュメントを作成しました。"
        }
        catch
        {
            $global:Common.HandleError("4004", "新規ドキュメント作成エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "新規ドキュメントの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # フッターに総ページと現在のページを設定
    [void] SetupFooterWithPageNumbers()
    {
        try
        {
            # 各セクションのフッターを設定
            foreach ($section in $this.word_document.Sections)
            {
                # フッターを取得
                $footer = $section.Footers.Item([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterPrimary)
                
                # フッターの内容をクリア
                $footer.Range.Text = ""
                
                # 中央揃えでページ番号を挿入
                $footer.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
                
                # 現在のページ番号フィールドを挿入
                $pageField = $footer.Range.Fields.Add($footer.Range, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldPage)
                $pageField.Update()
                
                # " / " を挿入
                $footer.Range.InsertAfter(" / ")
                
                # 総ページ数フィールドを挿入
                $numPagesField = $footer.Range.Fields.Add($footer.Range, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldNumPages)
                $numPagesField.Update()
                
                # フィールドを更新
                $footer.Range.Fields.Update()
                
                # 段落の後に改行を追加
                $footer.Range.InsertParagraphAfter()
            }
            
            # ドキュメント全体のフィールドを更新
            $this.word_document.Fields.Update()
            
            # ドキュメントを再計算
            $this.word_document.Repaginate()
            
            # フィールドの表示を確実にするために少し待機
            Start-Sleep -Milliseconds 100
            
            Write-Host "フッターにページ番号を設定しました（中央揃え）。"
        }
        catch
        {
            $global:Common.HandleError("4018", "フッターページ番号設定エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "フッターのページ番号設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # コンテンツ追加関連
    # ========================================

    # テキストを追加
    [void] AddText([string]$text)
    {
        try
        {
            if ([string]::IsNullOrEmpty($text))
            {
                throw "テキストが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $this.word_document.Content.Text += $text + "`r`n"
            Write-Host "テキストを追加しました: $text"
        }
        catch
        {
            $global:Common.HandleError("4005", "テキスト追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "テキストの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 見出しを追加
    [void] AddHeading([string]$text, [int]$level)
    {
        try
        {
            if ([string]::IsNullOrEmpty($text))
            {
                throw "見出しテキストが指定されていません。"
            }

            if ($level -eq $null)
            {
                $level = 1
            }

            if ($level -lt 1 -or $level -gt 9)
            {
                throw "見出しレベルは1から9の間である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $paragraph = $this.word_document.Content.Paragraphs.Add()
            $paragraph.Range.Text = $text

            # ローカライズ非依存で見出しスタイルを適用（BuiltinStyle を直接代入）
            $builtInStyle = switch ($level)
            {
                1 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading1 }
                2 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2 }
                3 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading3 }
                4 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading4 }
                5 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading5 }
                6 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading6 }
                7 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading7 }
                8 { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading8 }
                default { [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading9 }
            }

            $paragraph.Range.Style = $builtInStyle
            $paragraph.Range.InsertParagraphAfter()
            
            Write-Host "見出しを追加しました: $text (レベル: $level)"
        }
        catch
        {
            $global:Common.HandleError("4006", "見出し追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "見出しの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 段落を追加
    [void] AddParagraph([string]$text)
    {
        try
        {
            if ([string]::IsNullOrEmpty($text))
            {
                throw "段落テキストが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $paragraph = $this.word_document.Content.Paragraphs.Add()
            $paragraph.Range.Text = $text
            $paragraph.Range.InsertParagraphAfter()
            
            Write-Host "段落を追加しました: $text"
        }
        catch
        {
            $global:Common.HandleError("4007", "段落追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "段落の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 表を追加
    [void] AddTable([object[,]]$data, [string]$title)
    {
        try
        {
            if (-not $data)
            {
                throw "テーブルデータが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($title))
            {
                throw "テーブルタイトルが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)

            if ($rows -eq 0 -or $cols -eq 0)
            {
                throw "テーブルデータが空です。"
            }

            # タイトルを追加（指定されている場合）
            if (-not [string]::IsNullOrEmpty($title))
            {
                $this.AddHeading($title, 2)
            }

            # テーブルを作成
            $table = $this.word_document.Tables.Add($this.word_document.Content, $rows, $cols)

            # データを設定
            for ($i = 0; $i -lt $rows; $i++)
            {
                for ($j = 0; $j -lt $cols; $j++)
                {
                    $table.Cell($i + 1, $j + 1).Range.Text = $data[$i, $j].ToString()
                }
            }

            # テーブルの後に段落を追加
            $this.word_document.Content.InsertAfter("`r`n")

            Write-Host "テーブルを追加しました: $rows 行 x $cols 列"
        }
        catch
        {
            $global:Common.HandleError("4008", "テーブル追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "テーブルの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 画像を追加
    [void] AddImage([string]$image_path)
    {
        try
        {
            if ([string]::IsNullOrEmpty($image_path))
            {
                throw "画像パスが指定されていません。"
            }

            if (-not (Test-Path $image_path))
            {
                throw "指定された画像ファイルが見つかりません: $image_path"
            }

            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $this.word_document.InlineShapes.AddPicture($image_path)
            $this.word_document.Content.InsertAfter("`r`n")
            
            Write-Host "画像を追加しました: $image_path"
        }
        catch
        {
            $global:Common.HandleError("4009", "画像追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "画像の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # ページ区切りを追加
    [void] AddPageBreak()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $this.word_document.Content.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
            
            Write-Host "ページ区切りを追加しました。"
        }
        catch
        {
            $global:Common.HandleError("4010", "ページ区切り追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "ページ区切りの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 目次・スタイル関連
    # ========================================

    # 目次を追加
    [void] AddTableOfContents()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            # 目次を挿入する位置を設定
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseStart)

            # 目次を追加（明示的な型で渡して型不一致を回避）
            $this.word_document.TablesOfContents.Add(
                $range,
                $true,   # UseHeadingStyles [bool]
                1,       # UpperHeadingLevel [int]
                3,       # LowerHeadingLevel [int]
                $false,  # UseFields [bool]
                "",      # TableID [string]
                $true,   # RightAlignPageNumbers [bool]
                $true,   # IncludePageNumbers [bool]
                "",      # AddedStyles [string]
                $true,   # UseHyperlinks [bool]
                $true,   # HidePageNumbersInWeb [bool]
                $true    # UseOutlineLevels [bool]
            )
            
            Write-Host "目次を追加しました。"
        }
        catch
        {
            $global:Common.HandleError("4011", "目次追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "目次の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # フォントを設定
    [void] SetFont([string]$font_name, [double]$font_size)
    {
        try
        {
            if ([string]::IsNullOrEmpty($font_name))
            {
                throw "フォント名が指定されていません。"
            }

            if ($font_size -eq $null)
            {
                $font_size = 10.5
            }

            if ($font_size -le 0)
            {
                throw "フォントサイズは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $this.word_document.Content.Font.Name = $font_name
            $this.word_document.Content.Font.Size = $font_size
            
            Write-Host "フォントを設定しました: $font_name, サイズ: $font_size"
        }
        catch
        {
            $global:Common.HandleError("4015", "フォント設定エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "フォントの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ページの向きを設定（Portrait または Landscape）
    [void] SetPageOrientation([string]$orientation)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ([string]::IsNullOrWhiteSpace($orientation))
            {
                throw "ページ向きが指定されていません。"
            }

            $normalized = $orientation.Trim().ToLower()
            $wd = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientPortrait
            switch ($normalized)
            {
                'portrait' { $wd = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientPortrait }
                '縦' { $wd = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientPortrait }
                '縦向き' { $wd = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientPortrait }
                'landscape' { $wd = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape }
                '横' { $wd = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape }
                '横向き' { $wd = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape }
                default { throw "サポートされていないページ向きです: $orientation" }
            }

            foreach ($section in $this.word_document.Sections)
            {
                $section.PageSetup.Orientation = $wd
            }

            Write-Host "ページ向きを設定しました: $orientation"
        }
        catch
        {
            $global:Common.HandleError("4017", "ページ向き設定エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "ページ向きの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ファイル操作関連
    # ========================================

    # ドキュメントを保存
    [void] SaveDocument([string]$file_path = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            # 保存先パスの決定と更新
            if ([string]::IsNullOrEmpty($file_path))
            {
                if ([string]::IsNullOrEmpty($this.file_path))
                {
                    throw "保存先のファイルパスが指定されていません。"
                }
            }
            else
            {
                $this.file_path = $file_path
            }

            $this.word_document.SaveAs($this.file_path)
            $this.is_saved = $true
            
            Write-Host "ドキュメントを保存しました: $($this.file_path)"
        }
        catch
        {
            $global:Common.HandleError("4012", "ドキュメント保存エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "ドキュメントの保存に失敗しました: $($_.Exception.Message)"
        }
    }

    # 目次を更新
    [void] UpdateTableOfContents()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ($this.word_document.TablesOfContents.Count -gt 0)
            {
                $this.word_document.TablesOfContents.Item(1).Update()
                Write-Host "目次を更新しました。"
            }
            else
            {
                Write-Host "更新する目次が見つかりません。"
            }
        }
        catch
        {
            $global:Common.HandleError("4013", "目次更新エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "目次の更新に失敗しました: $($_.Exception.Message)"
        }
    }

    # ドキュメントを開く
    [void] OpenDocument([string]$file_path)
    {
        try
        {
            if ([string]::IsNullOrEmpty($file_path))
            {
                throw "ファイルパスが指定されていません。"
            }

            if (-not (Test-Path $file_path))
            {
                throw "指定されたファイルが見つかりません: $file_path"
            }

            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            # 既存のドキュメントを閉じる
            if ($this.word_document)
            {
                $this.word_document.Close($false)
            }

            $this.word_document = $this.word_app.Documents.Open($file_path)
            $this.file_path = $file_path
            
            Write-Host "ドキュメントを開きました: $file_path"
        }
        catch
        {
            $global:Common.HandleError("4014", "ドキュメント開くエラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            throw "ドキュメントを開くのに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # エラーハンドリング・クリーンアップ関連
    # ========================================

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            Write-Host "初期化失敗時のクリーンアップを開始します。" -ForegroundColor Yellow
            
            # ドキュメントを閉じる
            if ($this.word_document)
            {
                try
                {
                    $this.word_document.Close($false)
                    $this.word_document = $null
                }
                catch
                {
                    Write-Host "ドキュメントの閉じる際にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # Wordアプリケーションを終了
            if ($this.word_app)
            {
                try
                {
                    $this.word_app.Quit()
                    $this.word_app = $null
                }
                catch
                {
                    Write-Host "Wordアプリケーションの終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # 一時ディレクトリを削除
            if ($this.temp_directory -and (Test-Path $this.temp_directory))
            {
                try
                {
                    Remove-Item -Path $this.temp_directory -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "一時ディレクトリを削除しました: $($this.temp_directory)" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "一時ディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            Write-Host "クリーンアップが完了しました。" -ForegroundColor Yellow
        }
        catch
        {
            Write-Host "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # ========================================
    # リソース管理関連
    # ========================================

    # リソースを解放
    [void] Dispose()
    {
        try
        {
            Write-Host "WordDriverのリソースを解放します。" -ForegroundColor Cyan
            
            # ドキュメントを保存して閉じる
            if ($this.word_document -and -not $this.is_saved)
            {
                try
                {
                    $this.word_document.Save()
                    Write-Host "ドキュメントを自動保存しました。" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "ドキュメントの自動保存に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # ドキュメントを閉じる
            if ($this.word_document)
            {
                try
                {
                    $this.word_document.Close($true)
                    $this.word_document = $null
                }
                catch
                {
                    Write-Host "ドキュメントの閉じる際にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # Wordアプリケーションを終了
            if ($this.word_app)
            {
                try
                {
                    $this.word_app.Quit()
                    $this.word_app = $null
                }
                catch
                {
                    Write-Host "Wordアプリケーションの終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            $this.is_initialized = $false
            Write-Host "WordDriverのリソース解放が完了しました。" -ForegroundColor Green
        }
        catch
        {
            $global:Common.HandleError("4016", "WordDriver Disposeエラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
            Write-Host "WordDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} 




