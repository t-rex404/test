# Wordファイル操作クラス
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Word

# WordDriverエラー管理モジュールをインポート
. "$PSScriptRoot\WordDriverErrors.ps1"

class WordDriver
{
    [Microsoft.Office.Interop.Word.Application]$word_app
    [Microsoft.Office.Interop.Word.Document]$word_document
    [string]$file_path
    [bool]$is_initialized
    [bool]$is_saved
    [string]$temp_directory

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
            
            $this.is_initialized = $true
            Write-Host "WordDriverの初期化が完了しました。"
        }
        catch
        {
            $this.CleanupOnInitializationFailure()
                               LogWordDriverError $WordDriverErrorCodes.INIT_ERROR "WordDriver初期化エラー: $($_.Exception.Message)"
            throw "WordDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 一時ディレクトリを作成
    [string] CreateTempDirectory()
    {
        try
        {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $temp_dir = Join-Path $env:TEMP "WordDriver_$timestamp"
            if (-not (Test-Path $temp_dir))
            {
                New-Item -ItemType Directory -Path $temp_dir -Force | Out-Null
            }
            
            Write-Host "一時ディレクトリを作成しました: $temp_dir"
            return $temp_dir
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.TEMP_DIR_ERROR "一時ディレクトリ作成エラー: $($_.Exception.Message)"
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
                           LogWordDriverError $WordDriverErrorCodes.WORD_APP_ERROR "Wordアプリケーション初期化エラー: $($_.Exception.Message)"
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
                           LogWordDriverError $WordDriverErrorCodes.NEW_DOCUMENT_ERROR "新規ドキュメント作成エラー: $($_.Exception.Message)"
            throw "新規ドキュメントの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # テキストを追加
    [void] AddText([string]$text, [string]$style = "Normal")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ([string]::IsNullOrEmpty($text))
            {
                Write-Host "空のテキストは追加されません。"
                return
            }

            $range = $this.word_document.Range($this.word_document.Content.End - 1, $this.word_document.Content.End - 1)
            $range.Text = $text + "`r`n"
            $range.Style = $style
            
            $displayText = if ($text.Length -gt 50) { $text.Substring(0, 50) + "..." } else { $text }
            Write-Host "テキストを追加しました: $displayText"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.ADD_TEXT_ERROR "テキスト追加エラー: $($_.Exception.Message)"
            throw "テキストの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 見出しを追加
    [void] AddHeading([string]$text, [int]$level = 1)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ([string]::IsNullOrEmpty($text))
            {
                Write-Host "空の見出しは追加されません。"
                return
            }

            if ($level -lt 1 -or $level -gt 9)
            {
                $level = 1
            }

            $style_name = "Heading $level"
            $this.AddText($text, $style_name)
            
            Write-Host "見出しレベル$levelを追加しました: $text"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.ADD_HEADING_ERROR "見出し追加エラー: $($_.Exception.Message)"
            throw "見出しの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 段落を追加
    [void] AddParagraph([string]$text, [string]$alignment = "Left")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ([string]::IsNullOrEmpty($text))
            {
                Write-Host "空の段落は追加されません。"
                return
            }

            $paragraph = $this.word_document.Paragraphs.Add()
            $paragraph.Range.Text = $text
            
            # 配置を設定
            switch ($alignment.ToLower())
            {
                "left" { $paragraph.Alignment = 0 }    # wdAlignParagraphLeft
                "center" { $paragraph.Alignment = 1 }  # wdAlignParagraphCenter
                "right" { $paragraph.Alignment = 2 }   # wdAlignParagraphRight
                "justify" { $paragraph.Alignment = 3 } # wdAlignParagraphJustify
                default { $paragraph.Alignment = 0 }
            }
            
            $displayText = if ($text.Length -gt 50) { $text.Substring(0, 50) + "..." } else { $text }
            Write-Host "段落を追加しました: $displayText"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.ADD_PARAGRAPH_ERROR "段落追加エラー: $($_.Exception.Message)"
            throw "段落の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 表を追加
    [void] AddTable([object[]]$data, [int]$rows, [int]$columns, [string]$title = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ($rows -le 0 -or $columns -le 0)
            {
                throw "行数と列数は1以上である必要があります。"
            }

            # 表のタイトルを追加
            if (-not [string]::IsNullOrEmpty($title))
            {
                $this.AddParagraph($title, "Center")
            }

            # 表を作成
            $table = $this.word_document.Tables.Add($this.word_document.Range($this.word_document.Content.End - 1), $rows, $columns)
            
            # データを設定
            if ($data -and $data.Length -gt 0)
            {
                $data_index = 0
                for ($i = 1; $i -le $rows; $i++)
                {
                    for ($j = 1; $j -le $columns; $j++)
                    {
                        if ($data_index -lt $data.Length)
                        {
                            $table.Cell($i, $j).Range.Text = $data[$data_index].ToString()
                            $data_index++
                        }
                    }
                }
            }

            # 表の後に改行を追加
            $this.AddText("")
            
            Write-Host "表を追加しました: ${rows}行×${columns}列"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.ADD_TABLE_ERROR "表追加エラー: $($_.Exception.Message)"
            throw "表の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 画像を追加
    [void] AddImage([string]$image_path, [int]$width = 400, [int]$height = 300)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if (-not (Test-Path $image_path))
            {
                throw "指定された画像ファイルが見つかりません: $image_path"
            }

            $range = $this.word_document.Range($this.word_document.Content.End - 1, $this.word_document.Content.End - 1)
            $shape = $this.word_document.InlineShapes.AddPicture($image_path, $false, $true, $range)
            
            # サイズを設定
            $shape.Width = $width
            $shape.Height = $height
            
            # 画像の後に改行を追加
            $this.AddText("")
            
            Write-Host "画像を追加しました: $image_path"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.ADD_IMAGE_ERROR "画像追加エラー: $($_.Exception.Message)"
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

            $range = $this.word_document.Range($this.word_document.Content.End - 1, $this.word_document.Content.End - 1)
            $range.InsertBreak(7) # wdPageBreak
            
            Write-Host "ページ区切りを追加しました。"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.ADD_PAGE_BREAK_ERROR "ページ区切り追加エラー: $($_.Exception.Message)"
            throw "ページ区切りの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 目次を追加
    [void] AddTableOfContents([int]$levels = 3)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ($levels -lt 1 -or $levels -gt 9)
            {
                $levels = 3
            }

            # 目次の前にページ区切りを追加
            $this.AddPageBreak()
            
            # 目次タイトルを追加
            $this.AddHeading("目次", 1)
            
            # 目次を挿入
            $range = $this.word_document.Range($this.word_document.Content.End - 1, $this.word_document.Content.End - 1)
            $this.word_document.TablesOfContents.Add($range, $true, 1, $levels, 1, "", "", "", $true)
            
            # 目次の後にページ区切りを追加
            $this.AddPageBreak()
            
            Write-Host "目次を追加しました（レベル: $levels）"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.ADD_TOC_ERROR "目次追加エラー: $($_.Exception.Message)"
            throw "目次の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # ドキュメントを保存
    [void] Save([string]$file_path = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            # ファイルパスが指定されていない場合はデフォルトパスを使用
            if ([string]::IsNullOrEmpty($file_path))
            {
                $file_path = $this.file_path
            }

            # 目次を更新
            $this.UpdateTableOfContents()
            
            # ドキュメントを保存
            $this.word_document.SaveAs($file_path)
            $this.file_path = $file_path
            $this.is_saved = $true
            
            Write-Host "ドキュメントを保存しました: $file_path"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.SAVE_ERROR "ドキュメント保存エラー: $($_.Exception.Message)"
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

            # 目次が存在する場合は更新
            if ($this.word_document.TablesOfContents.Count -gt 0)
            {
                foreach ($toc in $this.word_document.TablesOfContents)
                {
                    $toc.Update()
                }
                Write-Host "目次を更新しました。"
            }
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.UPDATE_TOC_ERROR "目次更新エラー: $($_.Exception.Message)"
            Write-Host "目次の更新に失敗しましたが、処理を続行します: $($_.Exception.Message)"
        }
    }

    # ドキュメントを開く
    [void] OpenDocument([string]$file_path)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if (-not (Test-Path $file_path))
            {
                throw "指定されたファイルが見つかりません: $file_path"
            }

            # 既存のドキュメントを閉じる
            if ($this.word_document)
            {
                $this.word_document.Close($false)
                $this.word_document = $null
            }

            # 新しいドキュメントを開く
            $this.word_document = $this.word_app.Documents.Open($file_path)
            $this.file_path = $file_path
            $this.is_saved = $true
            
            Write-Host "ドキュメントを開きました: $file_path"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.OPEN_DOCUMENT_ERROR "ドキュメントオープンエラー: $($_.Exception.Message)"
            throw "ドキュメントのオープンに失敗しました: $($_.Exception.Message)"
        }
    }

    # フォントを設定
    [void] SetFont([string]$font_name = "MS Gothic", [int]$font_size = 12, [bool]$is_bold = $false, [bool]$is_italic = $false)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $this.word_document.Content.Font.Name = $font_name
            $this.word_document.Content.Font.Size = $font_size
            $this.word_document.Content.Font.Bold = $is_bold
            $this.word_document.Content.Font.Italic = $is_italic
            
            Write-Host "フォントを設定しました: $font_name, $font_size pt"
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.SET_FONT_ERROR "フォント設定エラー: $($_.Exception.Message)"
            throw "フォントの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            if ($this.word_document)
            {
                $this.word_document.Close($false)
                $this.word_document = $null
            }
            
            if ($this.word_app)
            {
                $this.word_app.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.word_app) | Out-Null
                $this.word_app = $null
            }
            
            $this.is_initialized = $false
            Write-Host "初期化失敗時のクリーンアップが完了しました。"
        }
        catch
        {
            Write-Host "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)"
        }
    }

    # リソースを解放
    [void] Dispose()
    {
        try
        {
            if ($this.is_initialized)
            {
                if ($this.word_document)
                {
                    $this.word_document.Close($false)
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.word_document) | Out-Null
                    $this.word_document = $null
                }
                
                if ($this.word_app)
                {
                    $this.word_app.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.word_app) | Out-Null
                    $this.word_app = $null
                }
                
                $this.is_initialized = $false
                Write-Host "WordDriverのリソースを正常に解放しました。"
            }
        }
        catch
        {
                           LogWordDriverError $WordDriverErrorCodes.DISPOSE_ERROR "WordDriver Disposeエラー: $($_.Exception.Message)"
            Write-Host "WordDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)"
        }
    }
} 