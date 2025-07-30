# Wordファイル操作クラス
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Word

# WordDriverエラー管理モジュールをインポート
import-module "$PSScriptRoot\WordDriverErrors.psm1"

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
            
            $this.is_initialized = $true
            Write-Host "WordDriverの初期化が完了しました。"
        }
        catch
        {
            $this.CleanupOnInitializationFailure()
            LogWordDriverError $($WordDriverErrorCodes.INIT_ERROR) "WordDriver初期化エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.TEMP_DIR_ERROR "一時ディレクトリ作成エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.WORD_APP_ERROR "Wordアプリケーション初期化エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.NEW_DOCUMENT_ERROR "新規ドキュメント作成エラー: $($_.Exception.Message)"
            throw "新規ドキュメントの作成に失敗しました: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.ADD_TEXT_ERROR "テキスト追加エラー: $($_.Exception.Message)"
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
            $paragraph.Range.Style = "Heading $level"
            $paragraph.Range.InsertParagraphAfter()
            
            Write-Host "見出しを追加しました: $text (レベル: $level)"
        }
        catch
        {
            LogWordDriverError WordDriverErrorCodes.ADD_HEADING_ERROR "見出し追加エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.ADD_PARAGRAPH_ERROR "段落追加エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.ADD_TABLE_ERROR "テーブル追加エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.ADD_IMAGE_ERROR "画像追加エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.ADD_PAGE_BREAK_ERROR "ページ区切り追加エラー: $($_.Exception.Message)"
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

            # 目次を追加
            $this.word_document.TablesOfContents.Add($range, $true, 1, 3, "", "", "", $true, "", $true, $true, 1)
            
            Write-Host "目次を追加しました。"
        }
        catch
        {
            LogWordDriverError WordDriverErrorCodes.ADD_TOC_ERROR "目次追加エラー: $($_.Exception.Message)"
            throw "目次の追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # フォントを設定
    [void] SetFont([string]$font_name, [int]$font_size)
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
            LogWordDriverError WordDriverErrorCodes.SET_FONT_ERROR "フォント設定エラー: $($_.Exception.Message)"
            throw "フォントの設定に失敗しました: $($_.Exception.Message)"
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

            if ([string]::IsNullOrEmpty($file_path))
            {
                $this.file_path = $this.file_path
            }

            $this.word_document.SaveAs($file_path)
            $this.is_saved = $true
            
            Write-Host "ドキュメントを保存しました: $file_path"
        }
        catch
        {
            LogWordDriverError WordDriverErrorCodes.SAVE_DOCUMENT_ERROR "ドキュメント保存エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.UPDATE_TOC_ERROR "目次更新エラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.OPEN_DOCUMENT_ERROR "ドキュメント開くエラー: $($_.Exception.Message)"
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
            LogWordDriverError WordDriverErrorCodes.DISPOSE_ERROR "WordDriver Disposeエラー: $($_.Exception.Message)"
            Write-Host "WordDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} 