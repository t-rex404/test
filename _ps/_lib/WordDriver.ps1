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
            Write-Host "WordDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()
            
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0001", "WordDriver初期化エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WordDriver初期化エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "WordDriverの初期化に失敗しました: $($_.Exception.Message)"
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
            $temp_dir = Join-Path $base_dir "WordDriver"
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
                    $global:Common.HandleError("WordError_0002", "一時ディレクトリ作成エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "一時ディレクトリ作成エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "一時ディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # Wordアプリケーションを初期化
    [void] InitializeWordApplication()
    {
        try
        {
            Write-Host "Wordアプリケーションの初期化を開始します..." -ForegroundColor Cyan
            
            $this.word_app = New-Object -ComObject Word.Application
            Write-Host "Wordアプリケーションオブジェクトを作成しました" -ForegroundColor Green
            
            $this.word_app.Visible = $false
            $this.word_app.DisplayAlerts = $false
            $this.word_app.ScreenUpdating = $false
            Write-Host "基本設定を適用しました" -ForegroundColor Green
            
            # 応答性を向上させる設定（安全に実行）
            try
            {
                Write-Host "応答性向上設定を適用中..." -ForegroundColor Cyan
                
                # 各オプションを個別に設定（エラーが発生しても続行）
                if ($this.word_app.Options -ne $null)
                {
                    # 基本的な設定（ほとんどのバージョンで利用可能）
                    try { $this.word_app.Options.CheckGrammarAsYouType = $false } catch { Write-Host "文法チェック設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                    try { $this.word_app.Options.CheckSpellingAsYouType = $false } catch { Write-Host "スペルチェック設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                    
                    # 自動フォーマット関連の設定（バージョンによって利用できない場合がある）
                    try { 
                        if ($this.word_app.Options.PSObject.Properties.Name -contains 'AutoFormatAsYouTypeApplyFormatting') {
                            $this.word_app.Options.AutoFormatAsYouTypeApplyFormatting = $false
                            Write-Host "自動フォーマット設定を無効化しました" -ForegroundColor Green
                        } else {
                            Write-Host "自動フォーマット設定はこのバージョンでは利用できません" -ForegroundColor Yellow
                        }
                    } catch { Write-Host "自動フォーマット設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                    
                    try { 
                        if ($this.word_app.Options.PSObject.Properties.Name -contains 'AutoFormatAsYouTypeApplyHeadings') {
                            $this.word_app.Options.AutoFormatAsYouTypeApplyHeadings = $false
                            Write-Host "自動見出し設定を無効化しました" -ForegroundColor Green
                        } else {
                            Write-Host "自動見出し設定はこのバージョンでは利用できません" -ForegroundColor Yellow
                        }
                    } catch { Write-Host "自動見出し設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                    
                    try { 
                        if ($this.word_app.Options.PSObject.Properties.Name -contains 'AutoFormatAsYouTypeApplyBullets') {
                            $this.word_app.Options.AutoFormatAsYouTypeApplyBullets = $false
                            Write-Host "自動箇条書き設定を無効化しました" -ForegroundColor Green
                        } else {
                            Write-Host "自動箇条書き設定はこのバージョンでは利用できません" -ForegroundColor Yellow
                        }
                    } catch { Write-Host "自動箇条書き設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                    
                    try { 
                        if ($this.word_app.Options.PSObject.Properties.Name -contains 'AutoFormatAsYouTypeApplyNumbering') {
                            $this.word_app.Options.AutoFormatAsYouTypeApplyNumbering = $false
                            Write-Host "自動番号設定を無効化しました" -ForegroundColor Green
                        } else {
                            Write-Host "自動番号設定はこのバージョンでは利用できません" -ForegroundColor Yellow
                        }
                    } catch { Write-Host "自動番号設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                    
                    # その他の応答性向上設定
                    try { 
                        if ($this.word_app.Options.PSObject.Properties.Name -contains 'ConfirmConversions') {
                            $this.word_app.Options.ConfirmConversions = $false
                            Write-Host "変換確認設定を無効化しました" -ForegroundColor Green
                        }
                    } catch { Write-Host "変換確認設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                    
                    try { 
                        if ($this.word_app.Options.PSObject.Properties.Name -contains 'UpdateLinksAtOpen') {
                            $this.word_app.Options.UpdateLinksAtOpen = $false
                            Write-Host "リンク更新設定を無効化しました" -ForegroundColor Green
                        }
                    } catch { Write-Host "リンク更新設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow }
                }
                else
                {
                    Write-Host "Wordアプリケーションのオプションが利用できません。基本設定のみで続行します。" -ForegroundColor Yellow
                }
                
                Write-Host "応答性向上設定が完了しました" -ForegroundColor Green
            }
            catch
            {
                Write-Host "応答性向上設定でエラーが発生しましたが、基本機能は継続します: $($_.Exception.Message)" -ForegroundColor Yellow
            }
            
            # 初期化完了を確認
            Write-Host "初期化完了を確認中..." -ForegroundColor Cyan
            Start-Sleep -Milliseconds 500
            
            # アプリケーションの状態を確認
            if ($this.word_app -eq $null)
            {
                throw "Wordアプリケーションオブジェクトが作成されていません"
            }
            
            Write-Host "Wordアプリケーションの初期化が完了しました。" -ForegroundColor Green
        }
        catch
        {
            Write-Host "Wordアプリケーション初期化で致命的なエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try {
                    $global:Common.HandleError("WordError_0003", "Wordアプリケーション初期化エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                } catch {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Wordアプリケーション初期化エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            #$this.word_document.SaveAs([ref]$this.file_path)
            Write-Host "新規ドキュメントを作成しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0004", "新規ドキュメント作成エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "新規ドキュメント作成エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "新規ドキュメントの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # フッターに総ページと現在のページを設定
    [void] SetupFooterWithPageNumbers()
    {
        try
        {
            Write-Host "フッターのページ番号設定を開始します..." -ForegroundColor Cyan
            
            # ドキュメントの状態を確認
            if ($this.word_document -eq $null)
            {
                throw "Wordドキュメントが存在しません。"
            }
            
            # セクション数を確認
            $sectionCount = $this.word_document.Sections.Count
            Write-Host "セクション数: $sectionCount" -ForegroundColor Yellow
            
            # 各セクションのフッターを設定
            for ($i = 1; $i -le $sectionCount; $i++)
            {
                try
                {
                    Write-Host "セクション $i のフッターを設定中..." -ForegroundColor Yellow
                    
                    # セクションを安全に取得
                    $section = $this.word_document.Sections.Item($i)
                    if ($section -eq $null)
                    {
                        Write-Host "セクション $i の取得に失敗しました。スキップします。" -ForegroundColor Yellow
                        continue
                    }
                    
                    # フッターを安全に取得
                    $footer = $section.Footers.Item([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterPrimary)
                    if ($footer -eq $null)
                    {
                        Write-Host "セクション $i のフッター取得に失敗しました。スキップします。" -ForegroundColor Yellow
                        continue
                    }

                    # フッターの内容をクリア
                    $footer.Range.Text = ""

                    # フッターの最後に移動
                    $footer.Range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd) | Out-Null
                    
                    # Wordのクイックパーツ機能を使用してページ番号を挿入
                    try
                    {
                        # ページ番号フィールドを挿入
                        $pageField = $footer.Range.Fields.Add($footer.Range.Duplicate, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldPage, $null, $true)
                        $pageField.Update()
                        $footer.Range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd) | Out-Null
                        
                        # " / " を挿入
                        $footer.Range.InsertAfter(" / ")
                        $footer.Range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd) | Out-Null
                        
                        # 総ページ数フィールドを挿入
                        $numPagesField = $footer.Range.Fields.Add($footer.Range.Duplicate, [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldNumPages, $null, $true)
                        $numPagesField.Update()
                        $footer.Range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd) | Out-Null

                        # フッター内のフィールドを更新
                        $footer.Range.Fields.Update()
                        $footer.Range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd) | Out-Null    

                        ## フィールドを更新
                        #$pageField.Update()
                        #$numPagesField.Update()
                        #$footer.Range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd) | Out-Null

                    }
                    catch
                    {
                        Write-Host "クイックパーツ機能でのページ番号挿入でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                        
                        # フォールバック: フィールドコードを直接挿入
                        Write-Host "フォールバック方法でページ番号を挿入します..." -ForegroundColor Yellow
                        $footer.Range.InsertAfter("PAGE / NUMPAGES")
                        $footer.Range.Fields.Update()

                    }

                    # 中央揃えでページ番号を挿入
                    $footer.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
                    
                    # フィールドの表示を確実にするための処理
                    try
                    {
                        if ($footer.Range.Fields.Count -gt 0)
                        {
                            Write-Host "フィールド数: $($footer.Range.Fields.Count)" -ForegroundColor Cyan
                            # 各フィールドを個別に更新
                            foreach ($field in $footer.Range.Fields)
                            {
                                if ($field -ne $null)
                                {
                                    $field.Update()
                                }
                            }
                        }
                        else
                        {
                            Write-Host "フィールドが挿入されていません。" -ForegroundColor Yellow
                        }
                    }
                    catch
                    {
                        Write-Host "フィールドの更新でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                    
                    Write-Host "セクション $i のフッター設定完了" -ForegroundColor Green
                }
                catch
                {
                    Write-Host "セクション $i のフッター設定でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                    continue
                }
            }
            
            # ドキュメント全体のフィールドを更新
            Write-Host "ドキュメント全体のフィールドを更新中..." -ForegroundColor Cyan
            $this.word_document.Fields.Update()
            
            # ドキュメントを再計算
            Write-Host "ドキュメントを再計算中..." -ForegroundColor Cyan
            $this.word_document.Repaginate()
            
            # フィールドの表示を確実にするために少し待機
            Start-Sleep -Milliseconds 500
            
            # 再度フィールドを更新（確実性のため）
            $this.word_document.Fields.Update()
            
            # フィールドの表示を確実にするための追加設定
            try
            {
                # 各セクションのフッター内のフィールドを個別に更新
                for ($i = 1; $i -le $sectionCount; $i++)
                {
                    try
                    {
                        $section = $this.word_document.Sections.Item($i)
                        if ($section -ne $null)
                        {
                            $footer = $section.Footers.Item([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterPrimary)
                            if ($footer -ne $null -and $footer.Range.Fields.Count -gt 0)
                            {
                                foreach ($field in $footer.Range.Fields)
                                {
                                    if ($field -ne $null)
                                    {
                                        $field.Update()
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Host "セクション $i のフィールド更新でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
                
                Write-Host "フィールドの個別更新が完了しました。" -ForegroundColor Green
            }
            catch
            {
                Write-Host "フィールドの個別更新でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
            
            Write-Host "フィールドの更新が完了しました。" -ForegroundColor Green
            
            Write-Host "フッターにページ番号を設定しました（中央揃え）。" -ForegroundColor Green
        }
        catch
        {
            Write-Host "フッターのページ番号設定でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0018", "フッターページ番号設定エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "フッターのページ番号設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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

            # より安全で確実な方法でテキストを追加
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertAfter($text)
            $range.InsertParagraphAfter()
            
            # カーソル位置を最後に移動
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

            Write-Host "テキストを追加しました: $text"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0005", "テキスト追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
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
            Write-Host "見出しの追加でエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0006", "見出し追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "見出しの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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

            # より安全で確実な方法で段落を追加
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertAfter($text)
            $range.InsertParagraphAfter()
            
            # カーソル位置を最後に移動
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            
            Write-Host "段落を追加しました: $text"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0007", "段落追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "段落の追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

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

            # ドキュメントの最後にテーブルを挿入するためのRangeを作成
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            
            # テーブルを作成（正しい位置に挿入）
            $table = $this.word_document.Tables.Add($range, $rows, $cols)

            # データを設定
            for ($i = 0; $i -lt $rows; $i++)
            {
                for ($j = 0; $j -lt $cols; $j++)
                {
                    $table.Cell($i + 1, $j + 1).Range.Text = $data[$i, $j].ToString()
                }
            }

            # テーブルの後に段落を追加
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertParagraphAfter()

            Write-Host "テーブルを追加しました: $rows 行 x $cols 列"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0008", "テーブル追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テーブルの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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

            # ドキュメントの最後に画像を挿入するためのRangeを作成
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            
            # 画像を挿入（正しい位置に挿入）
            $this.word_document.InlineShapes.AddPicture($image_path, $false, $true, $range)
            
            # 画像の後に段落を追加
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertParagraphAfter()
            
            Write-Host "画像を追加しました: $image_path"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0009", "画像追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
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

    # ページ区切りを追加
    [void] AddPageBreak()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
            
            Write-Host "ページ区切りを追加しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0010", "ページ区切り追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ページ区切りの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ページ区切りの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # セクション管理関連
    # ========================================

    # 新しいセクションを追加
    [void] AddSection()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            # ドキュメントの最後にセクション区切りを挿入
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage)
            
            Write-Host "新しいセクションを追加しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0020", "セクション追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "セクションの追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "セクションの追加に失敗しました: $($_.Exception.Message)"
        }
    }

    # 指定した位置にセクション区切りを挿入
    [void] InsertSectionBreak([int]$position, [string]$break_type = "NextPage")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ($position -lt 0)
            {
                throw "位置は0以上の値である必要があります。"
            }

            # ブレークタイプを正規化
            $normalized_break_type = $break_type.Trim().ToLower()
            $wd_break_type = [Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage
            
            switch ($normalized_break_type)
            {
                'nextpage' { $wd_break_type = [Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage }
                'continuous' { $wd_break_type = [Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakContinuous }
                'evenpage' { $wd_break_type = [Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakEvenPage }
                'oddpage' { $wd_break_type = [Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakOddPage }
                default { throw "サポートされていないセクション区切りタイプです: $break_type" }
            }

            # 指定した位置にセクション区切りを挿入
            $range = $this.word_document.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseStart)
            $range.Move([Microsoft.Office.Interop.Word.WdUnits]::wdCharacter, $position)
            $range.InsertBreak($wd_break_type)
            
            Write-Host "指定した位置にセクション区切りを挿入しました: 位置 $position, タイプ $break_type"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0021", "セクション区切り挿入エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "セクション区切りの挿入に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "セクション区切りの挿入に失敗しました: $($_.Exception.Message)"
        }
    }

    # セクション数を取得
    [int] GetSectionCount()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            $section_count = $this.word_document.Sections.Count
            Write-Host "セクション数: $section_count"
            return $section_count
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0022", "セクション数取得エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "セクション数の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "セクション数の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 現在のセクションインデックスを取得
    [int] GetCurrentSectionIndex()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            # 現在のカーソル位置のセクションを取得
            $current_range = $this.word_document.Application.Selection.Range
            $current_section = $current_range.Sections.Item(1)
            
            # セクションインデックスを取得
            for ($i = 1; $i -le $this.word_document.Sections.Count; $i++)
            {
                if ($this.word_document.Sections.Item($i) -eq $current_section)
                {
                    Write-Host "現在のセクションインデックス: $i"
                    return $i
                }
            }
            
            # 見つからない場合は1を返す
            Write-Host "現在のセクションインデックス: 1 (デフォルト)"
            return 1
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0023", "現在のセクションインデックス取得エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "現在のセクションインデックスの取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "現在のセクションインデックスの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 指定したセクションのページ設定を変更
    [void] SetSectionPageSetup([int]$section_index, [hashtable]$page_settings)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ($section_index -lt 1 -or $section_index -gt $this.word_document.Sections.Count)
            {
                throw "無効なセクションインデックスです: $section_index"
            }

            if (-not $page_settings)
            {
                throw "ページ設定が指定されていません。"
            }

            # 指定したセクションを取得
            $section = $this.word_document.Sections.Item($section_index)
            if ($section -eq $null)
            {
                throw "セクション $section_index の取得に失敗しました。"
            }

            # ページ設定を適用
            if ($page_settings.ContainsKey('Orientation'))
            {
                $orientation = $page_settings['Orientation'].ToString().ToLower()
                $wd_orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientPortrait
                
                switch ($orientation)
                {
                    'portrait' { $wd_orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientPortrait }
                    'landscape' { $wd_orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape }
                    default { throw "サポートされていないページ向きです: $orientation" }
                }
                
                $section.PageSetup.Orientation = $wd_orientation
                Write-Host "セクション $section_index のページ向きを設定しました: $orientation"
            }

            if ($page_settings.ContainsKey('TopMargin'))
            {
                $section.PageSetup.TopMargin = $page_settings['TopMargin']
                Write-Host "セクション $section_index の上マージンを設定しました: $($page_settings['TopMargin'])"
            }

            if ($page_settings.ContainsKey('BottomMargin'))
            {
                $section.PageSetup.BottomMargin = $page_settings['BottomMargin']
                Write-Host "セクション $section_index の下マージンを設定しました: $($page_settings['BottomMargin'])"
            }

            if ($page_settings.ContainsKey('LeftMargin'))
            {
                $section.PageSetup.LeftMargin = $page_settings['LeftMargin']
                Write-Host "セクション $section_index の左マージンを設定しました: $($page_settings['LeftMargin'])"
            }

            if ($page_settings.ContainsKey('RightMargin'))
            {
                $section.PageSetup.RightMargin = $page_settings['RightMargin']
                Write-Host "セクション $section_index の右マージンを設定しました: $($page_settings['RightMargin'])"
            }

            if ($page_settings.ContainsKey('PageWidth'))
            {
                $section.PageSetup.PageWidth = $page_settings['PageWidth']
                Write-Host "セクション $section_index のページ幅を設定しました: $($page_settings['PageWidth'])"
            }

            if ($page_settings.ContainsKey('PageHeight'))
            {
                $section.PageSetup.PageHeight = $page_settings['PageHeight']
                Write-Host "セクション $section_index のページ高さを設定しました: $($page_settings['PageHeight'])"
            }

            Write-Host "セクション $section_index のページ設定を完了しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0024", "セクションページ設定エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "セクションのページ設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "セクションのページ設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 指定したセクションに移動
    [void] GoToSection([int]$section_index)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WordDriverが初期化されていません。"
            }

            if ($section_index -lt 1 -or $section_index -gt $this.word_document.Sections.Count)
            {
                throw "無効なセクションインデックスです: $section_index"
            }

            # 指定したセクションの開始位置に移動
            $section = $this.word_document.Sections.Item($section_index)
            if ($section -eq $null)
            {
                throw "セクション $section_index の取得に失敗しました。"
            }

            $section.Range.Select()
            Write-Host "セクション $section_index に移動しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0025", "セクション移動エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "セクションへの移動に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "セクションへの移動に失敗しました: $($_.Exception.Message)"
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0011", "目次追加エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "目次の追加に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0015", "フォント設定エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0017", "ページ向き設定エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ページ向きの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0012", "ドキュメント保存エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ドキュメントの保存に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0013", "目次更新エラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "目次の更新に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0014", "ドキュメント開くエラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ドキュメントを開くのに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0019", "初期化失敗時のクリーンアップエラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
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
                    #$this.word_document.Save()
                    $this.SaveDocument($this.temp_directory + "\WordDriver_AutoSave.docx")
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WordError_0016", "WordDriver Disposeエラー: $($_.Exception.Message)", "WordDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WordDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            Write-Host "WordDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} 




