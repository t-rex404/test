# Accessファイル操作クラス
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Access

class AccessDriver
{
    [object]$access_app
    [object]$access_database
    [string]$file_path
    [bool]$is_initialized
    [bool]$is_saved
    [string]$temp_directory
    [object]$current_recordset

    # ログファイルパス（共有可能）
    static [string]$NormalLogFile = ".\AccessDriver_$($env:USERNAME)_Normal.log"
    static [string]$ErrorLogFile = ".\AccessDriver_$($env:USERNAME)_Error.log"

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
                $global:Common.WriteLog($message, "INFO", "AccessDriver")
            }
            catch
            {
                Write-Host "正常ログ出力に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([AccessDriver]::NormalLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Green
    }

    # エラーログを出力（共通のエラー処理に委譲）
    [void] LogError([string]$errorCode, [string]$message)
    {
        if ($global:Common)
        {
            try
            {
                $global:Common.HandleError($errorCode, $message, "AccessDriver")
            }
            catch
            {
                Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([AccessDriver]::ErrorLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Red
    }

    # ========================================
    # 初期化・接続関連
    # ========================================

    AccessDriver()
    {
        try
        {
            $this.is_initialized = $false
            $this.is_saved = $false

            # 一時ディレクトリの作成
            $this.temp_directory = $this.CreateTempDirectory()

            # Accessアプリケーションの初期化
            $this.InitializeAccessApplication()

            # 新規データベースの作成
            $this.CreateNewDatabase()

            $this.is_initialized = $true
            Write-Host "AccessDriverの初期化が完了しました。"

            # 正常ログ出力
            $this.LogInfo("AccessDriverの初期化が完了しました")
        }
        catch
        {
            Write-Host "AccessDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()

            $this.LogError("AccessDriverError_0001", "AccessDriver初期化エラー: $($_.Exception.Message)")

            throw "AccessDriverの初期化に失敗しました: $($_.Exception.Message)"
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
            $temp_dir = Join-Path $base_dir "AccessDriver"
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
            $this.LogError("AccessDriverError_0002", "一時ディレクトリ作成エラー: $($_.Exception.Message)")

            throw "一時ディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # Accessアプリケーションを初期化
    [void] InitializeAccessApplication()
    {
        try
        {
            Write-Host "Accessアプリケーションの初期化を開始します..." -ForegroundColor Cyan

            $this.access_app = New-Object -ComObject Access.Application
            Write-Host "Accessアプリケーションオブジェクトを作成しました" -ForegroundColor Green

            $this.access_app.Visible = $false
            Write-Host "基本設定を適用しました" -ForegroundColor Green

            # 初期化完了を確認
            Write-Host "初期化完了を確認中..." -ForegroundColor Cyan
            Start-Sleep -Milliseconds 500

            # アプリケーションの状態を確認
            if ($null -eq $this.access_app)
            {
                throw "Accessアプリケーションオブジェクトが作成されていません"
            }

            Write-Host "Accessアプリケーションの初期化が完了しました。" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("Accessアプリケーションの初期化が完了しました")
        }
        catch
        {
            Write-Host "Accessアプリケーション初期化で致命的なエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red

            $this.LogError("AccessDriverError_0010", "Accessアプリケーション初期化エラー: $($_.Exception.Message)")

            throw "Accessアプリケーションの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 新規データベースを作成
    [void] CreateNewDatabase()
    {
        try
        {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $this.file_path = Join-Path $this.temp_directory "Database_$timestamp.accdb"

            # 新規データベースを作成
            $this.access_database = $this.access_app.NewCurrentDatabase($this.file_path)

            Write-Host "新規データベースを作成しました: $($this.file_path)"

            # 正常ログ出力
            $this.LogInfo("新規データベースを作成しました: $($this.file_path)")
        }
        catch
        {
            $this.LogError("AccessDriverError_0011", "新規データベース作成エラー: $($_.Exception.Message)")

            throw "新規データベースの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # テーブル操作関連
    # ========================================

    # テーブルを作成
    [void] CreateTable([string]$tableName, [hashtable]$columns)
    {
        try
        {
            if ([string]::IsNullOrEmpty($tableName))
            {
                throw "テーブル名が指定されていません。"
            }

            if ($null -eq $columns -or $columns.Count -eq 0)
            {
                throw "カラム定義が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            # SQL文を構築
            $columnDefinitions = @()
            foreach ($columnName in $columns.Keys)
            {
                $dataType = $columns[$columnName]
                $columnDefinitions += "$columnName $dataType"
            }
            $columnsString = $columnDefinitions -join ", "

            $sql = "CREATE TABLE $tableName ($columnsString)"

            # SQLを実行
            $this.access_app.CurrentDb().Execute($sql)

            Write-Host "テーブルを作成しました: $tableName"

            # 正常ログ出力
            $this.LogInfo("テーブルを作成しました: $tableName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0020", "テーブル作成エラー: $($_.Exception.Message)")

            throw "テーブルの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # テーブル一覧を取得
    [array] GetTableList()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            $tableList = @()
            $db = $this.access_app.CurrentDb()

            foreach ($table in $db.TableDefs)
            {
                # システムテーブルを除外
                if (-not $table.Name.StartsWith("MSys"))
                {
                    $tableList += $table.Name
                }
            }

            Write-Host "テーブル一覧を取得しました。テーブル数: $($tableList.Count)"

            # 正常ログ出力
            $this.LogInfo("テーブル一覧を取得しました。テーブル数: $($tableList.Count)")

            return $tableList
        }
        catch
        {
            $this.LogError("AccessDriverError_0021", "テーブル一覧取得エラー: $($_.Exception.Message)")

            throw "テーブル一覧の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # テーブルを削除
    [void] DropTable([string]$tableName)
    {
        try
        {
            if ([string]::IsNullOrEmpty($tableName))
            {
                throw "テーブル名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            $sql = "DROP TABLE $tableName"
            $this.access_app.CurrentDb().Execute($sql)

            Write-Host "テーブルを削除しました: $tableName"

            # 正常ログ出力
            $this.LogInfo("テーブルを削除しました: $tableName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0022", "テーブル削除エラー: $($_.Exception.Message)")

            throw "テーブルの削除に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # データ操作関連
    # ========================================

    # データを挿入
    [void] InsertData([string]$tableName, [hashtable]$data)
    {
        try
        {
            if ($null -eq $tableName)
            {
                throw "tableNameパラメータは必須です。"
            }
            if ([string]::IsNullOrEmpty($tableName))
            {
                throw "テーブル名が指定されていません。"
            }
            if ($null -eq $data)
            {
                throw "dataパラメータは必須です。"
            }
            if ($data.Count -eq 0)
            {
                throw "データが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            # カラム名と値を取得
            $columns = $data.Keys -join ", "
            $values = @()
            foreach ($value in $data.Values)
            {
                if ($value -is [string])
                {
                    $values += "'$value'"
                }
                elseif ($null -eq $value)
                {
                    $values += "NULL"
                }
                else
                {
                    $values += $value.ToString()
                }
            }
            $valuesString = $values -join ", "

            $sql = "INSERT INTO $tableName ($columns) VALUES ($valuesString)"
            $this.access_app.CurrentDb().Execute($sql)

            Write-Host "データを挿入しました: $tableName"

            # 正常ログ出力
            $this.LogInfo("データを挿入しました: $tableName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0030", "データ挿入エラー: $($_.Exception.Message)")

            throw "データの挿入に失敗しました: $($_.Exception.Message)"
        }
    }

    # データを更新
    [void] UpdateData([string]$tableName, [hashtable]$data, [string]$whereClause)
    {
        try
        {
            if ($null -eq $tableName)
            {
                throw "tableNameパラメータは必須です。"
            }
            if ([string]::IsNullOrEmpty($tableName))
            {
                throw "テーブル名が指定されていません。"
            }
            if ($null -eq $data)
            {
                throw "dataパラメータは必須です。"
            }
            if ($data.Count -eq 0)
            {
                throw "更新データが指定されていません。"
            }
            if ($null -eq $whereClause)
            {
                throw "whereClauseパラメータは必須です。"
            }
            if ([string]::IsNullOrEmpty($whereClause))
            {
                throw "WHERE句が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            # SET句を構築
            $setClause = @()
            foreach ($column in $data.Keys)
            {
                $value = $data[$column]
                if ($value -is [string])
                {
                    $setClause += "$column = '$value'"
                }
                elseif ($null -eq $value)
                {
                    $setClause += "$column = NULL"
                }
                else
                {
                    $setClause += "$column = $($value.ToString())"
                }
            }
            $setString = $setClause -join ", "

            $sql = "UPDATE $tableName SET $setString WHERE $whereClause"
            $this.access_app.CurrentDb().Execute($sql)

            Write-Host "データを更新しました: $tableName"

            # 正常ログ出力
            $this.LogInfo("データを更新しました: $tableName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0031", "データ更新エラー: $($_.Exception.Message)")

            throw "データの更新に失敗しました: $($_.Exception.Message)"
        }
    }

    # データを削除
    [void] DeleteData([string]$tableName, [string]$whereClause)
    {
        try
        {
            if ([string]::IsNullOrEmpty($tableName))
            {
                throw "テーブル名が指定されていません。"
            }

            if ([string]::IsNullOrEmpty($whereClause))
            {
                throw "WHERE句が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            $sql = "DELETE FROM $tableName WHERE $whereClause"
            $this.access_app.CurrentDb().Execute($sql)

            Write-Host "データを削除しました: $tableName"

            # 正常ログ出力
            $this.LogInfo("データを削除しました: $tableName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0032", "データ削除エラー: $($_.Exception.Message)")

            throw "データの削除に失敗しました: $($_.Exception.Message)"
        }
    }

    # データを検索
    [array] SelectData([string]$sql)
    {
        try
        {
            if ([string]::IsNullOrEmpty($sql))
            {
                throw "SQLが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            $db = $this.access_app.CurrentDb()
            $recordset = $db.OpenRecordset($sql)

            $results = @()

            if (-not $recordset.EOF)
            {
                $recordset.MoveFirst()

                while (-not $recordset.EOF)
                {
                    $row = @{}
                    for ($i = 0; $i -lt $recordset.Fields.Count; $i++)
                    {
                        $field = $recordset.Fields.Item($i)
                        $row[$field.Name] = $field.Value
                    }
                    $results += $row
                    $recordset.MoveNext()
                }
            }

            $recordset.Close()

            Write-Host "データを取得しました。件数: $($results.Count)"

            # 正常ログ出力
            $this.LogInfo("データを取得しました。件数: $($results.Count)")

            return $results
        }
        catch
        {
            $this.LogError("AccessDriverError_0033", "データ検索エラー: $($_.Exception.Message)")

            throw "データの検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # クエリ操作関連
    # ========================================

    # クエリを作成
    [void] CreateQuery([string]$queryName, [string]$sql)
    {
        try
        {
            if ([string]::IsNullOrEmpty($queryName))
            {
                throw "クエリ名が指定されていません。"
            }

            if ([string]::IsNullOrEmpty($sql))
            {
                throw "SQLが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            $db = $this.access_app.CurrentDb()
            $queryDef = $db.CreateQueryDef($queryName, $sql)

            Write-Host "クエリを作成しました: $queryName"

            # 正常ログ出力
            $this.LogInfo("クエリを作成しました: $queryName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0040", "クエリ作成エラー: $($_.Exception.Message)")

            throw "クエリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # クエリを実行
    [array] ExecuteQuery([string]$queryName)
    {
        try
        {
            if ([string]::IsNullOrEmpty($queryName))
            {
                throw "クエリ名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            $db = $this.access_app.CurrentDb()
            $recordset = $db.QueryDefs.Item($queryName).OpenRecordset()

            $results = @()

            if (-not $recordset.EOF)
            {
                $recordset.MoveFirst()

                while (-not $recordset.EOF)
                {
                    $row = @{}
                    for ($i = 0; $i -lt $recordset.Fields.Count; $i++)
                    {
                        $field = $recordset.Fields.Item($i)
                        $row[$field.Name] = $field.Value
                    }
                    $results += $row
                    $recordset.MoveNext()
                }
            }

            $recordset.Close()

            Write-Host "クエリを実行しました: $queryName。件数: $($results.Count)"

            # 正常ログ出力
            $this.LogInfo("クエリを実行しました: $queryName。件数: $($results.Count)")

            return $results
        }
        catch
        {
            $this.LogError("AccessDriverError_0041", "クエリ実行エラー: $($_.Exception.Message)")

            throw "クエリの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # フォーム操作関連
    # ========================================

    # フォームを作成
    [void] CreateForm([string]$formName, [string]$recordSource = "")
    {
        try
        {
            if ([string]::IsNullOrEmpty($formName))
            {
                throw "フォーム名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            # フォームを作成
            $form = $this.access_app.CreateForm()
            $form.Name = $formName

            if (-not [string]::IsNullOrEmpty($recordSource))
            {
                $form.RecordSource = $recordSource
            }

            # フォームを保存
            $this.access_app.DoCmd.Save([Microsoft.Office.Interop.Access.AcObjectType]::acForm, $formName)

            Write-Host "フォームを作成しました: $formName"

            # 正常ログ出力
            $this.LogInfo("フォームを作成しました: $formName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0050", "フォーム作成エラー: $($_.Exception.Message)")

            throw "フォームの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # フォームを開く
    [void] OpenForm([string]$formName, [string]$view = "Normal")
    {
        try
        {
            if ([string]::IsNullOrEmpty($formName))
            {
                throw "フォーム名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            # ビューモードを決定
            $viewMode = switch ($view.ToLower())
            {
                "normal" { [Microsoft.Office.Interop.Access.AcFormView]::acNormal }
                "design" { [Microsoft.Office.Interop.Access.AcFormView]::acDesign }
                "preview" { [Microsoft.Office.Interop.Access.AcFormView]::acPreview }
                "datasheet" { [Microsoft.Office.Interop.Access.AcFormView]::acFormDS }
                default { [Microsoft.Office.Interop.Access.AcFormView]::acNormal }
            }

            $this.access_app.DoCmd.OpenForm($formName, $viewMode)

            Write-Host "フォームを開きました: $formName (ビュー: $view)"

            # 正常ログ出力
            $this.LogInfo("フォームを開きました: $formName (ビュー: $view)")
        }
        catch
        {
            $this.LogError("AccessDriverError_0051", "フォームを開くエラー: $($_.Exception.Message)")

            throw "フォームを開くのに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # レポート操作関連
    # ========================================

    # レポートを作成
    [void] CreateReport([string]$reportName, [string]$recordSource = "")
    {
        try
        {
            if ([string]::IsNullOrEmpty($reportName))
            {
                throw "レポート名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            # レポートを作成
            $report = $this.access_app.CreateReport()
            $report.Name = $reportName

            if (-not [string]::IsNullOrEmpty($recordSource))
            {
                $report.RecordSource = $recordSource
            }

            # レポートを保存
            $this.access_app.DoCmd.Save([Microsoft.Office.Interop.Access.AcObjectType]::acReport, $reportName)

            Write-Host "レポートを作成しました: $reportName"

            # 正常ログ出力
            $this.LogInfo("レポートを作成しました: $reportName")
        }
        catch
        {
            $this.LogError("AccessDriverError_0052", "レポート作成エラー: $($_.Exception.Message)")

            throw "レポートの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # レポートを開く
    [void] OpenReport([string]$reportName, [string]$view = "Preview")
    {
        try
        {
            if ([string]::IsNullOrEmpty($reportName))
            {
                throw "レポート名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            # ビューモードを決定
            $viewMode = switch ($view.ToLower())
            {
                "normal" { [Microsoft.Office.Interop.Access.AcView]::acViewNormal }
                "design" { [Microsoft.Office.Interop.Access.AcView]::acViewDesign }
                "preview" { [Microsoft.Office.Interop.Access.AcView]::acViewPreview }
                "report" { [Microsoft.Office.Interop.Access.AcView]::acViewReport }
                default { [Microsoft.Office.Interop.Access.AcView]::acViewPreview }
            }

            $this.access_app.DoCmd.OpenReport($reportName, $viewMode)

            Write-Host "レポートを開きました: $reportName (ビュー: $view)"

            # 正常ログ出力
            $this.LogInfo("レポートを開きました: $reportName (ビュー: $view)")
        }
        catch
        {
            $this.LogError("AccessDriverError_0053", "レポートを開くエラー: $($_.Exception.Message)")

            throw "レポートを開くのに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ファイル操作関連
    # ========================================

    # データベースを保存
    [void] SaveDatabase([string]$file_path = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
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
                # 新しいパスに保存する場合
                $this.access_app.SaveAsText(6, "", $file_path)
                $this.file_path = $file_path
            }

            $this.is_saved = $true

            Write-Host "データベースを保存しました: $($this.file_path)"

            # 正常ログ出力
            $this.LogInfo("データベースを保存しました: $($this.file_path)")
        }
        catch
        {
            $this.LogError("AccessDriverError_0060", "データベース保存エラー: $($_.Exception.Message)")

            throw "データベースの保存に失敗しました: $($_.Exception.Message)"
        }
    }

    # 既存のデータベースを開く
    [void] OpenDatabase([string]$file_path)
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
                throw "AccessDriverが初期化されていません。"
            }

            # 現在のデータベースを閉じる
            if ($this.access_database)
            {
                $this.access_app.CloseCurrentDatabase()
            }

            $this.access_app.OpenCurrentDatabase($file_path)
            $this.file_path = $file_path

            Write-Host "データベースを開きました: $file_path"

            # 正常ログ出力
            $this.LogInfo("データベースを開きました: $file_path")
        }
        catch
        {
            $this.LogError("AccessDriverError_0061", "データベースを開くエラー: $($_.Exception.Message)")

            throw "データベースを開くのに失敗しました: $($_.Exception.Message)"
        }
    }

    # データベースをコンパクト化
    [void] CompactDatabase([string]$source_path = "", [string]$dest_path = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "AccessDriverが初期化されていません。"
            }

            if ([string]::IsNullOrEmpty($source_path))
            {
                $source_path = $this.file_path
            }

            if ([string]::IsNullOrEmpty($dest_path))
            {
                $dest_path = $source_path + ".tmp"
            }

            # データベースを閉じる
            $this.access_app.CloseCurrentDatabase()

            # コンパクト化
            $this.access_app.DBEngine.CompactDatabase($source_path, $dest_path)

            # ファイルを置き換える
            if ($source_path -eq $this.file_path)
            {
                Remove-Item $source_path -Force
                Move-Item $dest_path $source_path -Force
                $this.access_app.OpenCurrentDatabase($source_path)
            }

            Write-Host "データベースをコンパクト化しました: $source_path"

            # 正常ログ出力
            $this.LogInfo("データベースをコンパクト化しました: $source_path")
        }
        catch
        {
            $this.LogError("AccessDriverError_0080", "データベースコンパクト化エラー: $($_.Exception.Message)")

            throw "データベースのコンパクト化に失敗しました: $($_.Exception.Message)"
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

            # データベースを閉じる
            if ($this.access_app)
            {
                try
                {
                    $this.access_app.CloseCurrentDatabase()
                }
                catch
                {
                    Write-Host "データベースを閉じる際にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # Accessアプリケーションを終了
            if ($this.access_app)
            {
                try
                {
                    $this.access_app.Quit()
                    $this.access_app = $null
                }
                catch
                {
                    Write-Host "Accessアプリケーションの終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
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
            $this.LogError("AccessDriverError_0090", "初期化失敗時のクリーンアップエラー: $($_.Exception.Message)")

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
            Write-Host "AccessDriverのリソースを解放します。" -ForegroundColor Cyan

            # データベースを保存して閉じる
            if ($this.access_app -and -not $this.is_saved)
            {
                try
                {
                    $this.SaveDatabase($this.temp_directory + "\AccessDriver_AutoSave.accdb")
                    Write-Host "データベースを自動保存しました。" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "データベースの自動保存に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # データベースを閉じる
            if ($this.access_app)
            {
                try
                {
                    $this.access_app.CloseCurrentDatabase()
                }
                catch
                {
                    Write-Host "データベースを閉じる際にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # Accessアプリケーションを終了
            if ($this.access_app)
            {
                try
                {
                    $this.access_app.Quit()
                    $this.access_app = $null
                }
                catch
                {
                    Write-Host "Accessアプリケーションの終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            $this.is_initialized = $false
            Write-Host "AccessDriverのリソース解放が完了しました。" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("AccessDriverのリソースを解放しました")
        }
        catch
        {
            $this.LogError("AccessDriverError_0091", "AccessDriver Disposeエラー: $($_.Exception.Message)")

            Write-Host "AccessDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}