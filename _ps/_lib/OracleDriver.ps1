# ORACLEデータベース操作クラス
# Oracle.ManagedDataAccess.Clientを使用してORACLEデータベースに接続・操作

# 共通ライブラリをインポート
#. "$PSScriptRoot\Common.ps1"
#$Common = New-Object -TypeName 'Common'

class OracleDriver
{
    [string]$ConnectionString
    [System.Data.OracleClient.OracleConnection]$Connection
    [System.Data.OracleClient.OracleTransaction]$Transaction
    [bool]$is_connected
    [bool]$is_transaction_active
    [string]$last_error_message
    [hashtable]$connection_parameters

    # ========================================
    # 初期化・接続関連
    # ========================================

    OracleDriver()
    {
        try
        {
            $this.is_connected = $false
            $this.is_transaction_active = $false
            $this.last_error_message = ""
            $this.connection_parameters = @{}
            
            Write-Host "OracleDriverの初期化が完了しました。"
        }
        catch
        {
            $global:Common.HandleError("6001", "OracleDriver初期化エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "OracleDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続パラメータを設定
    [void] SetConnectionParameters([string]$server, [string]$port, [string]$service_name, [string]$username, [string]$password, [string]$schema = "")
    {
        try
        {
            $this.connection_parameters = @{
                Server = $server
                Port = $port
                ServiceName = $service_name
                Username = $username
                Password = $password
                Schema = $schema
            }
            
            # 接続文字列を構築
            $this.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$server)(PORT=$port))(CONNECT_DATA=(SERVICE_NAME=$service_name)));User Id=$username;Password=$password;"
            
            Write-Host "接続パラメータを設定しました。"
        }
        catch
        {
            $global:Common.HandleError("6002", "接続パラメータ設定エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "接続パラメータの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # データベースに接続
    [void] Connect()
    {
        try
        {
            if ($this.is_connected)
            {
                Write-Host "既に接続されています。" -ForegroundColor Yellow
                return
            }

            if ([string]::IsNullOrEmpty($this.ConnectionString))
            {
                throw "接続文字列が設定されていません。SetConnectionParameters()を先に実行してください。"
            }

            $this.Connection = New-Object System.Data.OracleClient.OracleConnection($this.ConnectionString)
            $this.Connection.Open()
            $this.is_connected = $true
            
            Write-Host "ORACLEデータベースに接続しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6003", "データベース接続エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "データベースへの接続に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続を切断
    [void] Disconnect()
    {
        try
        {
            if ($this.is_transaction_active)
            {
                $this.RollbackTransaction()
            }

            if ($this.Connection -and $this.Connection.State -eq 'Open')
            {
                $this.Connection.Close()
                $this.Connection.Dispose()
                $this.is_connected = $false
                Write-Host "データベース接続を切断しました。" -ForegroundColor Green
            }
        }
        catch
        {
            $global:Common.HandleError("6004", "接続切断エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "接続の切断に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # DML (Data Manipulation Language) 操作
    # ========================================

    # SELECT文を実行
    [System.Data.DataTable] ExecuteSelect([string]$sql, [hashtable]$parameters = @{})
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            
            # パラメータを追加
            foreach ($key in $parameters.Keys)
            {
                $param = $command.Parameters.Add($key, $parameters[$key])
            }

            $adapter = New-Object System.Data.OracleClient.OracleDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $adapter.Fill($dataTable)
            
            Write-Host "SELECT文を実行しました。結果行数: $($dataTable.Rows.Count)" -ForegroundColor Green
            return $dataTable
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6005", "SELECT実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "SELECT文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # INSERT文を実行
    [int] ExecuteInsert([string]$sql, [hashtable]$parameters = @{})
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            
            # パラメータを追加
            foreach ($key in $parameters.Keys)
            {
                $param = $command.Parameters.Add($key, $parameters[$key])
            }

            $affectedRows = $command.ExecuteNonQuery()
            
            Write-Host "INSERT文を実行しました。影響を受けた行数: $affectedRows" -ForegroundColor Green
            return $affectedRows
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6006", "INSERT実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "INSERT文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # UPDATE文を実行
    [int] ExecuteUpdate([string]$sql, [hashtable]$parameters = @{})
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            
            # パラメータを追加
            foreach ($key in $parameters.Keys)
            {
                $param = $command.Parameters.Add($key, $parameters[$key])
            }

            $affectedRows = $command.ExecuteNonQuery()
            
            Write-Host "UPDATE文を実行しました。影響を受けた行数: $affectedRows" -ForegroundColor Green
            return $affectedRows
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6007", "UPDATE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "UPDATE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # DELETE文を実行
    [int] ExecuteDelete([string]$sql, [hashtable]$parameters = @{})
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            
            # パラメータを追加
            foreach ($key in $parameters.Keys)
            {
                $param = $command.Parameters.Add($key, $parameters[$key])
            }

            $affectedRows = $command.ExecuteNonQuery()
            
            Write-Host "DELETE文を実行しました。影響を受けた行数: $affectedRows" -ForegroundColor Green
            return $affectedRows
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6008", "DELETE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "DELETE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # MERGE文を実行
    [int] ExecuteMerge([string]$sql, [hashtable]$parameters = @{})
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            
            # パラメータを追加
            foreach ($key in $parameters.Keys)
            {
                $param = $command.Parameters.Add($key, $parameters[$key])
            }

            $affectedRows = $command.ExecuteNonQuery()
            
            Write-Host "MERGE文を実行しました。影響を受けた行数: $affectedRows" -ForegroundColor Green
            return $affectedRows
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6009", "MERGE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "MERGE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # DDL (Data Definition Language) 操作
    # ========================================

    # CREATE TABLE文を実行
    [void] ExecuteCreateTable([string]$sql)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "CREATE TABLE文を実行しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6010", "CREATE TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "CREATE TABLE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ALTER TABLE文を実行
    [void] ExecuteAlterTable([string]$sql)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "ALTER TABLE文を実行しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6011", "ALTER TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "ALTER TABLE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # DROP TABLE文を実行
    [void] ExecuteDropTable([string]$tableName)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $sql = "DROP TABLE $tableName"
            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "DROP TABLE文を実行しました。テーブル: $tableName" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6012", "DROP TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "DROP TABLE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # CREATE INDEX文を実行
    [void] ExecuteCreateIndex([string]$sql)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "CREATE INDEX文を実行しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6013", "CREATE INDEX実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "CREATE INDEX文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # CREATE VIEW文を実行
    [void] ExecuteCreateView([string]$sql)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "CREATE VIEW文を実行しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6014", "CREATE VIEW実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "CREATE VIEW文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # CREATE SEQUENCE文を実行
    [void] ExecuteCreateSequence([string]$sql)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "CREATE SEQUENCE文を実行しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6015", "CREATE SEQUENCE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "CREATE SEQUENCE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # DCL (Data Control Language) 操作
    # ========================================

    # GRANT文を実行
    [void] ExecuteGrant([string]$sql)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "GRANT文を実行しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6016", "GRANT実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "GRANT文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # REVOKE文を実行
    [void] ExecuteRevoke([string]$sql)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()
            
            Write-Host "REVOKE文を実行しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6017", "REVOKE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "REVOKE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # TCL (Transaction Control Language) 操作
    # ========================================

    # トランザクションを開始
    [void] BeginTransaction()
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            if ($this.is_transaction_active)
            {
                Write-Host "既にトランザクションが開始されています。" -ForegroundColor Yellow
                return
            }

            $this.Transaction = $this.Connection.BeginTransaction()
            $this.is_transaction_active = $true
            
            Write-Host "トランザクションを開始しました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6018", "トランザクション開始エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "トランザクションの開始に失敗しました: $($_.Exception.Message)"
        }
    }

    # トランザクションをコミット
    [void] CommitTransaction()
    {
        try
        {
            if (-not $this.is_transaction_active)
            {
                Write-Host "アクティブなトランザクションがありません。" -ForegroundColor Yellow
                return
            }

            $this.Transaction.Commit()
            $this.Transaction.Dispose()
            $this.is_transaction_active = $false
            
            Write-Host "トランザクションをコミットしました。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6019", "トランザクションコミットエラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "トランザクションのコミットに失敗しました: $($_.Exception.Message)"
        }
    }

    # トランザクションをロールバック
    [void] RollbackTransaction()
    {
        try
        {
            if (-not $this.is_transaction_active)
            {
                Write-Host "アクティブなトランザクションがありません。" -ForegroundColor Yellow
                return
            }

            $this.Transaction.Rollback()
            $this.Transaction.Dispose()
            $this.is_transaction_active = $false
            
            Write-Host "トランザクションをロールバックしました。" -ForegroundColor Yellow
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6020", "トランザクションロールバックエラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "トランザクションのロールバックに失敗しました: $($_.Exception.Message)"
        }
    }

    # SAVEPOINTを作成
    [void] CreateSavepoint([string]$savepointName)
    {
        try
        {
            if (-not $this.is_transaction_active)
            {
                throw "アクティブなトランザクションがありません。BeginTransaction()を先に実行してください。"
            }

            $sql = "SAVEPOINT $savepointName"
            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection, $this.Transaction)
            $command.ExecuteNonQuery()
            
            Write-Host "SAVEPOINTを作成しました: $savepointName" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6021", "SAVEPOINT作成エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "SAVEPOINTの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # SAVEPOINTにロールバック
    [void] RollbackToSavepoint([string]$savepointName)
    {
        try
        {
            if (-not $this.is_transaction_active)
            {
                throw "アクティブなトランザクションがありません。BeginTransaction()を先に実行してください。"
            }

            $sql = "ROLLBACK TO SAVEPOINT $savepointName"
            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection, $this.Transaction)
            $command.ExecuteNonQuery()
            
            Write-Host "SAVEPOINTにロールバックしました: $savepointName" -ForegroundColor Yellow
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6022", "SAVEPOINTロールバックエラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "SAVEPOINTへのロールバックに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ユーティリティメソッド
    # ========================================

    # テーブル一覧を取得
    [System.Data.DataTable] GetTableList([string]$schema = "")
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $sql = @"
SELECT TABLE_NAME, TABLESPACE_NAME, STATUS, LAST_ANALYZED
FROM ALL_TABLES
WHERE OWNER = :schema
ORDER BY TABLE_NAME
"@

            $parameters = @{
                ":schema" = if ([string]::IsNullOrEmpty($schema)) { $this.connection_parameters.Username } else { $schema }
            }

            return $this.ExecuteSelect($sql, $parameters)
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6023", "テーブル一覧取得エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "テーブル一覧の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # テーブル構造を取得
    [System.Data.DataTable] GetTableStructure([string]$tableName, [string]$schema = "")
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $sql = @"
SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE, DATA_DEFAULT
FROM ALL_TAB_COLUMNS
WHERE TABLE_NAME = :tableName AND OWNER = :schema
ORDER BY COLUMN_ID
"@

            $parameters = @{
                ":tableName" = $tableName
                ":schema" = if ([string]::IsNullOrEmpty($schema)) { $this.connection_parameters.Username } else { $schema }
            }

            return $this.ExecuteSelect($sql, $parameters)
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            $global:Common.HandleError("6024", "テーブル構造取得エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
            throw "テーブル構造の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続状態を確認
    [bool] IsConnected()
    {
        return $this.is_connected -and $this.Connection.State -eq 'Open'
    }

    # トランザクション状態を確認
    [bool] IsTransactionActive()
    {
        return $this.is_transaction_active
    }

    # 最後のエラーメッセージを取得
    [string] GetLastErrorMessage()
    {
        return $this.last_error_message
    }

    # デストラクタ
    [void] Dispose()
    {
        try
        {
            $this.Disconnect()
        }
        catch
        {
            Write-Host "OracleDriverの破棄中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "OracleDriverライブラリが正常にインポートされました。" -ForegroundColor Green 