# ORACLEデータベース操作クラス
# Oracle.ManagedDataAccess.Clientを使用してORACLEデータベースに接続・操作

Add-Type -AssemblyName "System.Data.OracleClient"

class OracleDriver
{
    [string]$ConnectionString
    [System.Data.OracleClient.OracleConnection]$Connection
    [System.Data.OracleClient.OracleTransaction]$Transaction
    [bool]$is_connected
    [bool]$is_transaction_active
    [string]$last_error_message
    [hashtable]$connection_parameters
    [string]$sqlplus_path
    [string]$tns_admin_path
    [string]$tns_alias

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
            $this.sqlplus_path = "sqlplus"
            $this.tns_admin_path = $env:TNS_ADMIN
            $this.tns_alias = ""
            
            Write-Host "OracleDriverの初期化が完了しました。"
        }
        catch
        {
            # 初期化に失敗した場合のクリーンアップ
            Write-Host "OracleDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()

            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0001", "OracleDriver初期化エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "OracleDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "OracleDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # TNS_ADMINパスを設定
    [void] SetTnsAdminPath([string]$path)
    {
        try
        {
            $this.tns_admin_path = $path
            $env:TNS_ADMIN = $path
            Write-Host "TNS_ADMINパスを設定しました: $path"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0002", "TNS_ADMINパス設定エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TNS_ADMINパスの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "TNS_ADMINパスの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # SQLPLUSのパスを設定
    [void] SetSqlPlusPath([string]$path)
    {
        try
        {
            $this.sqlplus_path = $path
            Write-Host "SQLPLUSのパスを設定しました: $path"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0003", "SQLPLUSパス設定エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SQLPLUSパスの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "SQLPLUSパスの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # TNS接続用のパラメータを設定
    [void] SetTnsConnectionParameters([string]$tnsAlias, [string]$username, [string]$password, [string]$schema = "")
    {
        try
        {
            $this.tns_alias = $tnsAlias
            $this.connection_parameters = @{
                TnsAlias = $tnsAlias
                Username = $username
                Password = $password
                Schema = $schema
            }

            # TNS接続文字列を構築
            $this.ConnectionString = "Data Source=$tnsAlias;User Id=$username;Password=$password;"

            Write-Host "TNS接続パラメータを設定しました。TNSエイリアス: $tnsAlias"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0004", "TNS接続パラメータ設定エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TNS接続パラメータの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "TNS接続パラメータの設定に失敗しました: $($_.Exception.Message)"
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0005", "接続パラメータ設定エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "接続パラメータの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "接続パラメータの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # TNSエイリアスを使用してSQLPLUSで接続
    [void] ConnectWithSqlPlusTns([string]$username, [string]$password, [string]$tnsAlias)
    {
        try
        {
            if ($this.is_connected)
            {
                Write-Host "既に接続されています。" -ForegroundColor Yellow
                return
            }

            # 接続パラメータを保存
            $this.connection_parameters.Username = $username
            $this.connection_parameters.Password = $password
            $this.connection_parameters.TnsAlias = $tnsAlias
            $this.tns_alias = $tnsAlias

            # SQLPLUS接続文字列を構築（TNSエイリアス使用）
            $connectionString = "$username/$password@$tnsAlias"

            # SQLPLUSプロセスを開始
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.sqlplus_path
            $processInfo.Arguments = $connectionString
            $processInfo.UseShellExecute = $false
            $processInfo.RedirectStandardInput = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.CreateNoWindow = $false

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start()

            # 接続テスト用のSQLを実行
            $process.StandardInput.WriteLine("SELECT 1 FROM DUAL;")
            $process.StandardInput.WriteLine("EXIT;")

            $output = $process.StandardOutput.ReadToEnd()
            $error = $process.StandardError.ReadToEnd()

            $process.WaitForExit()
            $process.Close()

            if ($process.ExitCode -eq 0 -and $output -match "1")
            {
                $this.is_connected = $true
                Write-Host "SQLPLUSでORACLEデータベースに接続しました（TNS）。ユーザ: $username, TNSエイリアス: $tnsAlias" -ForegroundColor Green
            }
            else
            {
                throw "SQLPLUS TNS接続に失敗しました。エラー: $error"
            }
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0006", "SQLPLUS TNS接続エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SQLPLUSでのTNS接続に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "SQLPLUSでのTNS接続に失敗しました: $($_.Exception.Message)"
        }
    }

    # SQLPLUSで接続（ユーザ名、パスワード、サービス名を指定）
    [void] ConnectWithSqlPlus([string]$username, [string]$password, [string]$service_name)
    {
        try
        {
            if ($this.is_connected)
            {
                Write-Host "既に接続されています。" -ForegroundColor Yellow
                return
            }

            # 接続パラメータを保存
            $this.connection_parameters.Username = $username
            $this.connection_parameters.Password = $password
            $this.connection_parameters.ServiceName = $service_name

            # SQLPLUS接続文字列を構築
            $connectionString = "$username/$password@$service_name"
            
            # SQLPLUSプロセスを開始
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.sqlplus_path
            $processInfo.Arguments = $connectionString
            $processInfo.UseShellExecute = $false
            $processInfo.RedirectStandardInput = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.CreateNoWindow = $false

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start()

            # 接続テスト用のSQLを実行
            $process.StandardInput.WriteLine("SELECT 1 FROM DUAL;")
            $process.StandardInput.WriteLine("EXIT;")
            
            $output = $process.StandardOutput.ReadToEnd()
            $error = $process.StandardError.ReadToEnd()
            
            $process.WaitForExit()
            $process.Close()

            if ($process.ExitCode -eq 0 -and $output -match "1")
            {
                $this.is_connected = $true
                Write-Host "SQLPLUSでORACLEデータベースに接続しました。ユーザ: $username, サービス: $service_name" -ForegroundColor Green
            }
            else
            {
                throw "SQLPLUS接続に失敗しました。エラー: $error"
            }
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0007", "SQLPLUS接続エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SQLPLUSでの接続に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "SQLPLUSでの接続に失敗しました: $($_.Exception.Message)"
        }
    }

    # SQLPLUSで接続（接続文字列を指定）
    [void] ConnectWithSqlPlusString([string]$connectionString)
    {
        try
        {
            if ($this.is_connected)
            {
                Write-Host "既に接続されています。" -ForegroundColor Yellow
                return
            }

            # SQLPLUSプロセスを開始
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.sqlplus_path
            $processInfo.Arguments = $connectionString
            $processInfo.UseShellExecute = $false
            $processInfo.RedirectStandardInput = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.CreateNoWindow = $false

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start()

            # 接続テスト用のSQLを実行
            $process.StandardInput.WriteLine("SELECT 1 FROM DUAL;")
            $process.StandardInput.WriteLine("EXIT;")
            
            $output = $process.StandardOutput.ReadToEnd()
            $error = $process.StandardError.ReadToEnd()
            
            $process.WaitForExit()
            $process.Close()

            if ($process.ExitCode -eq 0 -and $output -match "1")
            {
                $this.is_connected = $true
                Write-Host "SQLPLUSでORACLEデータベースに接続しました。接続文字列: $connectionString" -ForegroundColor Green
            }
            else
            {
                throw "SQLPLUS接続に失敗しました。エラー: $error"
            }
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0008", "SQLPLUS接続文字列接続エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SQLPLUSでの接続に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "SQLPLUSでの接続に失敗しました: $($_.Exception.Message)"
        }
    }

    # TNSエイリアスを使用してSQLPLUSでSQLを実行
    [string] ExecuteSqlPlusTns([string]$sql, [string]$username, [string]$password, [string]$tnsAlias)
    {
        try
        {
            # SQLの末尾にセミコロンまたはスラッシュがない場合は追加
            $trimmedSql = $sql.TrimEnd()
            if (-not ($trimmedSql.EndsWith(";") -or $trimmedSql.EndsWith("/")))
            {
                # PL/SQLブロックやストアドプロシージャの場合はスラッシュ、それ以外はセミコロン
                if ($trimmedSql -match "(?i)^\s*(CREATE|ALTER|DROP|BEGIN|DECLARE)" -or $trimmedSql -match "(?i)\bEND\s*$")
                {
                    $sql = $sql + "`n/"
                }
                else
                {
                    $sql = $sql + ";"
                }
            }

            # TNS_ADMINが設定されていることを確認
            if (-not [string]::IsNullOrEmpty($this.tns_admin_path))
            {
                $env:TNS_ADMIN = $this.tns_admin_path
            }

            # SQLPLUSプロセスを開始（-Sオプションでサイレントモード、TNSエイリアス使用）
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.sqlplus_path
            $processInfo.Arguments = "-S $username/$password@$tnsAlias"
            $processInfo.UseShellExecute = $false
            $processInfo.RedirectStandardInput = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.CreateNoWindow = $true
            $processInfo.WorkingDirectory = [System.IO.Directory]::GetCurrentDirectory()

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo

            # プロセス開始のエラーチェック
            if (-not $process.Start())
            {
                throw "SQLPLUSプロセスの開始に失敗しました。"
            }

            # SQLPLUSの設定とSQLを実行
            $process.StandardInput.WriteLine("SET PAGESIZE 0")
            $process.StandardInput.WriteLine("SET FEEDBACK OFF")
            $process.StandardInput.WriteLine("SET HEADING OFF")
            $process.StandardInput.WriteLine("SET LINESIZE 1000")
            $process.StandardInput.WriteLine("SET TRIMSPOOL ON")
            $process.StandardInput.WriteLine("SET ECHO OFF")
            $process.StandardInput.WriteLine("WHENEVER SQLERROR EXIT SQL.SQLCODE")
            $process.StandardInput.WriteLine($sql)
            $process.StandardInput.WriteLine("EXIT;")
            $process.StandardInput.Close()

            $output = $process.StandardOutput.ReadToEnd()
            $error = $process.StandardError.ReadToEnd()

            # タイムアウトを設定（30秒）
            $timeout = 30000
            if (-not $process.WaitForExit($timeout))
            {
                $process.Kill()
                throw "SQLPLUSの実行がタイムアウトしました。"
            }

            $exitCode = $process.ExitCode
            $process.Close()

            # エラーチェック（終了コードとエラー出力の両方を確認）
            if ($exitCode -ne 0 -or -not [string]::IsNullOrEmpty($error))
            {
                $errorMessage = if (-not [string]::IsNullOrEmpty($error)) { $error } else { "終了コード: $exitCode" }
                throw "SQLPLUSでのSQL実行に失敗しました。エラー: $errorMessage"
            }

            # ORA-エラーやSP2-エラーのチェック
            if ($output -match "ORA-\d+" -or $output -match "SP2-\d+")
            {
                throw "SQLPLUSでエラーが発生しました: $output"
            }

            Write-Host "SQLPLUSでSQLを実行しました（TNS）。" -ForegroundColor Green
            return $output.Trim()
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0009", "SQLPLUS TNS実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SQLPLUSでのSQL実行に失敗しました（TNS）: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "SQLPLUSでのSQL実行に失敗しました（TNS）: $($_.Exception.Message)"
        }
    }

    # SQLPLUSでSQLを実行
    [string] ExecuteSqlPlus([string]$sql, [string]$username, [string]$password, [string]$service_name)
    {
        try
        {
            # SQLの末尾にセミコロンまたはスラッシュがない場合は追加
            $trimmedSql = $sql.TrimEnd()
            if (-not ($trimmedSql.EndsWith(";") -or $trimmedSql.EndsWith("/")))
            {
                # PL/SQLブロックやストアドプロシージャの場合はスラッシュ、それ以外はセミコロン
                if ($trimmedSql -match "(?i)^\s*(CREATE|ALTER|DROP|BEGIN|DECLARE)" -or $trimmedSql -match "(?i)\bEND\s*$")
                {
                    $sql = $sql + "`n/"
                }
                else
                {
                    $sql = $sql + ";"
                }
            }

            # SQLPLUSプロセスを開始（-Sオプションでサイレントモード）
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.sqlplus_path
            $processInfo.Arguments = "-S $username/$password@$service_name"
            $processInfo.UseShellExecute = $false
            $processInfo.RedirectStandardInput = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.CreateNoWindow = $true
            $processInfo.WorkingDirectory = [System.IO.Directory]::GetCurrentDirectory()

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo

            # プロセス開始のエラーチェック
            if (-not $process.Start())
            {
                throw "SQLPLUSプロセスの開始に失敗しました。"
            }

            # SQLPLUSの設定とSQLを実行
            $process.StandardInput.WriteLine("SET PAGESIZE 0")
            $process.StandardInput.WriteLine("SET FEEDBACK OFF")
            $process.StandardInput.WriteLine("SET HEADING OFF")
            $process.StandardInput.WriteLine("SET LINESIZE 1000")
            $process.StandardInput.WriteLine("SET TRIMSPOOL ON")
            $process.StandardInput.WriteLine("SET ECHO OFF")
            $process.StandardInput.WriteLine("WHENEVER SQLERROR EXIT SQL.SQLCODE")
            $process.StandardInput.WriteLine($sql)
            $process.StandardInput.WriteLine("EXIT;")
            $process.StandardInput.Close()

            $output = $process.StandardOutput.ReadToEnd()
            $error = $process.StandardError.ReadToEnd()

            # タイムアウトを設定（30秒）
            $timeout = 30000
            if (-not $process.WaitForExit($timeout))
            {
                $process.Kill()
                throw "SQLPLUSの実行がタイムアウトしました。"
            }

            $exitCode = $process.ExitCode
            $process.Close()

            # エラーチェック（終了コードとエラー出力の両方を確認）
            if ($exitCode -ne 0 -or -not [string]::IsNullOrEmpty($error))
            {
                $errorMessage = if (-not [string]::IsNullOrEmpty($error)) { $error } else { "終了コード: $exitCode" }
                throw "SQLPLUSでのSQL実行に失敗しました。エラー: $errorMessage"
            }

            # ORA-エラーやSP2-エラーのチェック
            if ($output -match "ORA-\d+" -or $output -match "SP2-\d+")
            {
                throw "SQLPLUSでエラーが発生しました: $output"
            }

            Write-Host "SQLPLUSでSQLを実行しました。" -ForegroundColor Green
            return $output.Trim()
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0010", "SQLPLUS実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SQLPLUSでのSQL実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "SQLPLUSでのSQL実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # TNSエイリアスを使用してデータベースに接続
    [void] ConnectTns([string]$tnsAlias, [string]$username, [string]$password, [string]$schema = "")
    {
        try
        {
            if ($this.is_connected)
            {
                Write-Host "既に接続されています。" -ForegroundColor Yellow
                return
            }

            # TNS接続パラメータを設定
            $this.SetTnsConnectionParameters($tnsAlias, $username, $password, $schema)

            $this.Connection = New-Object System.Data.OracleClient.OracleConnection($this.ConnectionString)
            $this.Connection.Open()
            $this.is_connected = $true

            Write-Host "ORACLEデータベースに接続しました（TNS: $tnsAlias）。" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0011", "TNSデータベース接続エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TNSを使用したデータベースへの接続に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "TNSを使用したデータベースへの接続に失敗しました: $($_.Exception.Message)"
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0012", "データベース接続エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "データベースへの接続に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0013", "接続切断エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "接続の切断に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0014", "SELECT実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SELECT文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0015", "INSERT実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "INSERT文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0016", "UPDATE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "UPDATE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0017", "DELETE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "DELETE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0018", "MERGE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "MERGE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "MERGE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # CALL文を実行（ストアドプロシージャ呼び出し）
    [void] ExecuteCall([string]$procedureName, [hashtable]$parameters = @{})
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($procedureName, $this.Connection)
            $command.CommandType = [System.Data.CommandType]::StoredProcedure

            # パラメータを追加
            foreach ($key in $parameters.Keys)
            {
                $param = $command.Parameters.Add($key, $parameters[$key])
            }

            $command.ExecuteNonQuery()

            Write-Host "CALL文を実行しました。プロシージャ: $procedureName" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0019", "CALL実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "CALL文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "CALL文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # EXPLAIN PLANを実行
    [System.Data.DataTable] ExecuteExplainPlan([string]$sql, [string]$statementId = "PLAN_" + [System.Guid]::NewGuid().ToString())
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            # EXPLAIN PLANを実行
            $explainSql = "EXPLAIN PLAN SET STATEMENT_ID = '$statementId' FOR $sql"
            $command = New-Object System.Data.OracleClient.OracleCommand($explainSql, $this.Connection)
            $command.ExecuteNonQuery()

            # 実行計画を取得
            $selectPlan = @"
SELECT LPAD(' ', 2 * (LEVEL - 1)) || OPERATION || ' ' || OPTIONS AS OPERATION,
       OBJECT_NAME,
       COST,
       CARDINALITY,
       BYTES
FROM PLAN_TABLE
START WITH ID = 0 AND STATEMENT_ID = '$statementId'
CONNECT BY PRIOR ID = PARENT_ID AND STATEMENT_ID = '$statementId'
ORDER SIBLINGS BY ID
"@

            $planCommand = New-Object System.Data.OracleClient.OracleCommand($selectPlan, $this.Connection)
            $adapter = New-Object System.Data.OracleClient.OracleDataAdapter($planCommand)
            $dataTable = New-Object System.Data.DataTable
            $adapter.Fill($dataTable)

            # PLAN_TABLEから該当データを削除
            $deletePlan = "DELETE FROM PLAN_TABLE WHERE STATEMENT_ID = '$statementId'"
            $deleteCommand = New-Object System.Data.OracleClient.OracleCommand($deletePlan, $this.Connection)
            $deleteCommand.ExecuteNonQuery()

            Write-Host "EXPLAIN PLANを実行しました。" -ForegroundColor Green
            return $dataTable
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0020", "EXPLAIN PLAN実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "EXPLAIN PLANの実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "EXPLAIN PLANの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # LOCK TABLEを実行
    [void] ExecuteLockTable([string]$tableName, [string]$lockMode = "EXCLUSIVE")
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $sql = "LOCK TABLE $tableName IN $lockMode MODE"

            if (-not $this.is_transaction_active)
            {
                Write-Host "トランザクションが開始されていません。自動的に開始します。" -ForegroundColor Yellow
                $this.BeginTransaction()
            }

            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection, $this.Transaction)
            $command.ExecuteNonQuery()

            Write-Host "LOCK TABLEを実行しました。テーブル: $tableName, モード: $lockMode" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0021", "LOCK TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "LOCK TABLEの実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "LOCK TABLEの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # DDL (Data Definition Language) 操作
    # CREATE, ALTER, DROP, TRUNCATE, RENAME, FLASHBACK, ANALYZE, AUDIT, COMMENT, ASSOCIATE/DISASSOCIATE STATISTICS, PURGE, NOAUDIT
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0022", "CREATE TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "CREATE TABLE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0023", "ALTER TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ALTER TABLE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0024", "DROP TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "DROP TABLE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0025", "CREATE INDEX実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "CREATE INDEX文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0026", "CREATE VIEW実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "CREATE VIEW文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0027", "CREATE SEQUENCE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "CREATE SEQUENCE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "CREATE SEQUENCE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # TRUNCATE TABLEを実行
    [void] ExecuteTruncate([string]$tableName)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $sql = "TRUNCATE TABLE $tableName"
            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()

            Write-Host "TRUNCATE TABLEを実行しました。テーブル: $tableName" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0028", "TRUNCATE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TRUNCATE TABLEの実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "TRUNCATE TABLEの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # RENAMEを実行
    [void] ExecuteRename([string]$oldName, [string]$newName)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            $sql = "RENAME $oldName TO $newName"
            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()

            Write-Host "RENAMEを実行しました。$oldName -> $newName" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0029", "RENAME実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "RENAMEの実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "RENAMEの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # DCL (Data Control Language) 操作
    # GRANT, REVOKE
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0030", "GRANT実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "GRANT文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0031", "REVOKE実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "REVOKE文の実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "REVOKE文の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # TCL (Transaction Control Language) 操作
    # COMMIT, ROLLBACK, SAVEPOINT, SET TRANSACTION
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0032", "トランザクション開始エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "トランザクションの開始に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0033", "トランザクションコミットエラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "トランザクションのコミットに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0034", "トランザクションロールバックエラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "トランザクションのロールバックに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0035", "SAVEPOINT作成エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SAVEPOINTの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0036", "SAVEPOINTロールバックエラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SAVEPOINTへのロールバックに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "SAVEPOINTへのロールバックに失敗しました: $($_.Exception.Message)"
        }
    }

    # SET TRANSACTIONを実行
    [void] SetTransaction([string]$isolationLevel = "READ COMMITTED")
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "データベースに接続されていません。Connect()を先に実行してください。"
            }

            if ($this.is_transaction_active)
            {
                throw "既にトランザクションが開始されています。先に完了またはロールバックしてください。"
            }

            $sql = "SET TRANSACTION ISOLATION LEVEL $isolationLevel"
            $command = New-Object System.Data.OracleClient.OracleCommand($sql, $this.Connection)
            $command.ExecuteNonQuery()

            Write-Host "SET TRANSACTIONを実行しました。分離レベル: $isolationLevel" -ForegroundColor Green
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0037", "SET TRANSACTION実行エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "SET TRANSACTIONの実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "SET TRANSACTIONの実行に失敗しました: $($_.Exception.Message)"
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0038", "テーブル一覧取得エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テーブル一覧の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0039", "テーブル構造取得エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テーブル構造の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleError_0040", "OracleDriver破棄エラー: $($_.Exception.Message)", "OracleDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "OracleDriverの破棄中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "OracleDriverの破棄中にエラーが発生しました: $($_.Exception.Message)"
        }
    }
}

Write-Host "OracleDriverライブラリが正常にインポートされました。" -ForegroundColor Green 