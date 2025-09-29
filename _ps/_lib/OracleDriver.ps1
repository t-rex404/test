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

    # ログファイルパス（共有可能）
    static [string]$NormalLogFile = ".\OracleDriver_$($env:USERNAME)_Normal.log"
    static [string]$ErrorLogFile = ".\OracleDriver_$($env:USERNAME)_Error.log"

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
            # 環境変数TNS_ADMINがNULLの場合は空文字列を設定
            $this.tns_admin_path = if ($null -eq $env:TNS_ADMIN) { "" } else { $env:TNS_ADMIN }
            $this.tns_alias = ""

            Write-Host "OracleDriverの初期化が完了しました。"

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("OracleDriverの初期化が完了しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] OracleDriverの初期化が完了しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
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
                    $global:Common.HandleError("OracleDriverError_0001", "OracleDriver初期化エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("TNS_ADMINパスを設定しました: $path", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] TNS_ADMINパスを設定しました: $path" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0002", "TNS_ADMINパス設定エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("SQLPLUSのパスを設定しました: $path", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SQLPLUSのパスを設定しました: $path" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0003", "SQLPLUSパス設定エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("TNS接続パラメータを設定しました。TNSエイリアス: $tnsAlias", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] TNS接続パラメータを設定しました。TNSエイリアス: $tnsAlias" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0010", "TNS接続パラメータ設定エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("接続パラメータを設定しました。サーバー: $server, ポート: $port, サービス名: $service_name", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 接続パラメータを設定しました。サーバー: $server, ポート: $port, サービス名: $service_name" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0011", "接続パラメータ設定エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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
            $sqlPlusConnectionStr = "$username/$password@$tnsAlias"

            # SQLPLUSプロセスを開始
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.sqlplus_path
            $processInfo.Arguments = $sqlPlusConnectionStr
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
            $exitCode = $process.ExitCode
            $process.Close()

            if ($exitCode -eq 0 -and $output -match "1")
            {
                $this.is_connected = $true
                Write-Host "SQLPLUSでORACLEデータベースに接続しました（TNS）。ユーザ: $username, TNSエイリアス: $tnsAlias" -ForegroundColor Green

                # 正常ログ出力
                if ($global:Common)
                {
                    $global:Common.WriteLog("SQLPLUSでORACLEデータベースに接続しました（TNS）。ユーザ: $username, TNSエイリアス: $tnsAlias", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SQLPLUSでORACLEデータベースに接続しました（TNS）。ユーザ: $username, TNSエイリアス: $tnsAlias" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
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
                    $global:Common.HandleError("OracleDriverError_0012", "SQLPLUS TNS接続エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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
            $sqlPlusConnectionStr = "$username/$password@$service_name"

            # SQLPLUSプロセスを開始
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.sqlplus_path
            $processInfo.Arguments = $sqlPlusConnectionStr
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
            $exitCode = $process.ExitCode
            $process.Close()

            if ($exitCode -eq 0 -and $output -match "1")
            {
                $this.is_connected = $true
                Write-Host "SQLPLUSでORACLEデータベースに接続しました。ユーザ: $username, サービス: $service_name" -ForegroundColor Green

                # 正常ログ出力
                if ($global:Common)
                {
                    $global:Common.WriteLog("SQLPLUSでORACLEデータベースに接続しました。ユーザ: $username, サービス: $service_name", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SQLPLUSでORACLEデータベースに接続しました。ユーザ: $username, サービス: $service_name" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
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
                    $global:Common.HandleError("OracleDriverError_0013", "SQLPLUS接続エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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
    [void] ConnectWithSqlPlusString([string]$sqlPlusConnStr)
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
            $processInfo.Arguments = $sqlPlusConnStr
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
            $exitCode = $process.ExitCode
            $process.Close()

            if ($exitCode -eq 0 -and $output -match "1")
            {
                $this.is_connected = $true
                Write-Host "SQLPLUSでORACLEデータベースに接続しました。接続文字列: $sqlPlusConnStr" -ForegroundColor Green

                # 正常ログ出力
                if ($global:Common)
                {
                    $global:Common.WriteLog("SQLPLUSでORACLEデータベースに接続しました", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SQLPLUSでORACLEデータベースに接続しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
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
                    $global:Common.HandleError("OracleDriverError_0014", "SQLPLUS接続文字列接続エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("SQLPLUSでSQLを実行しました（TNS）。TNSエイリアス: $tnsAlias", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SQLPLUSでSQLを実行しました（TNS）。TNSエイリアス: $tnsAlias" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0020", "SQLPLUS TNS実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("SQLPLUSでSQLを実行しました。サービス名: $service_name", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SQLPLUSでSQLを実行しました。サービス名: $service_name" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0021", "SQLPLUS実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("ORACLEデータベースに接続しました（TNS: $tnsAlias）", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ORACLEデータベースに接続しました（TNS: $tnsAlias）" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0015", "TNSデータベース接続エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("ORACLEデータベースに接続しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ORACLEデータベースに接続しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0016", "データベース接続エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

                # 正常ログ出力
                if ($global:Common)
                {
                    $global:Common.WriteLog("データベース接続を切断しました", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] データベース接続を切断しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0017", "接続切断エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("SELECT文を実行しました。結果行数: $($dataTable.Rows.Count)", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SELECT文を実行しました。結果行数: $($dataTable.Rows.Count)" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0030", "SELECT実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("INSERT文を実行しました。影響を受けた行数: $affectedRows", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] INSERT文を実行しました。影響を受けた行数: $affectedRows" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0031", "INSERT実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("UPDATE文を実行しました。影響を受けた行数: $affectedRows", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] UPDATE文を実行しました。影響を受けた行数: $affectedRows" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0032", "UPDATE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("DELETE文を実行しました。影響を受けた行数: $affectedRows", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] DELETE文を実行しました。影響を受けた行数: $affectedRows" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0033", "DELETE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("MERGE文を実行しました。影響を受けた行数: $affectedRows", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] MERGE文を実行しました。影響を受けた行数: $affectedRows" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0034", "MERGE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("ストアドプロシージャを実行しました: $procedureName", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ストアドプロシージャを実行しました: $procedureName" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0035", "CALL実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("EXPLAIN PLANを実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] EXPLAIN PLANを実行しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

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
                    $global:Common.HandleError("OracleDriverError_0036", "EXPLAIN PLAN実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("テーブルをロックしました: $tableName (モード: $lockMode)", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] テーブルをロックしました: $tableName (モード: $lockMode)" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0037", "LOCK TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("CREATE TABLE文を実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] CREATE TABLE文を実行しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0040", "CREATE TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("ALTER TABLE文を実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ALTER TABLE文を実行しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0041", "ALTER TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("テーブルを削除しました: $tableName", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] テーブルを削除しました: $tableName" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0042", "DROP TABLE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("インデックスを作成しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] インデックスを作成しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0043", "CREATE INDEX実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("ビューを作成しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ビューを作成しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0044", "CREATE VIEW実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("シーケンスを作成しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] シーケンスを作成しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0045", "CREATE SEQUENCE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("テーブルをトランケートしました: $tableName", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] テーブルをトランケートしました: $tableName" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0046", "TRUNCATE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("オブジェクト名を変更しました: $oldName -> $newName", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] オブジェクト名を変更しました: $oldName -> $newName" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0047", "RENAME実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("権限を付与しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 権限を付与しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0048", "GRANT実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("権限を取り消しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 権限を取り消しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0049", "REVOKE実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("トランザクションを開始しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] トランザクションを開始しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0050", "トランザクション開始エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("トランザクションをコミットしました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] トランザクションをコミットしました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0051", "トランザクションコミットエラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("トランザクションをロールバックしました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] トランザクションをロールバックしました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0052", "トランザクションロールバックエラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("SAVEPOINTを作成しました: $savepointName", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SAVEPOINTを作成しました: $savepointName" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0053", "SAVEPOINT作成エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("SAVEPOINTにロールバックしました: $savepointName", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] SAVEPOINTにロールバックしました: $savepointName" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0054", "SAVEPOINTロールバックエラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("トランザクション分離レベルを設定しました: $isolationLevel", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] トランザクション分離レベルを設定しました: $isolationLevel" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                    $global:Common.HandleError("OracleDriverError_0055", "SET TRANSACTION実行エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

    # tnsnames.oraのパスを取得
    [string] GetTnsNamesPath()
    {
        try
        {
            Write-Host "tnsnames.oraファイルを検索しています..." -ForegroundColor Yellow

            # 1. まずレジストリからOracleホームパスを取得
            $tnsNamesPath = $this.GetTnsNamesFromRegistry()

            if (-not [string]::IsNullOrEmpty($tnsNamesPath) -and (Test-Path $tnsNamesPath))
            {
                Write-Host "レジストリから tnsnames.ora を発見しました: $tnsNamesPath" -ForegroundColor Green

                # 正常ログ出力
                if ($global:Common)
                {
                    $global:Common.WriteLog("tnsnames.oraファイルを発見しました: $tnsNamesPath", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] tnsnames.oraファイルを発見しました: $tnsNamesPath" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }

                return $tnsNamesPath
            }

            # 2. レジストリから見つからない場合はファイルシステムを検索
            Write-Host "レジストリから見つからないため、ファイルシステムを検索します..." -ForegroundColor Yellow
            $foundFiles = $this.FindTnsNamesFiles()

            if ($foundFiles.Count -eq 0)
            {
                throw "tnsnames.ora ファイルが見つかりませんでした。"
            }
            elseif ($foundFiles.Count -eq 1)
            {
                Write-Host "tnsnames.ora を発見しました: $($foundFiles[0])" -ForegroundColor Green

                # 正常ログ出力
                if ($global:Common)
                {
                    $global:Common.WriteLog("tnsnames.oraファイルを発見しました: $($foundFiles[0])", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] tnsnames.oraファイルを発見しました: $($foundFiles[0])" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }

                return $foundFiles[0]
            }
            else
            {
                # 複数見つかった場合は選択を促す
                Write-Host "複数の tnsnames.ora ファイルが見つかりました。" -ForegroundColor Yellow
                return $this.SelectTnsNamesFile($foundFiles)
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
                    $global:Common.HandleError("OracleDriverError_0082", "tnsnames.ora検索エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "tnsnames.oraの検索に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "tnsnames.oraの検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # レジストリからtnsnames.oraのパスを取得（プライベートヘルパーメソッド）
    hidden [string] GetTnsNamesFromRegistry()
    {
        try
        {
            # Oracle関連のレジストリキーをチェック
            $registryPaths = @(
                "HKLM:\SOFTWARE\ORACLE",
                "HKLM:\SOFTWARE\WOW6432Node\ORACLE",
                "HKCU:\SOFTWARE\ORACLE",
                "HKCU:\SOFTWARE\WOW6432Node\ORACLE"
            )

            foreach ($regPath in $registryPaths)
            {
                if (Test-Path $regPath)
                {
                    # Oracleホームを検索
                    $oracleKeys = Get-ChildItem $regPath -ErrorAction SilentlyContinue

                    foreach ($key in $oracleKeys)
                    {
                        # ORACLE_HOMEプロパティを探す
                        $oracleHome = $null
                        try
                        {
                            $oracleHome = (Get-ItemProperty -Path $key.PSPath -Name "ORACLE_HOME" -ErrorAction SilentlyContinue).ORACLE_HOME
                        }
                        catch
                        {
                            # このキーにはORACLE_HOMEがない
                            continue
                        }

                        if (-not [string]::IsNullOrEmpty($oracleHome))
                        {
                            # tnsnames.oraの標準的な場所を確認
                            $possiblePaths = @(
                                (Join-Path $oracleHome "network\admin\tnsnames.ora"),
                                (Join-Path $oracleHome "NETWORK\ADMIN\tnsnames.ora"),
                                (Join-Path $oracleHome "admin\network\tnsnames.ora"),
                                (Join-Path $oracleHome "ADMIN\NETWORK\tnsnames.ora")
                            )

                            foreach ($tnsPath in $possiblePaths)
                            {
                                if (Test-Path $tnsPath)
                                {
                                    return $tnsPath
                                }
                            }
                        }
                    }
                }
            }

            # TNS_ADMIN環境変数もチェック
            if (-not [string]::IsNullOrEmpty($env:TNS_ADMIN))
            {
                $tnsPath = Join-Path $env:TNS_ADMIN "tnsnames.ora"
                if (Test-Path $tnsPath)
                {
                    return $tnsPath
                }
            }

            return ""
        }
        catch
        {
            Write-Host "レジストリの検索中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
            return ""
        }
    }

    # ファイルシステムからtnsnames.oraを検索（プライベートヘルパーメソッド）
    hidden [string[]] FindTnsNamesFiles()
    {
        try
        {
            $foundFiles = @()

            # 検索対象のルートパス
            $searchPaths = @(
                "C:\oracle",
                "C:\app",
                "D:\oracle",
                "D:\app",
                $env:ProgramFiles,
                ${env:ProgramFiles(x86)}
            )

            foreach ($rootPath in $searchPaths)
            {
                if (Test-Path $rootPath)
                {
                    Write-Host "検索中: $rootPath" -ForegroundColor Gray

                    # tnsnames.oraファイルを再帰的に検索
                    $files = Get-ChildItem -Path $rootPath -Filter "tnsnames.ora" -Recurse -ErrorAction SilentlyContinue |
                             Where-Object { $_.PSIsContainer -eq $false }

                    foreach ($file in $files)
                    {
                        # 既に見つかったファイルと重複していないか確認
                        if ($foundFiles -notcontains $file.FullName)
                        {
                            $foundFiles += $file.FullName
                        }
                    }
                }
            }

            # インスタントクライアントの一般的な場所もチェック
            $instantClientPaths = @(
                "C:\instantclient*",
                "D:\instantclient*",
                "$env:USERPROFILE\instantclient*"
            )

            foreach ($pattern in $instantClientPaths)
            {
                $paths = Get-Item $pattern -ErrorAction SilentlyContinue
                foreach ($path in $paths)
                {
                    $tnsPath = Join-Path $path.FullName "network\admin\tnsnames.ora"
                    if ((Test-Path $tnsPath) -and ($foundFiles -notcontains $tnsPath))
                    {
                        $foundFiles += $tnsPath
                    }
                }
            }

            return $foundFiles
        }
        catch
        {
            Write-Host "ファイルシステムの検索中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
            return @()
        }
    }

    # 複数のtnsnames.oraファイルから選択（プライベートヘルパーメソッド）
    hidden [string] SelectTnsNamesFile([string[]]$files)
    {
        try
        {
            Write-Host "`n見つかった tnsnames.ora ファイル:" -ForegroundColor Cyan
            Write-Host "=================================" -ForegroundColor Cyan

            for ($i = 0; $i -lt $files.Count; $i++)
            {
                # ファイル情報を取得
                $fileInfo = Get-Item $files[$i]
                $lastModified = $fileInfo.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss")
                $fileSize = "{0:N0}" -f ($fileInfo.Length / 1KB) + " KB"

                Write-Host "$($i + 1). $($files[$i])" -ForegroundColor White
                Write-Host "   最終更新: $lastModified | サイズ: $fileSize" -ForegroundColor Gray
            }

            Write-Host "=================================" -ForegroundColor Cyan

            # ユーザーに選択を促す
            $isValid = $false
            $index = -1
            do
            {
                $selection = Read-Host "`n使用するファイルの番号を入力してください (1-$($files.Count))"

                # 入力値の検証
                $isValid = $false
                if ($selection -match '^\d+$')
                {
                    $index = [int]$selection - 1
                    if ($index -ge 0 -and $index -lt $files.Count)
                    {
                        $isValid = $true
                    }
                }

                if (-not $isValid)
                {
                    Write-Host "無効な入力です。1から$($files.Count)までの数字を入力してください。" -ForegroundColor Red
                }
            }
            while (-not $isValid)

            $selectedFile = $files[$index]
            Write-Host "`n選択されたファイル: $selectedFile" -ForegroundColor Green

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("tnsnames.oraファイルを選択しました: $selectedFile", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] tnsnames.oraファイルを選択しました: $selectedFile" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $selectedFile
        }
        catch
        {
            Write-Host "ファイル選択中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            # エラーの場合は最初のファイルを返す
            return $files[0]
        }
    }

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

            $result = $this.ExecuteSelect($sql, $parameters)

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("テーブル一覧を取得しました。テーブル数: $($result.Rows.Count)", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] テーブル一覧を取得しました。テーブル数: $($result.Rows.Count)" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $result
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0080", "テーブル一覧取得エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

            $result = $this.ExecuteSelect($sql, $parameters)

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("テーブル構造を取得しました: $tableName", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] テーブル構造を取得しました: $tableName" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $result
        }
        catch
        {
            $this.last_error_message = $_.Exception.Message
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0081", "テーブル構造取徖エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            if ($this.Connection)
            {
                if ($this.Connection.State -eq 'Open')
                {
                    $this.Connection.Close()
                }
                $this.Connection.Dispose()
                $this.Connection = $null
            }

            if ($this.Transaction)
            {
                $this.Transaction.Dispose()
                $this.Transaction = $null
            }

            $this.is_connected = $false
            $this.is_transaction_active = $false

            Write-Host "初期化失敗時のクリーンアップが完了しました。" -ForegroundColor Yellow

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("初期化失敗時のクリーンアップを実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 初期化失敗時のクリーンアップを実行しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            Write-Host "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # デストラクタ
    [void] Dispose()
    {
        try
        {
            $this.Disconnect()

            # 正常ログ出力
            if ($global:Common)
            {
                $global:Common.WriteLog("OracleDriverを破棄しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] OracleDriverを破棄しました" | Out-File -Append -FilePath ([OracleDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("OracleDriverError_0090", "OracleDriver破棄エラー: $($_.Exception.Message)", "OracleDriver", [OracleDriver]::ErrorLogFile)
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