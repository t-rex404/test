# TeraTerm操作クラス
# TeraTermのマクロ機能を使用してSSH/Telnet接続を自動化

class TeraTermDriver
{
    [string]$teraterm_exe_path
    [string]$macro_directory
    [string]$log_directory
    [string]$temp_directory
    [bool]$is_connected
    [bool]$is_initialized
    [string]$last_error_message
    [hashtable]$connection_parameters
    [System.Diagnostics.Process]$teraterm_process
    [string]$current_hostname
    [string]$current_protocol
    [int]$current_port

    # ========================================
    # 初期化・接続関連
    # ========================================

    TeraTermDriver()
    {
        try
        {
            $this.is_connected = $false
            $this.is_initialized = $false
            $this.last_error_message = ""
            $this.connection_parameters = @{}
            $this.teraterm_process = $null
            $this.current_hostname = ""
            $this.current_protocol = "SSH"
            $this.current_port = 22

            # TeraTermのパスを自動検出
            $this.teraterm_exe_path = $this.GetTeraTermExecutablePath()

            # 一時ディレクトリを作成
            $this.temp_directory = $this.CreateTempDirectory()

            # マクロとログディレクトリを設定
            $this.macro_directory = Join-Path $this.temp_directory "macros"
            $this.log_directory = Join-Path $this.temp_directory "logs"
            $this.CreateWorkingDirectories()

            $this.is_initialized = $true
            Write-Host "TeraTermDriverの初期化が完了しました。" -ForegroundColor Green
        }
        catch
        {
            # 初期化に失敗した場合のクリーンアップ
            Write-Host "TeraTermDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()

            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0001", "TeraTermDriver初期化エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TeraTermDriver初期化エラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "TeraTermDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # TeraTermの実行ファイルパスを取得
    [string] GetTeraTermExecutablePath()
    {
        try
        {
            # 一般的なTeraTermのインストールパスをチェック
            $possible_paths = @(
                "C:\Program Files (x86)\teraterm\ttermpro.exe",
                "C:\Program Files\teraterm\ttermpro.exe",
                "${env:ProgramFiles(x86)}\teraterm\ttermpro.exe",
                "${env:ProgramFiles}\teraterm\ttermpro.exe",
                "C:\Program Files (x86)\teraterm5\ttermpro.exe",
                "C:\Program Files\teraterm5\ttermpro.exe",
                "${env:ProgramFiles(x86)}\teraterm5\ttermpro.exe",
                "${env:ProgramFiles}\teraterm5\ttermpro.exe"
            )

            foreach ($path in $possible_paths)
            {
                if (Test-Path $path)
                {
                    Write-Host "TeraTermのパスを検出しました: $path"
                    return $path
                }
            }

            # PATH環境変数から検索
            $command = Get-Command "ttermpro.exe" -ErrorAction SilentlyContinue
            if ($command)
            {
                Write-Host "TeraTermのパスを検出しました: $($command.Source)"
                return $command.Source
            }

            throw "TeraTermの実行ファイルが見つかりません。TeraTermがインストールされていることを確認してください。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0002", "TeraTerm実行ファイルパス取得エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TeraTerm実行ファイルパス取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "TeraTermの実行ファイルパス取得に失敗しました: $($_.Exception.Message)"
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
            $temp_dir = Join-Path $base_dir "TeraTermDriver"
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
                    $global:Common.HandleError("TeraTermDriverError_0003", "一時ディレクトリ作成エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
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

    # 作業ディレクトリを作成
    [void] CreateWorkingDirectories()
    {
        try
        {
            if (-not (Test-Path $this.macro_directory))
            {
                New-Item -ItemType Directory -Path $this.macro_directory -Force | Out-Null
            }

            if (-not (Test-Path $this.log_directory))
            {
                New-Item -ItemType Directory -Path $this.log_directory -Force | Out-Null
            }

            Write-Host "作業ディレクトリを作成しました。"
        }
        catch
        {
            throw "作業ディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            Write-Host "初期化失敗時のクリーンアップを実行します。" -ForegroundColor Yellow

            # プロセスが起動している場合は終了
            if ($this.teraterm_process -and -not $this.teraterm_process.HasExited)
            {
                try
                {
                    $this.teraterm_process.Kill()
                    Write-Host "TeraTermプロセスを終了しました。" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "TeraTermプロセスの終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # 一時ディレクトリをクリーンアップ
            if (-not [string]::IsNullOrEmpty($this.temp_directory) -and (Test-Path $this.temp_directory))
            {
                try
                {
                    Remove-Item -Path $this.temp_directory -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "一時ディレクトリを削除しました。" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "一時ディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            Write-Host "初期化失敗時のクリーンアップが完了しました。" -ForegroundColor Yellow
        }
        catch
        {
            Write-Host "初期化失敗時のクリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # ========================================
    # 接続管理
    # ========================================

    # 接続パラメータを設定（パスワード認証）
    [void] SetConnectionParameters([string]$hostname, [string]$port = "22", [string]$protocol = "SSH", [string]$username = "", [string]$password = "")
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "TeraTermDriverが初期化されていません。"
            }

            $this.connection_parameters = @{
                Host = $hostname
                Port = $port
                Protocol = $protocol
                Username = $username
                Password = $password
                PemFilePath = ""  # PEMファイルパス（オプション）
            }

            $this.current_hostname = $hostname
            $this.current_port = [int]$port
            $this.current_protocol = $protocol

            Write-Host "接続パラメータを設定しました: $hostname`:$port ($protocol)"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0010", "接続パラメータ設定エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "接続パラメータ設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "接続パラメータの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続パラメータを設定（PEM認証）
    [void] SetConnectionParametersWithPem([string]$hostname, [string]$port, [string]$username, [string]$pem_file_path)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "TeraTermDriverが初期化されていません。"
            }

            if (-not (Test-Path $pem_file_path))
            {
                throw "PEMファイルが見つかりません: $pem_file_path"
            }

            $this.connection_parameters = @{
                Host = $hostname
                Port = $port
                Protocol = "SSH"
                Username = $username
                Password = ""
                PemFilePath = $pem_file_path
            }

            $this.current_hostname = $hostname
            $this.current_port = [int]$port
            $this.current_protocol = "SSH"

            Write-Host "接続パラメータを設定しました（PEM認証）: $hostname`:$port"
            Write-Host "  PEMファイル: $pem_file_path"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0013", "PEM接続パラメータ設定エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "PEM接続パラメータ設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "PEM接続パラメータの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # サーバーに接続
    [void] Connect()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "TeraTermDriverが初期化されていません。"
            }

            if ($this.is_connected)
            {
                Write-Host "既に接続されています。"
                return
            }

            if ($this.connection_parameters.Count -eq 0)
            {
                throw "接続パラメータが設定されていません。SetConnectionParameters()を先に実行してください。"
            }

            # 接続用マクロファイルを作成
            $macro_file = $this.CreateConnectionMacro()

            # TeraTermを起動してマクロを実行
            $this.StartTeraTermWithMacro($macro_file)

            $this.is_connected = $true
            Write-Host "TeraTermでサーバーに接続しました: $($this.current_hostname):$($this.current_port)" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0011", "サーバー接続エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "サーバー接続エラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "サーバーへの接続に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続を切断
    [void] Disconnect([bool]$confirm = $false)
    {
        try
        {
            if (-not $this.is_connected)
            {
                Write-Host "接続されていません。"
                return
            }

            # 切断用マクロファイルを作成
            $macro_file = $this.CreateDisconnectMacro($confirm)

            # マクロを実行
            $this.ExecuteMacro($macro_file)

            $this.is_connected = $false
            Write-Host "TeraTermの接続を切断しました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0012", "接続切断エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "接続切断エラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "接続の切断に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続状態を確認（プロセス状態も含めて実際の接続を検証）
    [bool] IsConnected()
    {
        # 接続フラグがfalseの場合は即座にfalseを返す
        if (-not $this.is_connected)
        {
            return $false
        }

        # TeraTermプロセスが実行中かを確認
        if ($null -eq $this.teraterm_process -or $this.teraterm_process.HasExited)
        {
            $this.is_connected = $false
            return $false
        }

        # プロセスが応答しているかを確認
        try
        {
            $this.teraterm_process.Refresh()
            if ($this.teraterm_process.Responding)
            {
                return $true
            }
            else
            {
                Write-Host "TeraTermプロセスが応答していません。" -ForegroundColor Yellow
                $this.is_connected = $false
                return $false
            }
        }
        catch
        {
            $this.is_connected = $false
            return $false
        }
    }

    # ========================================
    # マクロ作成・実行
    # ========================================

    # 接続用マクロファイルを作成
    [string] CreateConnectionMacro()
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macro_file_name = "connect_$timestamp.ttl"
            $macro_file_path = Join-Path $this.macro_directory $macro_file_name

            $hostname = $this.connection_parameters.Host
            $port = $this.connection_parameters.Port
            $protocol = $this.connection_parameters.Protocol
            $username = $this.connection_parameters.Username
            $password = $this.connection_parameters.Password

            # ログファイル名
            $log_file_name = "teraterm_$timestamp.log"
            $log_file_path = Join-Path $this.log_directory $log_file_name

            # PEM認証かパスワード認証かを判定
            $is_pem_auth = -not [string]::IsNullOrEmpty($this.connection_parameters.PemFilePath)

            # マクロ内容を作成
            if ($is_pem_auth)
            {
                # PEM認証用マクロ
                $pem_file = $this.connection_parameters.PemFilePath.Replace('\', '/')
                $macro_content = @"
; TeraTerm接続マクロ（PEM認証）
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; ログファイルを開く
logopen '$log_file_path'

; SSH鍵認証で接続
connect '$hostname`:$port /ssh /2 /auth=publickey /user=$username /keyfile=$pem_file'

; 接続確認
pause 3
sendln ''
wait '$' 10
if result = 1 then
    ; 接続成功
    messagebox '接続が完了しました（PEM認証）' 'TeraTermDriver'
else
    wait '#' 10
    if result = 1 then
        ; rootプロンプトで接続成功
        messagebox '接続が完了しました（PEM認証）' 'TeraTermDriver'
    else
        messagebox 'プロンプトが表示されませんでした' 'TeraTermDriver'
    endif
endif

; マクロ終了
"@
            }
            else
            {
                # パスワード認証用マクロ
                $macro_content = @"
; TeraTerm接続マクロ（パスワード認証）
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; ログファイルを開く
logopen '$log_file_path'

; 接続先ホストに接続
connect '$hostname' /$protocol /$port

; 接続待機
wait 'login:' 30
if result = 1 then
    ; ユーザー名を入力
    sendln '$username'
    wait 'Password:' 10
    if result = 1 then
        ; パスワードを入力
        sendln '$password'
        wait '$' 10
        if result = 1 then
            ; 接続成功
            messagebox '接続が完了しました' 'TeraTermDriver'
        else
            messagebox 'パスワード認証に失敗しました' 'TeraTermDriver'
        endif
    else
        messagebox 'パスワードプロンプトが表示されませんでした' 'TeraTermDriver'
    endif
else
    messagebox 'ログインプロンプトが表示されませんでした' 'TeraTermDriver'
endif

; マクロ終了
"@
            }

            # マクロファイルに書き込み
            $macro_content | Out-File -FilePath $macro_file_path -Encoding Default

            Write-Host "接続マクロファイルを作成しました: $macro_file_path"
            return $macro_file_path
        }
        catch
        {
            throw "接続マクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # 切断用マクロファイルを作成
    [string] CreateDisconnectMacro([bool]$confirm)
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macro_file_name = "disconnect_$timestamp.ttl"
            $macro_file_path = Join-Path $this.macro_directory $macro_file_name

            $confirm_flag = if ($confirm) { "1" } else { "0" }

            # マクロ内容を作成
            $macro_content = @"
; TeraTerm切断マクロ
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; 接続を切断
disconnect $confirm_flag

; ログファイルを閉じる
logclose

; マクロ終了
"@

            # マクロファイルに書き込み
            $macro_content | Out-File -FilePath $macro_file_path -Encoding Default

            Write-Host "切断マクロファイルを作成しました: $macro_file_path"
            return $macro_file_path
        }
        catch
        {
            throw "切断マクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # TeraTermを起動してマクロを実行
    [void] StartTeraTermWithMacro([string]$macro_file_path)
    {
        try
        {
            # TeraTermプロセスを起動
            $process_info = New-Object System.Diagnostics.ProcessStartInfo
            $process_info.FileName = $this.teraterm_exe_path
            $process_info.Arguments = "/M=`"$macro_file_path`""
            $process_info.UseShellExecute = $false
            $process_info.CreateNoWindow = $false

            $this.teraterm_process = [System.Diagnostics.Process]::Start($process_info)

            if (-not $this.teraterm_process)
            {
                throw "TeraTermプロセスの起動に失敗しました。"
            }

            Write-Host "TeraTermプロセスを起動しました (PID: $($this.teraterm_process.Id))"
        }
        catch
        {
            throw "TeraTermの起動に失敗しました: $($_.Exception.Message)"
        }
    }

    # マクロを実行
    [void] ExecuteMacro([string]$macro_file_path)
    {
        try
        {
            if (-not (Test-Path $macro_file_path))
            {
                throw "マクロファイルが見つかりません: $macro_file_path"
            }

            # TeraTermプロセスが実行中かチェック
            if ($this.teraterm_process -and -not $this.teraterm_process.HasExited)
            {
                # 新しいプロセスでマクロを実行
                $process_info = New-Object System.Diagnostics.ProcessStartInfo
                $process_info.FileName = $this.teraterm_exe_path
                $process_info.Arguments = "/M=`"$macro_file_path`""
                $process_info.UseShellExecute = $false
                $process_info.CreateNoWindow = $false

                $macro_process = [System.Diagnostics.Process]::Start($process_info)
                if (-not $macro_process)
                {
                    throw "マクロの実行に失敗しました。"
                }

                Write-Host "マクロを実行しました: $macro_file_path"
            }
            else
            {
                throw "TeraTermプロセスが実行されていません。"
            }
        }
        catch
        {
            throw "マクロの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # コマンド実行
    # ========================================

    # コマンドを実行
    [string] ExecuteCommand([string]$command, [string]$expected_prompt = "$", [int]$timeout_seconds = 30)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "接続されていません。先にConnect()を実行してください。"
            }

            # コマンド実行用マクロファイルを作成
            $macro_file = $this.CreateCommandMacro($command, $expected_prompt, $timeout_seconds)

            # マクロを実行
            $this.ExecuteMacro($macro_file)

            # ログファイルから結果を取得
            $result = $this.GetLastCommandResult()

            Write-Host "コマンドを実行しました: $command"
            return $result
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0020", "コマンド実行エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "コマンド実行エラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "コマンドの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # コマンド実行用マクロファイルを作成
    [string] CreateCommandMacro([string]$command, [string]$expected_prompt, [int]$timeout_seconds)
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macro_file_name = "command_$timestamp.ttl"
            $macro_file_path = Join-Path $this.macro_directory $macro_file_name

            # マクロ内容を作成
            $macro_content = @"
; TeraTermコマンド実行マクロ
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; コマンドを送信
sendln '$command'

; プロンプトを待機
wait '$expected_prompt' $timeout_seconds

; マクロ終了
"@

            # マクロファイルに書き込み
            $macro_content | Out-File -FilePath $macro_file_path -Encoding Default

            Write-Host "コマンド実行マクロファイルを作成しました: $macro_file_path"
            return $macro_file_path
        }
        catch
        {
            throw "コマンド実行マクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # 最後のコマンド結果を取得
    [string] GetLastCommandResult()
    {
        try
        {
            # 最新のログファイルを取得
            $log_files = Get-ChildItem -Path $this.log_directory -Filter "*.log" | Sort-Object LastWriteTime -Descending
            if ($log_files.Count -eq 0)
            {
                return ""
            }

            $latest_log_file = $log_files[0].FullName
            $log_content = Get-Content -Path $latest_log_file -Raw -Encoding Default

            # 最後のコマンドとその結果のみを取得
            $lines = $log_content -split "`n"
            $result_lines = @()
            $last_prompt_index = -1

            # 最後のプロンプトを見つける
            for ($i = $lines.Count - 1; $i -ge 0; $i--)
            {
                if ($lines[$i] -match '\$$' -or $lines[$i] -match '#$' -or $lines[$i] -match '>$')
                {
                    $last_prompt_index = $i
                    break
                }
            }

            # 最後のプロンプトの次の行から最後までを取得
            if ($last_prompt_index -ge 0 -and $last_prompt_index + 1 -lt $lines.Count)
            {
                for ($i = $last_prompt_index + 1; $i -lt $lines.Count; $i++)
                {
                    # 次のプロンプトが見つかったら終了
                    if ($lines[$i] -match '\$$' -or $lines[$i] -match '#$' -or $lines[$i] -match '>$')
                    {
                        break
                    }
                    $result_lines += $lines[$i]
                }
            }

            return ($result_lines -join "`n").Trim()
        }
        catch
        {
            Write-Host "コマンド結果の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            return ""
        }
    }

    # ========================================
    # ファイル転送
    # ========================================

    # ファイルをアップロード（ZMODEMプロトコル使用）
    [void] UploadFile([string]$local_file_path, [string]$remote_file_path)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "接続されていません。先にConnect()を実行してください。"
            }

            if (-not (Test-Path $local_file_path))
            {
                throw "ローカルファイルが見つかりません: $local_file_path"
            }

            # ZMODEMプロトコルを使用したファイルアップロード用マクロを作成
            $macro_file = $this.CreateFileUploadMacro($local_file_path, $remote_file_path)

            # マクロを実行
            $this.ExecuteMacro($macro_file)

            Write-Host "ファイルをアップロードしました: $local_file_path -> $remote_file_path" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0030", "ファイルアップロードエラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ファイルアップロードエラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "ファイルのアップロードに失敗しました: $($_.Exception.Message)"
        }
    }

    # ファイルをダウンロード（ZMODEMプロトコル使用）
    [void] DownloadFile([string]$remote_file_path, [string]$local_file_path)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "接続されていません。先にConnect()を実行してください。"
            }

            # ローカルディレクトリが存在しない場合は作成
            $local_dir = Split-Path $local_file_path -Parent
            if (-not (Test-Path $local_dir))
            {
                New-Item -ItemType Directory -Path $local_dir -Force | Out-Null
            }

            # ZMODEMプロトコルを使用したファイルダウンロード用マクロを作成
            $macro_file = $this.CreateFileDownloadMacro($remote_file_path, $local_file_path)

            # マクロを実行
            $this.ExecuteMacro($macro_file)

            Write-Host "ファイルをダウンロードしました: $remote_file_path -> $local_file_path" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0031", "ファイルダウンロードエラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ファイルダウンロードエラー: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "ファイルのダウンロードに失敗しました: $($_.Exception.Message)"
        }
    }

    # ZMODEMファイルアップロード用マクロファイルを作成
    [string] CreateFileUploadMacro([string]$local_file_path, [string]$remote_file_path)
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macro_file_name = "upload_$timestamp.ttl"
            $macro_file_path = Join-Path $this.macro_directory $macro_file_name

            # ローカルファイルパスをTeraTermのパス形式に変換
            $tt_local_path = $local_file_path.Replace('\', '/')

            # マクロ内容を作成
            $macro_content = @"
; TeraTermファイルアップロードマクロ (ZMODEM)
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; ファイル送信準備
sendln 'cd $(Split-Path $remote_file_path -Parent)'
wait '$' 10

; ZMODEMファイル送信を開始
zmodemfile '$tt_local_path' 1

; 送信完了待機
wait '$' 30

; ファイル名変更（必要な場合）
sendln 'mv $(Split-Path $local_file_path -Leaf) $(Split-Path $remote_file_path -Leaf)'
wait '$' 5

; マクロ終了
"@

            # マクロファイルに書き込み
            $macro_content | Out-File -FilePath $macro_file_path -Encoding Default

            Write-Host "ファイルアップロードマクロファイルを作成しました: $macro_file_path"
            return $macro_file_path
        }
        catch
        {
            throw "アップロードマクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # ZMODEMファイルダウンロード用マクロファイルを作成
    [string] CreateFileDownloadMacro([string]$remote_file_path, [string]$local_file_path)
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macro_file_name = "download_$timestamp.ttl"
            $macro_file_path = Join-Path $this.macro_directory $macro_file_name

            # ローカルディレクトリをTeraTermのパス形式に変換
            $tt_local_dir = (Split-Path $local_file_path -Parent).Replace('\', '/')

            # マクロ内容を作成
            $macro_content = @"
; TeraTermファイルダウンロードマクロ (ZMODEM)
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; ダウンロード先ディレクトリを設定
zmodemrecv '$tt_local_dir'

; ファイル受信開始
sendln 'sz $remote_file_path'

; 受信完了待機
wait '$' 30

; マクロ終了
"@

            # マクロファイルに書き込み
            $macro_content | Out-File -FilePath $macro_file_path -Encoding Default

            Write-Host "ファイルダウンロードマクロファイルを作成しました: $macro_file_path"
            return $macro_file_path
        }
        catch
        {
            throw "ダウンロードマクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ユーティリティ
    # ========================================

    # 接続情報を取得
    [hashtable] GetConnectionInfo()
    {
        return @{
            Host = $this.current_hostname
            Port = $this.current_port
            Protocol = $this.current_protocol
            IsConnected = $this.is_connected
            ProcessId = if ($null -ne $this.teraterm_process) { $this.teraterm_process.Id } else { $null }
        }
    }

    # ログファイル一覧を取得
    [array] GetLogFiles()
    {
        try
        {
            $log_files = Get-ChildItem -Path $this.log_directory -Filter "*.log" | Sort-Object LastWriteTime -Descending
            return $log_files
        }
        catch
        {
            Write-Host "ログファイル一覧の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            return @()
        }
    }

    # マクロファイル一覧を取得
    [array] GetMacroFiles()
    {
        try
        {
            $macro_files = Get-ChildItem -Path $this.macro_directory -Filter "*.ttl" | Sort-Object LastWriteTime -Descending
            return $macro_files
        }
        catch
        {
            Write-Host "マクロファイル一覧の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            return @()
        }
    }

    # 一時ファイルをクリーンアップ
    [void] CleanupTempFiles()
    {
        try
        {
            # 古いマクロファイルを削除（7日以上前）
            $cutoff_date = (Get-Date).AddDays(-7)
            $old_macros = Get-ChildItem -Path $this.macro_directory -Filter "*.ttl" | Where-Object { $_.LastWriteTime -lt $cutoff_date }
            foreach ($macro in $old_macros)
            {
                Remove-Item -Path $macro.FullName -Force
            }

            # 古いログファイルを削除（7日以上前）
            $old_logs = Get-ChildItem -Path $this.log_directory -Filter "*.log" | Where-Object { $_.LastWriteTime -lt $cutoff_date }
            foreach ($log in $old_logs)
            {
                Remove-Item -Path $log.FullName -Force
            }

            Write-Host "一時ファイルのクリーンアップが完了しました。"
        }
        catch
        {
            Write-Host "一時ファイルのクリーンアップに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # ========================================
    # 破棄
    # ========================================

    # TeraTermDriverを破棄
    [void] Dispose()
    {
        try
        {
            Write-Host "TeraTermDriverの破棄を開始します。" -ForegroundColor Yellow

            # 接続を切断
            if ($this.is_connected)
            {
                try
                {
                    $this.Disconnect($false)
                }
                catch
                {
                    Write-Host "接続の切断中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # TeraTermプロセスを終了
            if ($this.teraterm_process -and -not $this.teraterm_process.HasExited)
            {
                try
                {
                    $this.teraterm_process.Kill()
                    $this.teraterm_process.WaitForExit(5000)
                    Write-Host "TeraTermプロセスを終了しました。" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "TeraTermプロセスの終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # 一時ファイルをクリーンアップ
            $this.CleanupTempFiles()

            # 状態をリセット
            $this.is_connected = $false
            $this.is_initialized = $false
            $this.teraterm_process = $null
            $this.connection_parameters = @{}

            Write-Host "TeraTermDriverの破棄が完了しました。" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermDriverError_0090", "TeraTermDriver破棄エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TeraTermDriver破棄エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

Write-Host "TeraTermDriverクラスが正常にインポートされました。" -ForegroundColor Green