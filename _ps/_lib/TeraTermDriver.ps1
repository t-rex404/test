# TeraTerm操作クラス
# TeraTermのマクロ機能を使用してSSH/Telnet接続を自動化

class TeraTermDriver
{
    [string]$TeraTermPath
    [string]$MacroPath
    [string]$LogPath
    [bool]$is_connected
    [bool]$is_initialized
    [string]$last_error_message
    [hashtable]$connection_parameters
    [System.Diagnostics.Process]$teraterm_process
    [string]$current_host
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
            $this.current_host = ""
            $this.current_protocol = "SSH"
            $this.current_port = 22
            
            # TeraTermのパスを自動検出
            $this.TeraTermPath = $this.FindTeraTermPath()
            
            # 一時ディレクトリを作成
            $this.CreateTempDirectories()
            
            Write-Host "TeraTermDriverの初期化が完了しました。"
            $this.is_initialized = $true
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
                    $global:Common.HandleError("TeraTermError_0001", "TeraTermDriver初期化エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TeraTermDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "TeraTermDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # TeraTermの実行ファイルパスを検出
    [string] FindTeraTermPath()
    {
        try
        {
            # 一般的なTeraTermのインストールパスをチェック
            $possiblePaths = @(
                "C:\Program Files (x86)\teraterm\ttermpro.exe",
                "C:\Program Files\teraterm\ttermpro.exe",
                "C:\teraterm\ttermpro.exe",
                "ttermpro.exe"  # PATH環境変数から検索
            )

            foreach ($path in $possiblePaths)
            {
                if ($path -eq "ttermpro.exe")
                {
                    # PATH環境変数から検索
                    $foundPath = Get-Command "ttermpro.exe" -ErrorAction SilentlyContinue
                    if ($foundPath)
                    {
                        Write-Host "TeraTermのパスを検出しました: $($foundPath.Source)"
                        return $foundPath.Source
                    }
                }
                else
                {
                    if (Test-Path $path)
                    {
                        Write-Host "TeraTermのパスを検出しました: $path"
                        return $path
                    }
                }
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
                    $global:Common.HandleError("TeraTermError_0002", "TeraTerm実行ファイルパス取得エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TeraTermの実行ファイルパス取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "TeraTermの実行ファイルパス取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 一時ディレクトリを作成
    [void] CreateTempDirectories()
    {
        try
        {
            $tempDir = Join-Path $env:TEMP "TeraTermDriver"
            if (-not (Test-Path $tempDir))
            {
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            }

            $this.MacroPath = Join-Path $tempDir "macros"
            $this.LogPath = Join-Path $tempDir "logs"

            if (-not (Test-Path $this.MacroPath))
            {
                New-Item -ItemType Directory -Path $this.MacroPath -Force | Out-Null
            }

            if (-not (Test-Path $this.LogPath))
            {
                New-Item -ItemType Directory -Path $this.LogPath -Force | Out-Null
            }

            Write-Host "一時ディレクトリを作成しました: $tempDir"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0003", "一時ディレクトリ作成エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
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

    # 接続パラメータを設定
    [void] SetConnectionParameters([string]$host, [string]$port = "22", [string]$protocol = "SSH", [string]$username = "", [string]$password = "")
    {
        try
        {
            $this.connection_parameters = @{
                Host = $host
                Port = $port
                Protocol = $protocol
                Username = $username
                Password = $password
            }

            $this.current_host = $host
            $this.current_port = [int]$port
            $this.current_protocol = $protocol

            Write-Host "接続パラメータを設定しました: $host:$port ($protocol)"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0004", "接続パラメータ設定エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
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

            # マクロファイルを作成
            $macroFile = $this.CreateConnectionMacro()
            
            # TeraTermを起動してマクロを実行
            $this.StartTeraTermWithMacro($macroFile)

            $this.is_connected = $true
            Write-Host "TeraTermでサーバーに接続しました: $($this.current_host):$($this.current_port)"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0005", "サーバー接続エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "サーバーへの接続に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "サーバーへの接続に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続用マクロファイルを作成
    [string] CreateConnectionMacro()
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macroFileName = "connect_$timestamp.ttl"
            $macroFilePath = Join-Path $this.MacroPath $macroFileName

            $host = $this.connection_parameters.Host
            $port = $this.connection_parameters.Port
            $protocol = $this.connection_parameters.Protocol
            $username = $this.connection_parameters.Username
            $password = $this.connection_parameters.Password

            # ログファイル名
            $logFileName = "teraterm_$timestamp.log"
            $logFilePath = Join-Path $this.LogPath $logFileName

            # マクロ内容を作成
            $macroContent = @"
; TeraTerm接続マクロ
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; ログファイルを開く
logopen '$logFilePath'

; 接続先ホストに接続
connect '$host' /$protocol /$port

; 接続待機
wait 'login:'
if result = 1 then
    ; ユーザー名を入力
    sendln '$username'
    wait 'Password:'
    if result = 1 then
        ; パスワードを入力
        sendln '$password'
        wait '$'
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

            # マクロファイルに書き込み
            $macroContent | Out-File -FilePath $macroFilePath -Encoding UTF8

            Write-Host "接続マクロファイルを作成しました: $macroFilePath"
            return $macroFilePath
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0006", "接続マクロ作成エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "接続マクロの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "接続マクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # TeraTermを起動してマクロを実行
    [void] StartTeraTermWithMacro([string]$macroFilePath)
    {
        try
        {
            # TeraTermプロセスを起動
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $this.TeraTermPath
            $processInfo.Arguments = "/M=`"$macroFilePath`""
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $false

            $this.teraterm_process = [System.Diagnostics.Process]::Start($processInfo)

            if (-not $this.teraterm_process)
            {
                throw "TeraTermプロセスの起動に失敗しました。"
            }

            Write-Host "TeraTermプロセスを起動しました (PID: $($this.teraterm_process.Id))"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0007", "TeraTerm起動エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TeraTermの起動に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "TeraTermの起動に失敗しました: $($_.Exception.Message)"
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
            $macroFile = $this.CreateDisconnectMacro($confirm)
            
            # マクロを実行
            $this.ExecuteMacro($macroFile)

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
                    $global:Common.HandleError("TeraTermError_0008", "接続切断エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
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

    # 切断用マクロファイルを作成
    [string] CreateDisconnectMacro([bool]$confirm)
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macroFileName = "disconnect_$timestamp.ttl"
            $macroFilePath = Join-Path $this.MacroPath $macroFileName

            $confirmFlag = if ($confirm) { "1" } else { "0" }

            # マクロ内容を作成
            $macroContent = @"
; TeraTerm切断マクロ
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; 接続を切断
disconnect $confirmFlag

; ログファイルを閉じる
logclose

; マクロ終了
"@

            # マクロファイルに書き込み
            $macroContent | Out-File -FilePath $macroFilePath -Encoding UTF8

            Write-Host "切断マクロファイルを作成しました: $macroFilePath"
            return $macroFilePath
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0009", "切断マクロ作成エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "切断マクロの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "切断マクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # マクロを実行
    [void] ExecuteMacro([string]$macroFilePath)
    {
        try
        {
            if (-not (Test-Path $macroFilePath))
            {
                throw "マクロファイルが見つかりません: $macroFilePath"
            }

            # TeraTermプロセスが実行中かチェック
            if ($this.teraterm_process -and -not $this.teraterm_process.HasExited)
            {
                # 既存のプロセスでマクロを実行
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $this.TeraTermPath
                $processInfo.Arguments = "/M=`"$macroFilePath`""
                $processInfo.UseShellExecute = $false
                $processInfo.CreateNoWindow = $false

                $macroProcess = [System.Diagnostics.Process]::Start($processInfo)
                if (-not $macroProcess)
                {
                    throw "マクロの実行に失敗しました。"
                }

                Write-Host "マクロを実行しました: $macroFilePath"
            }
            else
            {
                throw "TeraTermプロセスが実行されていません。"
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0010", "マクロ実行エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "マクロの実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "マクロの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # コマンド実行関連
    # ========================================

    # コマンドを実行
    [string] ExecuteCommand([string]$command, [string]$expectedPrompt = "$", [int]$timeoutSeconds = 30)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "接続されていません。先にConnect()を実行してください。"
            }

            # コマンド実行用マクロファイルを作成
            $macroFile = $this.CreateCommandMacro($command, $expectedPrompt, $timeoutSeconds)
            
            # マクロを実行
            $this.ExecuteMacro($macroFile)

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
                    $global:Common.HandleError("TeraTermError_0011", "コマンド実行エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "コマンドの実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "コマンドの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # コマンド実行用マクロファイルを作成
    [string] CreateCommandMacro([string]$command, [string]$expectedPrompt, [int]$timeoutSeconds)
    {
        try
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $macroFileName = "command_$timestamp.ttl"
            $macroFilePath = Join-Path $this.MacroPath $macroFileName

            # マクロ内容を作成
            $macroContent = @"
; TeraTermコマンド実行マクロ
; 作成日時: $(Get-Date -Format "yyyy/MM/dd HH:mm:ss")

; コマンドを送信
sendln '$command'

; プロンプトを待機
wait '$expectedPrompt' $timeoutSeconds

; マクロ終了
"@

            # マクロファイルに書き込み
            $macroContent | Out-File -FilePath $macroFilePath -Encoding UTF8

            Write-Host "コマンド実行マクロファイルを作成しました: $macroFilePath"
            return $macroFilePath
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0012", "コマンド実行マクロ作成エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "コマンド実行マクロの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "コマンド実行マクロの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # 最後のコマンド結果を取得
    [string] GetLastCommandResult()
    {
        try
        {
            # 最新のログファイルを取得
            $logFiles = Get-ChildItem -Path $this.LogPath -Filter "*.log" | Sort-Object LastWriteTime -Descending
            if ($logFiles.Count -eq 0)
            {
                return ""
            }

            $latestLogFile = $logFiles[0].FullName
            $logContent = Get-Content -Path $latestLogFile -Raw -Encoding UTF8

            # 最後の数行を取得（プロンプト以降の内容）
            $lines = $logContent -split "`n"
            $resultLines = @()
            $foundPrompt = $false

            foreach ($line in $lines)
            {
                if ($foundPrompt)
                {
                    $resultLines += $line
                }
                elseif ($line -match '\$' -or $line -match '#' -or $line -match '>')
                {
                    $foundPrompt = $true
                }
            }

            return ($resultLines -join "`n").Trim()
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0013", "コマンド結果取得エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "コマンド結果の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            return ""
        }
    }

    # ========================================
    # ファイル転送関連
    # ========================================

    # ファイルをアップロード（SCP使用）
    [void] UploadFile([string]$localFilePath, [string]$remoteFilePath)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "接続されていません。先にConnect()を実行してください。"
            }

            if (-not (Test-Path $localFilePath))
            {
                throw "ローカルファイルが見つかりません: $localFilePath"
            }

            # SCPコマンドを実行
            $scpCommand = "scp `"$localFilePath`" $($this.connection_parameters.Username)@$($this.connection_parameters.Host):`"$remoteFilePath`""
            $result = $this.ExecuteCommand($scpCommand, "$", 60)

            Write-Host "ファイルをアップロードしました: $localFilePath -> $remoteFilePath"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0014", "ファイルアップロードエラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ファイルのアップロードに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ファイルのアップロードに失敗しました: $($_.Exception.Message)"
        }
    }

    # ファイルをダウンロード（SCP使用）
    [void] DownloadFile([string]$remoteFilePath, [string]$localFilePath)
    {
        try
        {
            if (-not $this.is_connected)
            {
                throw "接続されていません。先にConnect()を実行してください。"
            }

            # ローカルディレクトリが存在しない場合は作成
            $localDir = Split-Path $localFilePath -Parent
            if (-not (Test-Path $localDir))
            {
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
            }

            # SCPコマンドを実行
            $scpCommand = "scp $($this.connection_parameters.Username)@$($this.connection_parameters.Host):`"$remoteFilePath`" `"$localFilePath`""
            $result = $this.ExecuteCommand($scpCommand, "$", 60)

            Write-Host "ファイルをダウンロードしました: $remoteFilePath -> $localFilePath"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("TeraTermError_0015", "ファイルダウンロードエラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ファイルのダウンロードに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ファイルのダウンロードに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ユーティリティ関連
    # ========================================

    # 接続状態を確認
    [bool] IsConnected()
    {
        return $this.is_connected
    }

    # 接続情報を取得
    [hashtable] GetConnectionInfo()
    {
        return @{
            Host = $this.current_host
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
            $logFiles = Get-ChildItem -Path $this.LogPath -Filter "*.log" | Sort-Object LastWriteTime -Descending
            return $logFiles
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
            $macroFiles = Get-ChildItem -Path $this.MacroPath -Filter "*.ttl" | Sort-Object LastWriteTime -Descending
            return $macroFiles
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
            $cutoffDate = (Get-Date).AddDays(-7)
            $oldMacros = Get-ChildItem -Path $this.MacroPath -Filter "*.ttl" | Where-Object { $_.LastWriteTime -lt $cutoffDate }
            foreach ($macro in $oldMacros)
            {
                Remove-Item -Path $macro.FullName -Force
            }

            # 古いログファイルを削除（7日以上前）
            $oldLogs = Get-ChildItem -Path $this.LogPath -Filter "*.log" | Where-Object { $_.LastWriteTime -lt $cutoffDate }
            foreach ($log in $oldLogs)
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
    # 破棄・クリーンアップ関連
    # ========================================

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
            if (-not [string]::IsNullOrEmpty($this.MacroPath) -and (Test-Path $this.MacroPath))
            {
                try
                {
                    Remove-Item -Path $this.MacroPath -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch
                {
                    Write-Host "マクロディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            if (-not [string]::IsNullOrEmpty($this.LogPath) -and (Test-Path $this.LogPath))
            {
                try
                {
                    Remove-Item -Path $this.LogPath -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch
                {
                    Write-Host "ログディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            Write-Host "初期化失敗時のクリーンアップが完了しました。" -ForegroundColor Yellow
        }
        catch
        {
            Write-Host "初期化失敗時のクリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

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
                    $global:Common.HandleError("TeraTermError_0016", "TeraTermDriver破棄エラー: $($_.Exception.Message)", "TeraTermDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TeraTermDriverの破棄に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

Write-Host "TeraTermDriverクラスが正常にインポートされました。" -ForegroundColor Green

