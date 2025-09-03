# WinSCPDriver.ps1
# WinSCPを操作するためのPowerShellクラス
# ファイル転送、接続管理などの機能を提供

# 共通ライブラリをインポート
. "$PSScriptRoot\Common.ps1"

class WinSCPDriver
{
    # プライベートプロパティ
    [object]$Session
    [object]$SessionOptions
    [object]$TransferOptions
    [string]$WinSCPPath
    [bool]$IsConnected = $false
    [Common]$Common

    # コンストラクタ
    WinSCPDriver()
    {
        try
        {
            $this.Common = $global:Common
            $this.InitializeWinSCP()
            $this.Common.WriteLog("WinSCPDriverが正常に初期化されました。", "INFO")
        }
        catch
        {
            # 初期化失敗時のクリーンアップ
            Write-Host "WinSCPDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()

            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0001", "WinSCPDriver初期化エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WinSCPDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "WinSCPDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # WinSCPの初期化
    [void] InitializeWinSCP()
    {
        try
        {
            # WinSCPのインストールパスを検索
            $this.WinSCPPath = $this.FindWinSCPPath()
            
            if ([string]::IsNullOrEmpty($this.WinSCPPath))
            {
                throw "WinSCPのインストールパスが見つかりません。"
            }

            # WinSCPアセンブリを読み込み
            $assemblyPath = Join-Path $this.WinSCPPath "WinSCPnet.dll"
            
            if (-not (Test-Path $assemblyPath))
            {
                throw "WinSCPアセンブリが見つかりません: $assemblyPath"
            }

            # アセンブリを読み込み
            Add-Type -Path $assemblyPath
            
            # セッションオプションを初期化
            $this.SessionOptions = New-Object WinSCP.SessionOptions
            
            # 転送オプションを初期化
            $this.TransferOptions = New-Object WinSCP.TransferOptions
            $this.TransferOptions.TransferMode = [WinSCP.TransferMode]::Binary
            
            # セッションを初期化
            $this.Session = New-Object WinSCP.Session
            
            $this.Common.WriteLog("WinSCPアセンブリが正常に読み込まれました。", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0002", "WinSCPアセンブリ読み込みエラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WinSCPアセンブリの読み込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "WinSCPアセンブリの読み込みに失敗しました: $($_.Exception.Message)"
        }
    }

    # WinSCPのインストールパスを検索
    [string] FindWinSCPPath()
    {
        try
        {
            # 一般的なインストールパスをチェック
            $possiblePaths = @(
                "${env:ProgramFiles}\WinSCP",
                "${env:ProgramFiles(x86)}\WinSCP",
                "C:\Program Files\WinSCP",
                "C:\Program Files (x86)\WinSCP"
            )

            foreach ($path in $possiblePaths)
            {
                if (Test-Path $path)
                {
                    $this.Common.WriteLog("WinSCPパスが見つかりました: $path", "INFO")
                    return $path
                }
            }

            # レジストリから検索
            try
            {
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
                $installedPrograms = Get-ItemProperty $regPath | Where-Object { $_.DisplayName -like "*WinSCP*" }
                
                if ($installedPrograms)
                {
                    $installLocation = $installedPrograms[0].InstallLocation
                    if (-not [string]::IsNullOrEmpty($installLocation) -and (Test-Path $installLocation))
                    {
                        $this.Common.WriteLog("レジストリからWinSCPパスが見つかりました: $installLocation", "INFO")
                        return $installLocation
                    }
                }
            }
            catch
            {
                $this.Common.WriteLog("レジストリ検索でエラーが発生しました: $($_.Exception.Message)", "WARNING")
            }

            return ""
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0002", "WinSCPパス検索エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WinSCPパスの検索に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            return ""
        }
    }

    # 接続パラメータを設定
    [void] SetConnectionParameters([string]$hostName, [string]$userName, [string]$password, [string]$protocol = "SFTP", [int]$port = 0, [string]$hostKeyFingerprint = "")
    {
        try
        {
            if ($this.SessionOptions -eq $null)
            {
                throw "セッションオプションが初期化されていません。"
            }

            # プロトコルを設定
            switch ($protocol.ToUpper())
            {
                "SFTP" { $this.SessionOptions.Protocol = [WinSCP.Protocol]::Sftp }
                "FTP" { $this.SessionOptions.Protocol = [WinSCP.Protocol]::Ftp }
                "SCP" { $this.SessionOptions.Protocol = [WinSCP.Protocol]::Scp }
                "FTPS" { $this.SessionOptions.Protocol = [WinSCP.Protocol]::Ftps }
                default { throw "サポートされていないプロトコルです: $protocol" }
            }

            # 接続情報を設定
            $this.SessionOptions.HostName = $hostName
            $this.SessionOptions.UserName = $userName
            $this.SessionOptions.Password = $password

            # ポートが指定されている場合は設定
            if ($port -gt 0)
            {
                $this.SessionOptions.PortNumber = $port
            }

            # ホストキーフィンガープリントが指定されている場合は設定
            if (-not [string]::IsNullOrEmpty($hostKeyFingerprint))
            {
                $this.SessionOptions.SshHostKeyFingerprint = $hostKeyFingerprint
            }

            # タイムアウト設定
            $this.SessionOptions.Timeout = [TimeSpan]::FromSeconds(30)

            $this.Common.WriteLog("接続パラメータが設定されました。ホスト: $hostName, ユーザー: $userName, プロトコル: $protocol", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0003", "接続パラメータ設定エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
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
            if ($this.IsConnected)
            {
                $this.Common.WriteLog("既に接続されています。", "WARNING")
                return
            }

            if ($this.SessionOptions -eq $null)
            {
                throw "接続パラメータが設定されていません。SetConnectionParameters()を先に呼び出してください。"
            }

            # セッションを開く
            $this.Session.Open($this.SessionOptions)
            $this.IsConnected = $true

            $this.Common.WriteLog("サーバーに正常に接続されました。", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0004", "サーバー接続エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
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

    # 接続を切断
    [void] Disconnect()
    {
        try
        {
            if (-not $this.IsConnected)
            {
                $this.Common.WriteLog("接続されていません。", "WARNING")
                return
            }

            if ($this.Session -ne $null)
            {
                $this.Session.Close()
                $this.IsConnected = $false
                $this.Common.WriteLog("接続が正常に切断されました。", "INFO")
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0005", "接続切断エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
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

    # ファイルをアップロード
    [void] UploadFile([string]$localPath, [string]$remotePath)
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            if (-not (Test-Path $localPath))
            {
                throw "ローカルファイルが見つかりません: $localPath"
            }

            # ファイルをアップロード
            $transferResult = $this.Session.PutFiles($localPath, $remotePath, $false, $this.TransferOptions)
            
            # 転送結果をチェック
            $transferResult.Check()

            $this.Common.WriteLog("ファイルが正常にアップロードされました。ローカル: $localPath, リモート: $remotePath", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0008", "ファイルアップロードエラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
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

    # ファイルをダウンロード
    [void] DownloadFile([string]$remotePath, [string]$localPath)
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            # ローカルディレクトリが存在しない場合は作成
            $localDir = Split-Path $localPath -Parent
            if (-not (Test-Path $localDir))
            {
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
            }

            # ファイルをダウンロード
            $transferResult = $this.Session.GetFiles($remotePath, $localPath, $false, $this.TransferOptions)
            
            # 転送結果をチェック
            $transferResult.Check()

            $this.Common.WriteLog("ファイルが正常にダウンロードされました。リモート: $remotePath, ローカル: $localPath", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0009", "ファイルダウンロードエラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
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

    # ファイルを削除
    [void] RemoveFile([string]$remotePath)
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            # ファイルを削除
            $this.Session.RemoveFiles($remotePath)

            $this.Common.WriteLog("ファイルが正常に削除されました: $remotePath", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0010", "ファイル削除エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ファイルの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ファイルの削除に失敗しました: $($_.Exception.Message)"
        }
    }

    # ディレクトリを作成
    [void] CreateDirectory([string]$remotePath)
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            # ディレクトリを作成
            $this.Session.CreateDirectory($remotePath)

            $this.Common.WriteLog("ディレクトリが正常に作成されました: $remotePath", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0011", "ディレクトリ作成エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ディレクトリの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # ディレクトリを削除
    [void] RemoveDirectory([string]$remotePath)
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            # ディレクトリを削除
            $this.Session.RemoveFiles($remotePath)

            $this.Common.WriteLog("ディレクトリが正常に削除されました: $remotePath", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0012", "ディレクトリ削除エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ディレクトリの削除に失敗しました: $($_.Exception.Message)"
        }
    }

    # ファイル一覧を取得
    [array] GetFileList([string]$remotePath = "/")
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            # ファイル一覧を取得
            $fileInfos = $this.Session.ListDirectory($remotePath)
            
            $fileList = @()
            foreach ($fileInfo in $fileInfos.Files)
            {
                # 親ディレクトリとカレントディレクトリを除外
                if ($fileInfo.Name -ne "." -and $fileInfo.Name -ne "..")
                {
                    $fileList += @{
                        Name = $fileInfo.Name
                        FullName = $fileInfo.FullName
                        IsDirectory = $fileInfo.IsDirectory
                        Size = $fileInfo.Length
                        LastWriteTime = $fileInfo.LastWriteTime
                        Permissions = $fileInfo.FilePermissions
                    }
                }
            }

            $this.Common.WriteLog("ファイル一覧を取得しました。パス: $remotePath, 件数: $($fileList.Count)", "INFO")
            return $fileList
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0013", "ファイル一覧取得エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ファイル一覧の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ファイル一覧の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ファイルの存在確認
    [bool] FileExists([string]$remotePath)
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            # ファイル情報を取得
            $fileInfo = $this.Session.GetFileInfo($remotePath)
            $exists = $fileInfo -ne $null

            $this.Common.WriteLog("ファイル存在確認完了: $remotePath, 存在: $exists", "DEBUG")
            return $exists
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0014", "ファイル存在確認エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ファイルの存在確認に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            return $false
        }
    }

    # ディレクトリの存在確認
    [bool] DirectoryExists([string]$remotePath)
    {
        try
        {
            if (-not $this.IsConnected)
            {
                throw "接続されていません。先にConnect()を呼び出してください。"
            }

            # ディレクトリ情報を取得
            $fileInfo = $this.Session.GetFileInfo($remotePath)
            $exists = $fileInfo -ne $null -and $fileInfo.IsDirectory

            $this.Common.WriteLog("ディレクトリ存在確認完了: $remotePath, 存在: $exists", "DEBUG")
            return $exists
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0015", "ディレクトリ存在確認エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ディレクトリの存在確認に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            return $false
        }
    }

    # 転送オプションを設定
    [void] SetTransferOptions([string]$transferMode = "Binary", [bool]$overwrite = $false)
    {
        try
        {
            if ($this.TransferOptions -eq $null)
            {
                $this.TransferOptions = New-Object WinSCP.TransferOptions
            }

            # 転送モードを設定
            switch ($transferMode.ToUpper())
            {
                "BINARY" { $this.TransferOptions.TransferMode = [WinSCP.TransferMode]::Binary }
                "ASCII" { $this.TransferOptions.TransferMode = [WinSCP.TransferMode]::Ascii }
                "AUTOMATIC" { $this.TransferOptions.TransferMode = [WinSCP.TransferMode]::Automatic }
                default { throw "サポートされていない転送モードです: $transferMode" }
            }

            # 上書き設定
            $this.TransferOptions.OverwriteMode = if ($overwrite) { [WinSCP.OverwriteMode]::Overwrite } else { [WinSCP.OverwriteMode]::Resume }

            $this.Common.WriteLog("転送オプションが設定されました。モード: $transferMode, 上書き: $overwrite", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0020", "転送オプション設定エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "転送オプションの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "転送オプションの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続状態を取得
    [bool] IsConnectedToServer()
    {
        return $this.IsConnected
    }

    # リソースの破棄
    [void] Dispose()
    {
        try
        {
            if ($this.IsConnected)
            {
                $this.Disconnect()
            }

            if ($this.Session -ne $null)
            {
                $this.Session.Dispose()
                $this.Session = $null
            }

            $this.SessionOptions = $null
            $this.TransferOptions = $null

            $this.Common.WriteLog("WinSCPDriverが正常に破棄されました。", "INFO")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WinSCPError_0006", "WinSCPDriver破棄エラー: $($_.Exception.Message)", "WinSCPDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WinSCPDriverの破棄に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            # セッションが存在する場合は破棄
            if ($this.Session -ne $null)
            {
                try
                {
                    if ($this.IsConnected)
                    {
                        $this.Session.Close()
                    }
                    $this.Session.Dispose()
                }
                catch
                {
                    # セッションの破棄でエラーが発生しても無視
                }
                $this.Session = $null
            }

            # その他のオブジェクトをクリア
            $this.SessionOptions = $null
            $this.TransferOptions = $null
            $this.IsConnected = $false

            Write-Host "WinSCPDriver初期化失敗時のクリーンアップが完了しました。" -ForegroundColor Yellow
        }
        catch
        {
            Write-Host "WinSCPDriver初期化失敗時のクリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # デストラクタ
    [void] Finalize()
    {
        $this.Dispose()
    }
}

Write-Host "WinSCPDriverクラスが正常にインポートされました。" -ForegroundColor Green
