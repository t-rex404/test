# WinSCPを操作するためのPowerShellクラス
# ファイル転送、接続管理などの機能を提供

# WinSCPアセンブリの事前読み込み
# クラス定義前にアセンブリを読み込む必要がある
$script:WinSCPAssemblyLoaded = $false
if (-not $script:WinSCPAssemblyLoaded)
{
    # WinSCPのインストールパスを検索
    $possiblePaths = @(
        "${env:ProgramFiles}\WinSCP\WinSCPnet.dll",
        "${env:ProgramFiles(x86)}\WinSCP\WinSCPnet.dll",
        "C:\Program Files\WinSCP\WinSCPnet.dll",
        "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"
    )

    $assemblyPath = $null
    foreach ($path in $possiblePaths)
    {
        if (Test-Path $path)
        {
            $assemblyPath = $path
            break
        }
    }

    if ($assemblyPath)
    {
        try
        {
            Add-Type -Path $assemblyPath
            $script:WinSCPAssemblyLoaded = $true
            Write-Host "WinSCPアセンブリを読み込みました: $assemblyPath" -ForegroundColor Green
        }
        catch
        {
            Write-Host "WinSCPアセンブリの読み込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "WinSCPがインストールされていることを確認してください。" -ForegroundColor Yellow
            throw
        }
    }
    else
    {
        Write-Host "WinSCPが見つかりません。WinSCPをインストールしてください。" -ForegroundColor Red
        Write-Host "ダウンロード: https://winscp.net/" -ForegroundColor Yellow
        throw "WinSCPアセンブリが見つかりません"
    }
}

class WinSCPDriver
{
    # プロパティ
    [object]$Session
    [object]$SessionOptions
    [object]$TransferOptions
    [string]$WinSCPPath
    [bool]$IsConnected = $false

    # ログファイルパス（共有可能）
    static [string]$NormalLogFile = ".\WinSCPDriver_$($env:USERNAME)_Normal.log"
    static [string]$ErrorLogFile = ".\WinSCPDriver_$($env:USERNAME)_Error.log"

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
                $global:Common.WriteLog($message, "INFO", "WinSCPDriver")
            }
            catch
            {
                Write-Host "正常ログ出力に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([WinSCPDriver]::NormalLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Green
    }

    # エラーログを出力
    [void] LogError([string]$errorCode, [string]$message)
    {
        if ($global:Common)
        {
            try
            {
                $global:Common.HandleError($errorCode, $message, "WinSCPDriver")
            }
            catch
            {
                Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([WinSCPDriver]::ErrorLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Red
    }

    # ========================================
    # 初期化・接続関連
    # ========================================

    # コンストラクタ
    WinSCPDriver()
    {
        try
        {
            $this.IsConnected = $false

            # セッションオプションを初期化
            $this.SessionOptions = New-Object WinSCP.SessionOptions

            # 転送オプションを初期化
            $this.TransferOptions = New-Object WinSCP.TransferOptions
            $this.TransferOptions.TransferMode = [WinSCP.TransferMode]::Binary

            # セッションを初期化
            $this.Session = New-Object WinSCP.Session

            Write-Host "WinSCPDriverが正常に初期化されました。" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("WinSCPDriver initialization completed")
        }
        catch
        {
            # 初期化失敗時のクリーンアップ
            Write-Host "WinSCPDriver初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            $this.CleanupOnInitializationFailure()

            $this.LogError("WinSCPDriverError_0001", "WinSCPDriver初期化エラー: $($_.Exception.Message)")
            throw "WinSCPDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続パラメータを設定（パスワード認証）
    [void] SetConnectionParameters([string]$hostName, [string]$userName, [string]$password, [string]$protocol = "SFTP", [int]$port = 0, [string]$hostKeyFingerprint = "")
    {
        try
        {
            if ($null -eq $this.SessionOptions)
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

            # 正常ログ出力
            $this.LogInfo("接続パラメータが設定されました。ホスト: $hostName, ユーザー: $userName, プロトコル: $protocol")
            Write-Host "接続パラメータが設定されました。ホスト: $hostName, ユーザー: $userName" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0010", "接続パラメータ設定エラー: $($_.Exception.Message)")

            throw "接続パラメータの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続パラメータを設定（秘密鍵認証）
    [void] SetConnectionParametersWithKey([string]$hostName, [string]$userName, [string]$privateKeyPath, [string]$passphrase = "", [string]$protocol = "SFTP", [int]$port = 0, [string]$hostKeyFingerprint = "")
    {
        try
        {
            if ($null -eq $this.SessionOptions)
            {
                throw "セッションオプションが初期化されていません。"
            }

            # 秘密鍵ファイルの存在確認
            if (-not (Test-Path $privateKeyPath))
            {
                throw "秘密鍵ファイルが見つかりません: $privateKeyPath"
            }

            # プロトコルを設定（SSH系のみサポート）
            switch ($protocol.ToUpper())
            {
                "SFTP" { $this.SessionOptions.Protocol = [WinSCP.Protocol]::Sftp }
                "SCP" { $this.SessionOptions.Protocol = [WinSCP.Protocol]::Scp }
                default { throw "秘密鍵認証は SFTP/SCP プロトコルでのみサポートされています: $protocol" }
            }

            # 接続情報を設定
            $this.SessionOptions.HostName = $hostName
            $this.SessionOptions.UserName = $userName

            # 秘密鍵の設定
            $this.SessionOptions.SshPrivateKeyPath = $privateKeyPath
            if (-not [string]::IsNullOrEmpty($passphrase))
            {
                $this.SessionOptions.SshPrivateKeyPassphrase = $passphrase
            }

            # ポートが指定されている場合は設定
            if ($port -gt 0)
            {
                $this.SessionOptions.PortNumber = $port
            }
            else
            {
                # SSH系のデフォルトポート
                $this.SessionOptions.PortNumber = 22
            }

            # ホストキーフィンガープリントが指定されている場合は設定
            if (-not [string]::IsNullOrEmpty($hostKeyFingerprint))
            {
                $this.SessionOptions.SshHostKeyFingerprint = $hostKeyFingerprint
            }
            else
            {
                # セキュリティ警告: 本番環境では必ず正しいフィンガープリントを設定してください
                $this.SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
            }

            # タイムアウト設定
            $this.SessionOptions.Timeout = [TimeSpan]::FromSeconds(30)

            # 正常ログ出力
            $this.LogInfo("接続パラメータが設定されました（秘密鍵認証）。ホスト: $hostName, ユーザー: $userName, 鍵: $privateKeyPath")
            Write-Host "接続パラメータが設定されました（秘密鍵認証）。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0013", "秘密鍵接続パラメータ設定エラー: $($_.Exception.Message)")

            throw "秘密鍵接続パラメータの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # サーバーに接続
    [void] Connect()
    {
        try
        {
            if ($this.IsConnected)
            {
                $this.LogInfo("既に接続されています。")
                Write-Host "既に接続されています。" -ForegroundColor Yellow
                return
            }

            if ($null -eq $this.SessionOptions)
            {
                throw "接続パラメータが設定されていません。SetConnectionParameters()を先に呼び出してください。"
            }

            # セッションを開く
            $this.Session.Open($this.SessionOptions)
            $this.IsConnected = $true

            # 正常ログ出力
            $this.LogInfo("サーバーに正常に接続されました。")
            Write-Host "サーバーに正常に接続されました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0011", "サーバー接続エラー: $($_.Exception.Message)")

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
                $this.LogInfo("接続されていません。")
                Write-Host "接続されていません。" -ForegroundColor Yellow
                return
            }

            if ($null -ne $this.Session)
            {
                $this.Session.Close()
                $this.IsConnected = $false

                # 正常ログ出力
                $this.LogInfo("接続が正常に切断されました。")
                Write-Host "接続が正常に切断されました。" -ForegroundColor Green
            }
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0012", "接続切断エラー: $($_.Exception.Message)")

            throw "接続の切断に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ファイル操作関連
    # ========================================

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

            # 正常ログ出力
            $this.LogInfo("ファイルが正常にアップロードされました。ローカル: $localPath, リモート: $remotePath")
            Write-Host "ファイルが正常にアップロードされました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0020", "ファイルアップロードエラー: $($_.Exception.Message)")

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

            # 正常ログ出力
            $this.LogInfo("ファイルが正常にダウンロードされました。リモート: $remotePath, ローカル: $localPath")
            Write-Host "ファイルが正常にダウンロードされました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0021", "ファイルダウンロードエラー: $($_.Exception.Message)")

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

            # 正常ログ出力
            $this.LogInfo("ファイルが正常に削除されました: $remotePath")
            Write-Host "ファイルが正常に削除されました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0022", "ファイル削除エラー: $($_.Exception.Message)")

            throw "ファイルの削除に失敗しました: $($_.Exception.Message)"
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
            try
            {
                $fileInfo = $this.Session.GetFileInfo($remotePath)
                $exists = $null -ne $fileInfo

                # 正常ログ出力
                $this.LogInfo("ファイル存在確認完了: $remotePath, 存在: $exists")

                return $exists
            }
            catch [WinSCP.SessionRemoteException]
            {
                # ファイルが存在しない場合の正常な動作
                $this.LogInfo("ファイル存在確認完了: $remotePath, 存在: False")

                return $false
            }
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0041", "ファイル存在確認エラー: $($_.Exception.Message)")

            return $false
        }
    }

    # ========================================
    # ディレクトリ操作関連
    # ========================================

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

            # 正常ログ出力
            $this.LogInfo("ディレクトリが正常に作成されました: $remotePath")
            Write-Host "ディレクトリが正常に作成されました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0030", "ディレクトリ作成エラー: $($_.Exception.Message)")

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

            # 正常ログ出力
            $this.LogInfo("ディレクトリが正常に削除されました: $remotePath")
            Write-Host "ディレクトリが正常に削除されました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0031", "ディレクトリ削除エラー: $($_.Exception.Message)")

            throw "ディレクトリの削除に失敗しました: $($_.Exception.Message)"
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
            try
            {
                $fileInfo = $this.Session.GetFileInfo($remotePath)
                $exists = ($null -ne $fileInfo) -and $fileInfo.IsDirectory

                # 正常ログ出力
                $this.LogInfo("ディレクトリ存在確認完了: $remotePath, 存在: $exists")

                return $exists
            }
            catch [WinSCP.SessionRemoteException]
            {
                # ディレクトリが存在しない場合の正常な動作
                $this.LogInfo("ディレクトリ存在確認完了: $remotePath, 存在: False")

                return $false
            }
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0042", "ディレクトリ存在確認エラー: $($_.Exception.Message)")

            return $false
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

            # 正常ログ出力
            $this.LogInfo("ファイル一覧を取得しました。パス: $remotePath, 件数: $($fileList.Count)")
            Write-Host "ファイル一覧を取得しました。件数: $($fileList.Count)" -ForegroundColor Green

            return $fileList
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0040", "ファイル一覧取得エラー: $($_.Exception.Message)")

            throw "ファイル一覧の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 設定・ユーティリティ関連
    # ========================================

    # 転送オプションを設定
    [void] SetTransferOptions([string]$transferMode = "Binary", [bool]$overwrite = $false)
    {
        try
        {
            if ($null -eq $this.TransferOptions)
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
            $this.TransferOptions.OverwriteMode = if ($overwrite) { [WinSCP.OverwriteMode]::Overwrite } else { [WinSCP.OverwriteMode]::Exception }

            # 正常ログ出力
            $this.LogInfo("転送オプションが設定されました。モード: $transferMode, 上書き: $overwrite")
            Write-Host "転送オプションが設定されました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0080", "転送オプション設定エラー: $($_.Exception.Message)")

            throw "転送オプションの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 接続状態を取得
    [bool] IsConnectedToServer()
    {
        if ($null -eq $this.Session)
        {
            $this.IsConnected = $false
            return $false
        }

        try
        {
            # 実際のセッション状態を確認
            $opened = $this.Session.Opened
            $this.IsConnected = $opened
            return $opened
        }
        catch
        {
            $this.IsConnected = $false
            return $false
        }
    }

    # ========================================
    # クリーンアップ・破棄関連
    # ========================================

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            # セッションが存在する場合は破棄
            if ($null -ne $this.Session)
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

    # リソースの破棄
    [void] Dispose()
    {
        try
        {
            if ($this.IsConnected)
            {
                $this.Disconnect()
            }

            if ($null -ne $this.Session)
            {
                $this.Session.Dispose()
                $this.Session = $null
            }

            $this.SessionOptions = $null
            $this.TransferOptions = $null

            # 正常ログ出力
            $this.LogInfo("WinSCPDriverが正常に破棄されました。")
            Write-Host "WinSCPDriverが正常に破棄されました。" -ForegroundColor Green
        }
        catch
        {
            $this.LogError("WinSCPDriverError_0090", "WinSCPDriver破棄エラー: $($_.Exception.Message)")
        }
    }
}

Write-Host "WinSCPDriverクラスが正常にインポートされました。" -ForegroundColor Green
