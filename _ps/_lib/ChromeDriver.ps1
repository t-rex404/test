# ChromeDriverエラー管理モジュールをインポート
. "$PSScriptRoot\ChromeDriverErrors.ps1"

class ChromeDriver : WebDriver
{
    [string]$browser_exe_path
    [string]$browser_user_data_dir
    [bool]$is_chrome_initialized

    ChromeDriver()
    {
        try
        {
            $this.is_chrome_initialized = $false
            
            # ブラウザの実行ファイルのパスを取得
            $this.browser_exe_path = $this.GetChromeExecutablePath()
            
            # ユーザーデータディレクトリの設定
            $this.browser_user_data_dir = $this.GetUserDataDirectory()
            
            # ブラウザの実行
            $this.StartBrowser($this.browser_exe_path, $this.browser_user_data_dir)

            # タブ情報を取得
            $tab_infomation = $this.GetTabInfomation()

            # WebSocket接続
            $this.GetWebSocketInfomation($tab_infomation.webSocketDebuggerUrl)

            # タブをアクティブにする
            $this.SetActiveTab($tab_infomation.id)

            # デバッグモードを有効化
            $this.EnableDebugMode()

            $this.is_chrome_initialized = $true
            Write-Host "ChromeDriverの初期化が完了しました。"
        }
        catch
        {
            # 初期化に失敗した場合のクリーンアップ
            $this.CleanupOnInitializationFailure()
            LogChromeDriverError $ChromeDriverErrorCodes.INIT_ERROR "ChromeDriver初期化エラー: $($_.Exception.Message)"
            throw "ChromeDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # Chrome実行ファイルのパスを取得
    [string] GetChromeExecutablePath()
    {
        try
        {
            # 複数のレジストリパスを試行
            $registry_paths = @(
                'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\',
                'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\',
                'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\'
            )

            foreach ($registry_path in $registry_paths)
            {
                try
                {
                    $browser_exe_path = Get-ItemPropertyValue -Path $registry_path -Name '(default)' -ErrorAction Stop
                    if ($browser_exe_path -and (Test-Path $browser_exe_path))
                    {
                        Write-Host "Chrome実行ファイルが見つかりました: $browser_exe_path"
                        return $browser_exe_path
                    }
                }
                catch
                {
                    Write-Host "レジストリパス $registry_path での検索に失敗しました: $($_.Exception.Message)"
                    continue
                }
            }

            # レジストリで見つからない場合、一般的なパスを試行
            $common_paths = @(
                "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
                "${env:LOCALAPPDATA}\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles}\Google\Chrome Beta\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome Beta\Application\chrome.exe",
                "${env:ProgramFiles}\Google\Chrome SxS\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome SxS\Application\chrome.exe"
            )

            foreach ($common_path in $common_paths)
            {
                if (Test-Path $common_path)
                {
                    Write-Host "Chrome実行ファイルが見つかりました: $common_path"
                    return $common_path
                }
            }

            throw "Chrome実行ファイルが見つかりませんでした。Chromeがインストールされているか確認してください。"
        }
        catch
        {
            LogChromeDriverError $ChromeDriverErrorCodes.EXECUTABLE_PATH_ERROR "Chrome実行ファイルパス取得エラー: $($_.Exception.Message)"
            throw "Chrome実行ファイルのパス取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ユーザーデータディレクトリを取得
    [string] GetUserDataDirectory()
    {
        try
        {
            # 環境変数から取得を試行
            $user_data_dir = $env:CHROME_USER_DATA_DIR
            if (-not [string]::IsNullOrEmpty($user_data_dir))
            {
                if (-not (Test-Path $user_data_dir))
                {
                    New-Item -ItemType Directory -Path $user_data_dir -Force | Out-Null
                }
                return $user_data_dir
            }

            # デフォルトパスを使用
            $default_dir = Join-Path $env:TEMP "ChromeDriver_UserData_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            if (-not (Test-Path $default_dir))
            {
                New-Item -ItemType Directory -Path $default_dir -Force | Out-Null
            }
            
            Write-Host "ユーザーデータディレクトリを作成しました: $default_dir"
            return $default_dir
        }
        catch
        {
            LogChromeDriverError $ChromeDriverErrorCodes.USER_DATA_DIR_ERROR "ユーザーデータディレクトリ作成エラー: $($_.Exception.Message)"
            throw "ユーザーデータディレクトリの作成に失敗しました: $($_.Exception.Message)"
        }
    }

    # デバッグモードを有効化
    [void] EnableDebugMode()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # エミュレーションメディアを設定
            $this.SendWebSocketMessage('Emulation.setEmulatedMedia', @{ media = 'screen' })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                Write-Host "エミュレーションメディア設定に失敗しましたが、処理を続行します: $($response.error.message)"
            }
            else
            {
                Write-Host "デバッグモードが有効化されました。"
            }
        }
        catch
        {
            LogChromeDriverError $ChromeDriverErrorCodes.DEBUG_MODE_ERROR "デバッグモード有効化エラー: $($_.Exception.Message)"
            Write-Host "デバッグモードの有効化に失敗しましたが、処理を続行します: $($_.Exception.Message)"
        }
    }

    # 初期化失敗時のクリーンアップ
    [void] CleanupOnInitializationFailure()
    {
        try
        {
            if ($this.is_initialized)
            {
                $this.Dispose()
            }
            else
            {
                # 部分的な初期化状態のクリーンアップ
                if ($this.web_socket -and $this.web_socket.State -eq [System.Net.WebSockets.WebSocketState]::Open)
                {
                    $this.web_socket.Dispose()
                    $this.web_socket = $null
                }
                
                if ($this.browser_exe_process_id -gt 0)
                {
                    try
                    {
                        Stop-Process -Id $this.browser_exe_process_id -Force -ErrorAction SilentlyContinue
                    }
                    catch
                    {
                        Write-Host "プロセス終了に失敗しました: $($_.Exception.Message)"
                    }
                    finally
                    {
                        $this.browser_exe_process_id = 0
                    }
                }
            }
            
            $this.is_chrome_initialized = $false
            Write-Host "初期化失敗時のクリーンアップが完了しました。"
        }
        catch
        {
            Write-Host "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)"
        }
    }

    # カスタムDisposeメソッド
    [void] Dispose()
    {
        try
        {
            if ($this.is_chrome_initialized)
            {
                # 親クラスのDisposeを呼び出し
                [base]::Dispose()
                
                # Chrome固有のクリーンアップ
                if (-not [string]::IsNullOrEmpty($this.browser_user_data_dir) -and (Test-Path $this.browser_user_data_dir))
                {
                    try
                    {
                        # ユーザーデータディレクトリの削除（オプション）
                        # Remove-Item -Path $this.browser_user_data_dir -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Host "ユーザーデータディレクトリを保持しました: $($this.browser_user_data_dir)"
                    }
                    catch
                    {
                        Write-Host "ユーザーデータディレクトリの削除に失敗しました: $($_.Exception.Message)"
                    }
                }
                
                $this.is_chrome_initialized = $false
                Write-Host "ChromeDriverのリソースを正常に解放しました。"
            }
        }
        catch
        {
            LogChromeDriverError $ChromeDriverErrorCodes.DISPOSE_ERROR "ChromeDriver Disposeエラー: $($_.Exception.Message)"
            Write-Host "ChromeDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)"
        }
    }
}
