# ChromeDriverエラー管理モジュールをインポート
#import-module "$PSScriptRoot\ChromeDriverErrors.psm1"

# 共通ライブラリをインポート
. "$PSScriptRoot\Common.ps1"

class ChromeDriver : WebDriver
{
    [string]$browser_exe_path
    [string]$browser_user_data_dir
    [bool]$is_chrome_initialized

    # ========================================
    # 初期化・接続関連
    # ========================================

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
            $global:Common.HandleError("3001", "ChromeDriver初期化エラー: $($_.Exception.Message)", "ChromeDriver", ".\ChromeDriver_Error.log")
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
                    $this.browser_exe_path = Get-ItemPropertyValue -Path $registry_path -Name '(default)' -ErrorAction Stop
                    if ($this.browser_exe_path -and (Test-Path $this.browser_exe_path))
                    {
                        Write-Host "Chrome実行ファイルが見つかりました: $this.browser_exe_path"
                        return $this.browser_exe_path
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
            $global:Common.HandleError("3002", "Chrome実行ファイルパス取得エラー: $($_.Exception.Message)", "ChromeDriver", ".\ChromeDriver_Error.log")
            throw "Chrome実行ファイルのパス取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ユーザーデータディレクトリを取得
    [string] GetUserDataDirectory()
    {
        try
        {
            # ユーザーデータディレクトリのパスを生成
            #$user_data_dir = Join-Path $env:TEMP "ChromeDriver_UserData_$(Get-Random)"
            $user_data_dir = 'C:\temp\ChromeDriver_UserData'
            
            if (Test-Path $user_data_dir)
            {
                Remove-Item -Path $user_data_dir -Recurse -Force
            }

            Write-Host "Chromeユーザーデータディレクトリを作成しました: $user_data_dir"
            return $user_data_dir
        }
        catch
        {
            $global:Common.HandleError("3003", "ユーザーデータディレクトリ取得エラー: $($_.Exception.Message)", "ChromeDriver", ".\ChromeDriver_Error.log")
            throw "ユーザーデータディレクトリの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 設定・初期化関連
    # ========================================

    # デバッグモードを有効化
    [void] EnableDebugMode()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # ページイベントを有効化
            $this.EnablePageEvents()
            
            Write-Host "Chromeデバッグモードが有効化されました。"
        }
        catch
        {
            $global:Common.HandleError("3004", "デバッグモード有効化エラー: $($_.Exception.Message)", "ChromeDriver", ".\ChromeDriver_Error.log")
            throw "デバッグモードの有効化に失敗しました: $($_.Exception.Message)"
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
            
            # WebSocket接続を閉じる
            if ($this.web_socket -and $this.web_socket.State -eq 'Open')
            {
                try
                {
                    $this.web_socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "Initialization failed", $null).Wait()
                }
                catch
                {
                    Write-Host "WebSocket接続の閉じる際にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # ブラウザプロセスを終了
            if ($this.browser_exe_process_id -gt 0)
            {
                try
                {
                    $process = Get-Process -Id $this.browser_exe_process_id -ErrorAction SilentlyContinue
                    if ($process)
                    {
                        $process.Kill()
                        Write-Host "ブラウザプロセスを終了しました。プロセスID: $($this.browser_exe_process_id)" -ForegroundColor Yellow
                    }
                }
                catch
                {
                    Write-Host "ブラウザプロセスの終了に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # ユーザーデータディレクトリを削除
            if ($this.browser_user_data_dir -and (Test-Path $this.browser_user_data_dir))
            {
                try
                {
                    Remove-Item -Path $this.browser_user_data_dir -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "ユーザーデータディレクトリを削除しました: $($this.browser_user_data_dir)" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "ユーザーデータディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            Write-Host "クリーンアップが完了しました。" -ForegroundColor Yellow
        }
        catch
        {
            Write-Host "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # ========================================
    # リソース管理関連
    # ========================================

    # カスタムDisposeメソッド
    [void] Dispose()
    {
        try
        {
            Write-Host "ChromeDriverのリソースを解放します。" -ForegroundColor Cyan
            
            # 親クラスのDisposeを呼び出し
            [WebDriver]::Dispose($this)
            
            # Chrome固有のクリーンアップ
            $this.CleanupOnInitializationFailure()
            
            $this.is_chrome_initialized = $false
            Write-Host "ChromeDriverのリソース解放が完了しました。" -ForegroundColor Green
        }
        catch
        {
            $global:Common.HandleError("3005", "ChromeDriver Disposeエラー: $($_.Exception.Message)", "ChromeDriver", ".\ChromeDriver_Error.log")
            Write-Host "ChromeDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}



