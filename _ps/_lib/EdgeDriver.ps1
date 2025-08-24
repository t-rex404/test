# EdgeDriverエラー管理モジュールをインポート
#import-module "$PSScriptRoot\EdgeDriverErrors.psm1"

# 共通ライブラリをインポート
#. "$PSScriptRoot\Common.ps1"
#$Common = New-Object -TypeName 'Common'

class EdgeDriver : WebDriver
{
    [string]$browser_exe_path
    [string]$browser_user_data_dir
    [bool]$is_edge_initialized

    # ========================================
    # 初期化・接続関連
    # ========================================

    EdgeDriver()
    {
        try
        {
            $this.is_edge_initialized = $false
            
            # ブラウザの実行ファイルのパスを取得
            $this.browser_exe_path = $this.GetEdgeExecutablePath()
            
            # ユーザーデータディレクトリの設定
            $this.browser_user_data_dir = $this.GetUserDataDirectory()
            
            # ブラウザの実行
            Write-Host "ブラウザを起動しています..." -ForegroundColor Yellow
            $this.StartBrowser($this.browser_exe_path, $this.browser_user_data_dir)

            # タブ情報を取得（リトライ機能付き）
            Write-Host "タブ情報を取得しています..." -ForegroundColor Yellow
            $tab_infomation = $null
            $retry_count = 0
            $max_retries = 5
            
            while ($retry_count -lt $max_retries -and -not $tab_infomation)
            {
                try
                {
                    Start-Sleep -Seconds 2
                    $tab_infomation = $this.GetTabInfomation()
                    
                    if ($tab_infomation -and $tab_infomation.webSocketDebuggerUrl)
                    {
                        Write-Host "タブ情報の取得に成功しました。" -ForegroundColor Green
                        break
                    }
                    else
                    {
                        Write-Host "タブ情報の取得に失敗しました。再試行中... ($($retry_count + 1)/$max_retries)" -ForegroundColor Yellow
                        $tab_infomation = $null
                    }
                }
                catch
                {
                    Write-Host "タブ情報取得エラー（再試行 $($retry_count + 1)/$max_retries）: $($_.Exception.Message)" -ForegroundColor Yellow
                    $tab_infomation = $null
                }
                $retry_count++
            }
            
            if (-not $tab_infomation -or [string]::IsNullOrEmpty($tab_infomation.webSocketDebuggerUrl))
            {
                throw "タブ情報の取得に失敗しました。WebSocketデバッガーURLが取得できません。"
            }

            # WebSocket接続
            Write-Host "WebSocket接続を確立しています..." -ForegroundColor Yellow
            $this.GetWebSocketInfomation($tab_infomation.webSocketDebuggerUrl)

            # タブをアクティブにする
            Write-Host "タブをアクティブにしています..." -ForegroundColor Yellow
            $this.SetActiveTab($tab_infomation.id)

            # デバッグモードを有効化
            Write-Host "デバッグモードを有効化しています..." -ForegroundColor Yellow
            $this.EnableDebugMode()

            $this.is_edge_initialized = $true
            Write-Host "EdgeDriverの初期化が完了しました。" -ForegroundColor Green
        }
        catch
        {
            # 初期化に失敗した場合のクリーンアップ
            Write-Host "EdgeDriver初期化に失敗した場合のクリーンアップを開始します。" -ForegroundColor Yellow
            $this.CleanupOnInitializationFailure()

            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("2001", "EdgeDriver初期化エラー: $($_.Exception.Message)", "EdgeDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "EdgeDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "EdgeDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # Edge実行ファイルのパスを取得
    [string] GetEdgeExecutablePath()
    {
        try
        {
            # 複数のレジストリパスを試行
            $registry_paths = @(
                'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\',
                'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\',
                'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\'
            )

            foreach ($registry_path in $registry_paths)
            {
                try
                {
                    $this.browser_exe_path = Get-ItemPropertyValue -Path $registry_path -Name '(default)' -ErrorAction Stop
                    if ($this.browser_exe_path -and (Test-Path $this.browser_exe_path))
                    {
                        Write-Host "Edge実行ファイルが見つかりました: $this.browser_exe_path"
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
                "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
                "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                "${env:LOCALAPPDATA}\Microsoft\Edge\Application\msedge.exe"
            )

            foreach ($common_path in $common_paths)
            {
                if (Test-Path $common_path)
                {
                    Write-Host "Edge実行ファイルが見つかりました: $common_path"
                    return $common_path
                }
            }

            throw "Edge実行ファイルが見つかりませんでした。Edgeがインストールされているか確認してください。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("2002", "Edge実行ファイルパス取得エラー: $($_.Exception.Message)", "EdgeDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Edge実行ファイルのパス取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

            throw "Edge実行ファイルのパス取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ユーザーデータディレクトリを取得
    [string] GetUserDataDirectory()
    {
        try
        {
            # ユーザーデータディレクトリのパスを生成
            $base_dir = "C:\temp"
            if (-not (Test-Path $base_dir))
            {
                $base_dir = "C:\temp"
            }
            
            $user_data_dir = Join-Path $base_dir "EdgeDriver_UserData"
            
            # 既存のディレクトリが存在する場合は削除
            if (Test-Path $user_data_dir)
            {
                try
                {
                    Remove-Item -Path $user_data_dir -Recurse -Force -ErrorAction Stop
                    Write-Host "既存のユーザーデータディレクトリを削除しました: $user_data_dir" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "既存のユーザーデータディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # 新しいディレクトリを作成
            try
            {
                New-Item -ItemType Directory -Path $user_data_dir -Force -ErrorAction Stop | Out-Null
                Write-Host "Edgeユーザーデータディレクトリを作成しました: $user_data_dir" -ForegroundColor Green
            }
            catch
            {
                throw "ユーザーデータディレクトリの作成に失敗しました: $user_data_dir - $($_.Exception.Message)"
            }
            
            return $user_data_dir
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("2003", "ユーザーデータディレクトリ取得エラー: $($_.Exception.Message)", "EdgeDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ユーザーデータディレクトリの取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            
            Write-Host "Edgeデバッグモードが有効化されました。"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("2004", "デバッグモード有効化エラー: $($_.Exception.Message)", "EdgeDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "デバッグモードの有効化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
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
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("2005", "初期化失敗時のクリーンアップエラー: $($_.Exception.Message)", "EdgeDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "クリーンアップ中にエラーが発生しました: $($_.Exception.Message)"
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
            Write-Host "EdgeDriverのリソースを解放します。" -ForegroundColor Cyan
            
            # 親クラスのDisposeを呼び出し
            #[WebDriver]::Dispose($this)
            ([WebDriver]$this).Dispose()

            # Edge固有のクリーンアップ
            $this.CleanupOnInitializationFailure()
            
            $this.is_edge_initialized = $false
            Write-Host "EdgeDriverのリソース解放が完了しました。" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("2006", "EdgeDriver Disposeエラー: $($_.Exception.Message)", "EdgeDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "EdgeDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "EdgeDriverのリソース解放中にエラーが発生しました: $($_.Exception.Message)"
        }
    }
}
