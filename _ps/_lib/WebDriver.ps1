# WebDriverエラー管理モジュールをインポート
. "$PSScriptRoot\WebDriverErrors.ps1"

class WebDriver
{
    [int]$browser_exe_process_id
    [System.Net.WebSockets.ClientWebSocket]$web_socket
    [int]$message_id
    [bool]$is_initialized

    WebDriver()
    {
        try
        {
            $this.message_id = 0
            $this.is_initialized = $false
            $this.browser_exe_process_id = 0
            $this.web_socket = $null
        }
        catch
        {
            # 初期化・接続関連エラー (1001)
            LogWebDriverError $WebDriverErrorCodes.INIT_ERROR "WebDriver初期化エラー: $($_.Exception.Message)"
            throw "WebDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 初期化・接続関連
    # ========================================

    # ブラウザ起動
    [void] StartBrowser([string]$browser_exe_path, [string]$browser_user_data_dir)
    {
        try
        {
            # パラメータの検証
            if ([string]::IsNullOrEmpty($browser_exe_path))
            {
                throw "ブラウザの実行ファイルパスが指定されていません。"
            }
            
            if (-not (Test-Path $browser_exe_path))
            {
                throw "指定されたブラウザの実行ファイルが見つかりません: $browser_exe_path"
            }

            if ([string]::IsNullOrEmpty($browser_user_data_dir))
            {
                throw "ブラウザのユーザーデータディレクトリが指定されていません。"
            }

            # ユーザーデータディレクトリが存在しない場合は作成
            if (-not (Test-Path $browser_user_data_dir))
            {
                try
                {
                    New-Item -ItemType Directory -Path $browser_user_data_dir -Force | Out-Null
                }
                catch
                {
                    throw "ユーザーデータディレクトリの作成に失敗しました: $browser_user_data_dir"
                }
            }

            $url = 'about:blank'
            $argument_list = '--new-window ' + $url + ' --remote-debugging-port=9222 --disable-popup-blocking --no-first-run --user-data-dir=' + $browser_user_data_dir
                # 引数の意味
                # --new-window                 : 新しいウィンドウを開く。
                # --remote-debugging-port=9222 : デバッグ用 WebSocket をポート 9222 で有効化。
                # --disable-popup-blocking     : パップアップを無効化。
                # --no-first-run               : 最初の起動を無効化。
                # --user-data-dir              : ユーザーデータを指定(別のプロファイルを指定)（既存のブラウザに影響を与えないようにするため）。
                
            # ブラウザをデバッグモードで開く
            $browser_exe_process = Start-Process -FilePath $browser_exe_path -ArgumentList $argument_list -PassThru -ErrorAction Stop
            
            if (-not $browser_exe_process)
            {
                throw "ブラウザの起動に失敗しました。"
            }

            $this.browser_exe_process_id = $browser_exe_process.Id
            
            # ブラウザの起動を待機
            Start-Sleep -Seconds 3
            
            # プロセスが正常に起動しているか確認
            if (-not (Get-Process -Id $this.browser_exe_process_id -ErrorAction SilentlyContinue))
            {
                throw "ブラウザプロセスが正常に起動していません。"
            }

            Write-Host "ブラウザが正常に起動しました。プロセスID: $($this.browser_exe_process_id)"
        }
        catch
        {
            # 初期化・接続関連エラー (1002)
            LogWebDriverError $WebDriverErrorCodes.BROWSER_START_ERROR "ブラウザ起動エラー: $($_.Exception.Message)"
            throw "ブラウザの起動に失敗しました: $($_.Exception.Message)"
        }
    }

    # タブ情報を取得
    [Object] GetTabInfomation()
    {
        try
        {
            # デバッグ対象のWebSocket URLを取得(タブ情報を取得)
            $retry_count = 5
            $retry_delay = 2  # 秒
            
            for ($i = 0; $i -lt $retry_count; $i++)
            {
                try
                {
                    $tabs = Invoke-RestMethod -Uri 'http://localhost:9222/json' -ErrorAction Stop -TimeoutSec 10
                    
                    if (-not $tabs)
                    {
                        throw 'タブ情報を取得できません。'
                    }

                    $tab = $null
                    # 「about:blank」を選択
                    foreach ($tab in $tabs)
                    {
                        if ($tab.title -like 'about:blank' -and $tab.type -like 'page')
                        {
                            return $tab
                        }
                    }
                    
                    if ($i -lt $retry_count - 1)
                    {
                        Write-Host "about:blankタブが見つかりません。再試行します... ($($i+1)/$retry_count)"
                        Start-Sleep -Seconds $retry_delay
                    }
                    else
                    {
                        throw 'about:blankタブが見つかりません。'
                    }
                }
                catch
                {
                    if ($i -lt $retry_count - 1)
                    {
                        Write-Host "タブ情報の取得に失敗しました。再試行します... ($($i+1)/$retry_count)"
                        Start-Sleep -Seconds $retry_delay
                    }
                    else
                    {
                        throw
                    }
                }
            }
        }
        catch
        {
            # 初期化・接続関連エラー (1003)
            LogWebDriverError $WebDriverErrorCodes.TAB_INFO_ERROR "タブ情報取得エラー: $($_.Exception.Message)"
            throw "タブ情報の取得に失敗しました: $($_.Exception.Message)"
        }
    }
    
    # WebSocket接続
    [void] GetWebSocketInfomation([string]$web_socket_debugger_url)
    {
        try
        {
            if ([string]::IsNullOrEmpty($web_socket_debugger_url))
            {
                throw "WebSocketデバッガーURLが指定されていません。"
            }

            $retry_count = 3
            $retry_delay = 2  # 秒
            
            for ($i = 0; $i -lt $retry_count; $i++)
            {
                try
                {
                    # WebSocket接続の準備
                    $this.web_socket = [System.Net.WebSockets.ClientWebSocket]::new()
                    $uri = [System.Uri]::new($web_socket_debugger_url)
                    
                    # 接続タイムアウトを設定
                    $cancellationTokenSource = [System.Threading.CancellationTokenSource]::new()
                    $cancellationTokenSource.CancelAfter([TimeSpan]::FromSeconds(10))
                    
                    $this.web_socket.ConnectAsync($uri, $cancellationTokenSource.Token).Wait()
                    
                    if ($this.web_socket.State -eq [System.Net.WebSockets.WebSocketState]::Open)
                    {
                        $this.is_initialized = $true
                        Write-Host "WebSocket接続が確立されました。"
                        return
                    }
                    else
                    {
                        throw "WebSocket接続が確立されませんでした。状態: $($this.web_socket.State)"
                    }
                }
                catch
                {
                    if ($this.web_socket)
                    {
                        $this.web_socket.Dispose()
                        $this.web_socket = $null
                    }
                    
                    if ($i -lt $retry_count - 1)
                    {
                        Write-Host "WebSocket接続に失敗しました。再試行します... ($($i+1)/$retry_count)"
                        Start-Sleep -Seconds $retry_delay
                    }
                    else
                    {
                        throw
                    }
                }
                finally
                {
                    if ($cancellationTokenSource)
                    {
                        $cancellationTokenSource.Dispose()
                    }
                }
            }
        }
        catch
        {
            # 初期化・接続関連エラー (1004)
            LogWebDriverError $WebDriverErrorCodes.WEBSOCKET_CONNECTION_ERROR "WebSocket接続エラー: $($_.Exception.Message)"
            throw "WebSocket接続に失敗しました: $($_.Exception.Message)"
        }
    }

    # WebSocketメッセージ送信
    [void] SendWebSocketMessage([string]$method, [hashtable]$params)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            if ($this.web_socket.State -ne [System.Net.WebSockets.WebSocketState]::Open)
            {
                throw 'WebSocketが切断されています。ブラウザの状態を確認してください。'
            }
            
            if ([string]::IsNullOrEmpty($method))
            {
                throw "メソッド名が指定されていません。"
            }
            
            # WebSocketメッセージ送信メッセージ作成
            $this.message_id++
            $message = @{
                id = $this.message_id
                method = $method
                params = $params
            } | ConvertTo-Json -Depth 10 -Compress -ErrorAction Stop
            
            # WebSocketメッセージ送信
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($message)
            $buffer = [System.ArraySegment[byte]]::new($bytes)
            
            # 送信タイムアウトを設定
            $cancellationTokenSource = [System.Threading.CancellationTokenSource]::new()
            $cancellationTokenSource.CancelAfter([TimeSpan]::FromSeconds(10))
            
            $this.web_socket.SendAsync($buffer, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cancellationTokenSource.Token).Wait()
        }
        catch
        {
            # 初期化・接続関連エラー (1005)
            LogWebDriverError $WebDriverErrorCodes.WEBSOCKET_SEND_ERROR "WebSocket メッセージ送信エラー: $($_.Exception.Message)"
            throw "WebSocketメッセージ送信に失敗しました: $($_.Exception.Message)"
        }
        finally
        {
            if ($cancellationTokenSource)
            {
                $cancellationTokenSource.Dispose()
            }
        }
    }

    # WebSocketメッセージ受信
    [string] ReceiveWebSocketMessage()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            if ($this.web_socket.State -ne [System.Net.WebSockets.WebSocketState]::Open)
            {
                throw 'WebSocketが切断されています。ブラウザの状態を確認してください。'
            }
            
            $timeout = [datetime]::Now.AddSeconds(30)  # タイムアウトを30秒に延長

            # WebSocketメッセージ受信
            $responce_result = $false
            $response_json   = $null
            $retry_count = 0
            $max_retries = 3
            
            while(-not($responce_result))
            {
                if ([datetime]::Now -gt $timeout)
                {
                    throw 'WebSocketメッセージの受信がタイムアウトしました。'
                }
                
                if ($retry_count -ge $max_retries)
                {
                    throw "WebSocketメッセージの受信に失敗しました。最大再試行回数に達しました。"
                }

                $buffer_size = 4096
                $buffer = New-Object byte[] $buffer_size
                $message_stream = New-Object System.IO.MemoryStream
                
                try
                {
                    do
                    {
                        $segment = [System.ArraySegment[byte]]::new($buffer)
                        
                        # 受信タイムアウトを設定
                        $cancellationTokenSource = [System.Threading.CancellationTokenSource]::new()
                        $cancellationTokenSource.CancelAfter([TimeSpan]::FromSeconds(10))
                        
                        $receive_task = $this.web_socket.ReceiveAsync($segment, $cancellationTokenSource.Token)
                        $receive_task.Wait()
                        $result = $receive_task.Result
                        $message_stream.Write($buffer, 0, $result.Count)
                    } while (-not $result.EndOfMessage)
                        
                    # UTF8文字列に変換
                    $message_stream.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
                    $reader = New-Object System.IO.StreamReader($message_stream, [System.Text.Encoding]::UTF8)
                    $response_json = $reader.ReadToEnd()

                    try
                    {
                        $response_json_object = ConvertFrom-Json $response_json -ErrorAction Stop
                        
                        # エラーレスポンスのチェック
                        if ($response_json_object.error)
                        {
                            throw "WebSocketエラーレスポンス: $($response_json_object.error.message)"
                        }
                        
                        # 受信時には、送信に対する受信だけでなく、ブラウザからのイベント通知も受信する可能性がある
                        # イベント通知の場合、response_jsonにidが含まれないため、idの有無で送信に対する受信なのかイベントの通知なのかを判断する
                        if ($response_json_object.id -eq $this.message_id)
                        {
                            return $response_json
                        }
                    }
                    catch
                    {
                        Write-Host "受信内容をJSON化できませんでした: $response_json"
                        $retry_count++
                        Start-Sleep -Milliseconds 100
                        continue
                    }
                }
                catch
                {
                    $retry_count++
                    Write-Host "WebSocketメッセージ受信エラー（再試行 $retry_count/$max_retries）: $($_.Exception.Message)"
                    Start-Sleep -Milliseconds 500
                }
                finally
                {
                    if ($message_stream)
                    {
                        $message_stream.Dispose()
                    }
                    if ($cancellationTokenSource)
                    {
                        $cancellationTokenSource.Dispose()
                    }
                }
            }
        }
        catch
        {
            # 初期化・接続関連エラー (1006)
            LogWebDriverError $WebDriverErrorCodes.WEBSOCKET_RECEIVE_ERROR "WebSocket メッセージ受信エラー: $($_.Exception.Message)"
            throw "WebSocketメッセージの受信に失敗しました: $($_.Exception.Message)"
        }
        return $response_json
    }

    # ========================================
    # ナビゲーション関連
    # ========================================

    # ページ遷移
    [void] Navigate([string]$url)
    {
        try
        {
            if ([string]::IsNullOrEmpty($url))
            {
                throw "URLが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # URLの形式を検証
            if (-not [System.Uri]::IsWellFormedUriString($url, [System.UriKind]::AbsoluteOrRelative))
            {
                throw "無効なURL形式です: $url"
            }

            # ページ遷移リクエスト送信
            $this.SendWebSocketMessage('Page.navigate', @{ url = $url })
            $response = $this.ReceiveWebSocketMessage()
            
            # レスポンスのエラーチェック
            $response_obj = $response | ConvertFrom-Json
            if ($response_obj.error)
            {
                throw "ページ遷移エラー: $($response_obj.error.message)"
            }

            # ページロード完了待機
            $timeout = [datetime]::Now.AddSeconds(60)
            $retry_count = 0
            $max_retries = 10
            
            try
            {
                while ($true)
                {
                    if ([datetime]::Now -gt $timeout)
                    {
                        throw 'ページロード完了待機タイムアウト'
                    }
                    
                    $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.readyState;"; returnByValue = $true })
                    $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                    
                    if ($response_json.error)
                    {
                        throw "JavaScript実行エラー: $($response_json.error.message)"
                    }
                    
                    if ($response_json.result.result.value -eq 'complete')
                    {
                        break
                    }
                    
                    $retry_count++
                    if ($retry_count -gt $max_retries)
                    {
                        Write-Host "ページロード待機中... 現在の状態: $($response_json.result.result.value)"
                        $retry_count = 0
                    }
                    
                    Start-Sleep -Milliseconds 500
                }
            }
            catch
            {
                throw "ページロード完了待機に失敗しました: $($_.Exception.Message)"
            }

            # ページイベントの有効化
            try
            {
                $this.EnablePageEvents()
            }
            catch
            {
                Write-Host "ページイベントの有効化に失敗しましたが、処理を続行します: $($_.Exception.Message)"
            }

            # 利用可能なタブ情報を取得
            try
            {
                $this.DiscoverTargets()
            }
            catch
            {
                Write-Host "ターゲット発見に失敗しましたが、処理を続行します: $($_.Exception.Message)"
            }

            # JavaScript実行環境を有効化し、`document` オブジェクトにアクセスできるようにする
            try
            {
                $this.SendWebSocketMessage('Runtime.enable', @{ })
                $this.ReceiveWebSocketMessage() | Out-Null
            }
            catch
            {
                Write-Host "Runtime.enableに失敗しましたが、処理を続行します: $($_.Exception.Message)"
            }

            # 現在のページのDOMツリーを取得し、documentノード直下の要素まで取得
            try
            {
                $this.SendWebSocketMessage('DOM.getDocument', @{ depth = 1; })
                $this.ReceiveWebSocketMessage() | Out-Null
            }
            catch
            {
                Write-Host "DOM.getDocumentに失敗しましたが、処理を続行します: $($_.Exception.Message)"
            }

            Write-Host "ページ遷移が完了しました: $url"
        }
        catch
        {
            # ナビゲーション関連エラー (1011)
            LogWebDriverError $WebDriverErrorCodes.NAVIGATION_ERROR "ページ遷移エラー: $($_.Exception.Message)"
            throw "ページ遷移に失敗しました: $($_.Exception.Message)"
        }
    }

    # 広告の読み込み待機
    [void] WaitForAdToLoad()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $timeout = [datetime]::Now.AddSeconds(60)
            $retry_count = 0
            $max_retries = 5
            
            while ([datetime]::Now -lt $timeout)
            {
                try
                {
                    $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.querySelector('.ad, .adsbygoogle, .banner, .popup') !== null"; returnByValue = $true })
                    $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                    
                    if ($response_json.error)
                    {
                        throw "JavaScript実行エラー: $($response_json.error.message)"
                    }
                    
                    $loaded = $response_json.result.result.value
                    if ($loaded -eq $true)
                    {
                        Write-Host "広告が読み込まれました。"
                        return
                    }
                }
                catch
                {
                    $retry_count++
                    if ($retry_count -ge $max_retries)
                    {
                        throw "広告読み込み確認に失敗しました: $($_.Exception.Message)"
                    }
                    Write-Host "広告読み込み確認エラー（再試行 $retry_count/$max_retries）: $($_.Exception.Message)"
                }
                
                Start-Sleep -Milliseconds 500
            }
            throw "広告の読み込みを待機中にタイムアウトしました。"
        }
        catch
        {
            # ナビゲーション関連エラー (1016)
            LogWebDriverError $WebDriverErrorCodes.AD_LOAD_WAIT_ERROR "広告読み込み待機エラー: $($_.Exception.Message)"
            throw "広告の読み込み待機に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウを閉じる
    [void] CloseWindow()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                Write-Host "WebDriverが初期化されていないため、ウィンドウを閉じる処理をスキップします。"
                return
            }

            $this.SendWebSocketMessage('Browser.close', @{ })
            $response = $this.ReceiveWebSocketMessage()
            
            # レスポンスのエラーチェック
            $response_obj = $response | ConvertFrom-Json
            if ($response_obj.error)
            {
                Write-Host "ウィンドウを閉じる際にエラーが発生しましたが、処理を続行します: $($response_obj.error.message)"
            }
            else
            {
                Write-Host "ウィンドウを正常に閉じました。"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1077)
            LogWebDriverError $WebDriverErrorCodes.CLOSE_WINDOW_ERROR "ウィンドウを閉じるエラー: $($_.Exception.Message)"
            Write-Host "ウィンドウを閉じる際にエラーが発生しました: $($_.Exception.Message)"
        }
    }

    # WebSocketを閉じる
    [void] Dispose()
    {
        try
        {
            # WebSocket接続を閉じる
            if ($this.web_socket -and $this.web_socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) 
            {
                try
                {
                    $cancellationTokenSource = [System.Threading.CancellationTokenSource]::new()
                    $cancellationTokenSource.CancelAfter([TimeSpan]::FromSeconds(5))
                    
                    $this.web_socket.CloseAsync('NormalClosure', 'Closing', $cancellationTokenSource.Token).Wait()
                    Write-Host "WebSocket接続を正常に閉じました。"
                }
                catch
                {
                    Write-Host "WebSocket接続を閉じる際にエラーが発生しました: $($_.Exception.Message)"
                }
                finally
                {
                    if ($cancellationTokenSource)
                    {
                        $cancellationTokenSource.Dispose()
                    }
                }
                
                $this.web_socket.Dispose()
                $this.web_socket = $null
            }
            
            # ブラウザプロセスを終了
            if ($this.browser_exe_process_id -gt 0) 
            {
                try
                {
                    $process = Get-Process -Id $this.browser_exe_process_id -ErrorAction SilentlyContinue
                    if ($process)
                    {
                        Stop-Process -Id $this.browser_exe_process_id -Force -ErrorAction Stop
                        Write-Host "ブラウザプロセスを正常に終了しました。プロセスID: $($this.browser_exe_process_id)"
                    }
                    else
                    {
                        Write-Host "ブラウザプロセスは既に終了しています。プロセスID: $($this.browser_exe_process_id)"
                    }
                }
                catch
                {
                    Write-Host "ブラウザプロセスを終了する際にエラーが発生しました: $($_.Exception.Message)"
                }
                finally
                {
                    $this.browser_exe_process_id = 0
                }
            }
            
            $this.is_initialized = $false
            Write-Host "WebDriverのリソースを正常に解放しました。"
        }
        catch
        {
            # 初期化・接続関連エラー (1007)
            LogWebDriverError $WebDriverErrorCodes.DISPOSE_ERROR "Disposeエラー: $($_.Exception.Message)"
            Write-Host "リソースの解放中にエラーが発生しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # タブ・ウィンドウ管理
    # ========================================

    # 新しいタブやページが開かれた際に、そのターゲット情報を自動で検出できるようにする
    # discover = $true でターゲットの発見機能を有効にする
    [string] DiscoverTargets()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Target.setDiscoverTargets', @{ discover = $true })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ターゲット発見エラー: $($response.error.message)"
            }
            
            return $response.result.targetId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1081)
            LogWebDriverError $WebDriverErrorCodes.DISCOVER_TARGETS_ERROR "ターゲット発見エラー: $($_.Exception.Message)"
            throw "ターゲットの発見に失敗しました: $($_.Exception.Message)"
        }
    }
    
    # 現在開かれているすべてのターゲット（タブやページ）の情報を取得
    # ターゲットには、タブのID、URL、タイトル、タイプなどが含まれる
    [hashtable] GetAvailableTabs()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Target.getTargets', @{ })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "タブ情報取得エラー: $($response_json.error.message)"
            }
            
            return $response_json.result
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1082)
            LogWebDriverError $WebDriverErrorCodes.GET_AVAILABLE_TABS_ERROR "タブ情報取得エラー: $($_.Exception.Message)"
            throw "利用可能なタブ情報の取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # タブをアクティブにする
    [void] SetActiveTab($tab_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($tab_id))
            {
                throw "タブIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Target.attachToTarget', @{ targetId = $tab_id })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "タブアクティブ化エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1083)
            LogWebDriverError $WebDriverErrorCodes.SET_ACTIVE_TAB_ERROR "タブアクティブ化エラー: $($_.Exception.Message)"
            throw "タブのアクティブ化に失敗しました: $($_.Exception.Message)"
        }
    }

    # タブを閉じる
    [void] CloseTab($tab_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($tab_id))
            {
                throw "タブIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Target.detachFromTarget', @{ targetId = $tab_id })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "タブ切断エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1084)
            LogWebDriverError $WebDriverErrorCodes.CLOSE_TAB_ERROR "タブ切断エラー: $($_.Exception.Message)"
            throw "タブの切断に失敗しました: $($_.Exception.Message)"
        }
    }

    # ページイベント有効化
    [void] EnablePageEvents()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Page.enable', @{ })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ページイベント有効化エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1085)
            LogWebDriverError $WebDriverErrorCodes.ENABLE_PAGE_EVENTS_ERROR "ページイベント有効化エラー: $($_.Exception.Message)"
            throw "ページイベントの有効化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 要素検索関連
    # ========================================

    # CSSセレクタで要素を検索
    [hashtable] FindElement([string]$selector)
    {
        try
        {
            if ([string]::IsNullOrEmpty($selector))
            {
                throw "CSSセレクタが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.querySelector('$selector')" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            if ($response_json.result.result.objectId -ne $null)
            {
                return @{ nodeId = $response_json.result.result.objectId; selector = $selector }
            }
            else
            {
                throw "CSSセレクタで要素を取得できません。セレクタ：$selector"
            }
        }
        catch
        {
            # 要素検索関連エラー (1021)
            LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENT_ERROR "要素検索エラー (CSS): $($_.Exception.Message)"
            throw "CSSセレクタでの要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # CSSセレクタで複数の要素を検索
    [array] FindElements([string]$selector)
    {
        try
        {
            if ([string]::IsNullOrEmpty($selector))
            {
                throw "CSSセレクタが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "...document.querySelectorAll('$selector').map(e => e.outerHTML)" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            if ($response_json.result.result.objectIds.Count -eq 0)
            {
                return @{ nodeId = $response_json.result.result.objectIds; selector = $selector }
            }
            else
            {
                throw "CSSセレクタで複数の要素を取得できません。セレクタ：$selector"
            }
        }
        catch
        {
            # 要素検索関連エラー (1028)
            LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENTS_ERROR "複数要素検索エラー (CSS): $($_.Exception.Message)"
            throw "CSSセレクタでの複数要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # JavaScriptで要素を検索
    [hashtable] FindElementGeneric([string]$expression, [string]$query_type, [string]$element)
    {
        try
        {
            if ([string]::IsNullOrEmpty($expression))
            {
                throw "JavaScript式が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            if ($response_json.result.result.objectId)
            {
                return @{ nodeId = $response_json.result.result.objectId; query_type = $query_type; element = $element } 
            }
            else
            {
                throw "$query_type による要素取得に失敗しました。element：$element"
            }
        }
        catch
        {
            # 要素検索関連エラー (1022)
            LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENT_GENERIC_ERROR "要素検索エラー ($query_type): $($_.Exception.Message)"
            throw "$query_type による要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # JavaScriptで複数の要素を検索
    [hashtable] FindElementsGeneric([string]$expression, [string]$query_type, [string]$element)
    {
        try
        {
            if ([string]::IsNullOrEmpty($expression))
            {
                throw "JavaScript式が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            if ($response_json.result.result.value -gt 0)
            {
                return @{ count = $response_json.result.result.value; query_type = $query_type; element = $element }
            }
            else
            {
                throw "$query_type による複数の要素取得に失敗しました。element：$element"
            }
        }
        catch
        {
            # 要素検索関連エラー (1029)
            LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENTS_GENERIC_ERROR "複数要素検索エラー ($query_type): $($_.Exception.Message)"
            throw "$query_type による複数要素検索に失敗しました: $($_.Exception.Message)"
        }
    }
    
    <#
    # XPathで複数の要素を検索
    [array] FindElementsByXPath([string]$xpath)
    {
        $expression = "...document.evaluate('$xpath', document, null, XPathResult.ANY_TYPE, null).iterateNext().outerHTML"
        return $this.FindElementsGeneric($expression, 'XPath', $xpath)
    }
    #>

    # class属性で複数の要素を検索
    [array] FindElementsByClassName([string]$class_name)
    {
        try
        {
            if ([string]::IsNullOrEmpty($class_name))
            {
                throw "クラス名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.getElementsByClassName('$class_name').length"
            $element_count = $this.FindElementsGeneric($expression, 'ClassName', $class_name).count

            $element_list = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $element_count; $i++)
            {
                try
                {
                    $element_list.add($this.FindElementByClassName($class_name, $i))
                }
                catch
                {
                    Write-Host "インデックス $i の要素取得に失敗しましたが、処理を続行します: $($_.Exception.Message)"
                }
            }
            return $element_list
        }
        catch
        {
            # 要素検索関連エラー (1030)
            LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENTS_BY_CLASSNAME_ERROR "複数要素検索エラー (ClassName): $($_.Exception.Message)"
            throw "ClassNameでの複数要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # name属性で複数の要素を検索
    [array] FindElementsByName([string]$name)
    {
        try
        {
            if ([string]::IsNullOrEmpty($name))
            {
                throw "name属性が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.getElementsByName('$name').length"
            $element_count = $this.FindElementsGeneric($expression, 'Name', $name).count

            $element_list = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $element_count; $i++)
            {
                try
                {
                    $element_list.add($this.FindElementByName($name, $i))
                }
                catch
                {
                    Write-Host "インデックス $i の要素取得に失敗しましたが、処理を続行します: $($_.Exception.Message)"
                }
            }
            return $element_list
        }
        catch
        {
            # 要素検索関連エラー (1031)
            LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENTS_BY_NAME_ERROR "複数要素検索エラー (Name): $($_.Exception.Message)"
            throw "Nameでの複数要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # tag名で複数の要素を検索
    [array] FindElementsByTagName([string]$tag_name)
    {
        try
        {
            if ([string]::IsNullOrEmpty($tag_name))
            {
                throw "タグ名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.getElementsByTagName('$tag_name').length"
            $element_count = $this.FindElementsGeneric($expression, 'TagName', $tag_name).count        
            
            $element_list = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $element_count; $i++)
            {
                try
                {
                    $element_list.add($this.FindElementByTagName($tag_name, $i))
                }
                catch
                {
                    Write-Host "インデックス $i の要素取得に失敗しましたが、処理を続行します: $($_.Exception.Message)"
                }
            }
            return $element_list
        }
        catch
        {
            # 要素検索関連エラー (1032)
            LogWebDriverError $WebDriverErrorCodes.FIND_ELEMENTS_BY_TAGNAME_ERROR "複数要素検索エラー (TagName): $($_.Exception.Message)"
            throw "TagNameでの複数要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # JavaScriptで要素有無を検索
    [bool] IsExistsElementGeneric([string]$expression, [string]$query_type, [string]$element)
    {
        try
        {
            if ([string]::IsNullOrEmpty($expression))
            {
                throw "JavaScript式が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            if ($response_json.result.result.objectId)
            {
                return $true
            }
            else
            {
                return $false
            }
        }
        catch
        {
            # 要素検索関連エラー (1033)
            LogWebDriverError $WebDriverErrorCodes.IS_EXISTS_ELEMENT_GENERIC_ERROR "要素存在確認エラー ($query_type): $($_.Exception.Message)"
            throw "$query_type による要素存在確認に失敗しました: $($_.Exception.Message)"
        }
    }

    # XPathで要素有無を検索
    [bool] IsExistsElementByXPath([string]$xpath)
    {
        try
        {
            if ([string]::IsNullOrEmpty($xpath))
            {
                throw "XPathが指定されていません。"
            }

            $expression = "document.evaluate('$xpath', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
            return $this.IsExistsElementGeneric($expression, 'XPath', $xpath)
        }
        catch
        {
            # 要素検索関連エラー (1034)
            LogWebDriverError $WebDriverErrorCodes.IS_EXISTS_ELEMENT_XPATH_ERROR "XPath要素存在確認エラー: $($_.Exception.Message)"
            throw "XPathでの要素存在確認に失敗しました: $($_.Exception.Message)"
        }
    }

    # class属性で要素有無を検索
    [bool] IsExistsElementByClassName([string]$class_name, [int]$index)
    {
        try
        {
            if ([string]::IsNullOrEmpty($class_name))
            {
                throw "クラス名が指定されていません。"
            }

            if ($index -lt 0)
            {
                throw "インデックスは0以上である必要があります。"
            }

            $expression = "document.getElementsByClassName('$class_name')[$index]"
            return $this.IsExistsElementGeneric($expression, 'ClassName', $class_name)
        }
        catch
        {
            # 要素検索関連エラー (1035)
            LogWebDriverError $WebDriverErrorCodes.IS_EXISTS_ELEMENT_CLASSNAME_ERROR "ClassName要素存在確認エラー: $($_.Exception.Message)"
            throw "ClassNameでの要素存在確認に失敗しました: $($_.Exception.Message)"
        }
    }

    # id属性で要素有無を検索
    [bool] IsExistsElementById([string]$id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($id))
            {
                throw "IDが指定されていません。"
            }

            $expression = "document.getElementById('$id')"
            return $this.IsExistsElementGeneric($expression, 'Id', $id)
        }
        catch
        {
            # 要素検索関連エラー (1036)
            LogWebDriverError $WebDriverErrorCodes.IS_EXISTS_ELEMENT_ID_ERROR "Id要素存在確認エラー: $($_.Exception.Message)"
            throw "Idでの要素存在確認に失敗しました: $($_.Exception.Message)"
        }
    }

    # name属性で要素有無を検索
    [bool] IsExistsElementByName([string]$name, [int]$index)
    {
        try
        {
            if ([string]::IsNullOrEmpty($name))
            {
                throw "name属性が指定されていません。"
            }

            if ($index -lt 0)
            {
                throw "インデックスは0以上である必要があります。"
            }

            $expression = "document.getElementsByName('$name')[$index]"
            return $this.IsExistsElementGeneric($expression, 'Name', $name)
        }
        catch
        {
            # 要素検索関連エラー (1037)
            LogWebDriverError $WebDriverErrorCodes.IS_EXISTS_ELEMENT_NAME_ERROR "Name要素存在確認エラー: $($_.Exception.Message)"
            throw "Nameでの要素存在確認に失敗しました: $($_.Exception.Message)"
        }
    }

    # tag名で要素有無を検索
    [bool] IsExistsElementByTagName([string]$tag_name, [int]$index)
    {
        try
        {
            if ([string]::IsNullOrEmpty($tag_name))
            {
                throw "タグ名が指定されていません。"
            }

            if ($index -lt 0)
            {
                throw "インデックスは0以上である必要があります。"
            }

            $expression = "document.getElementsByTagName('$tag_name')[$index]"
            return $this.IsExistsElementGeneric($expression, 'TagName', $tag_name)
        }
        catch
        {
            # 要素検索関連エラー (1038)
            LogWebDriverError $WebDriverErrorCodes.IS_EXISTS_ELEMENT_TAGNAME_ERROR "TagName要素存在確認エラー: $($_.Exception.Message)"
            throw "TagNameでの要素存在確認に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 要素操作関連
    # ========================================

    # 要素のテキストを取得
    [string] GetElementText([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { return this.textContent; }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1051)
            LogWebDriverError $WebDriverErrorCodes.GET_ELEMENT_TEXT_ERROR "要素テキスト取得エラー: $($_.Exception.Message)"
            throw "要素のテキスト取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素にテキストを入力
    [void] SetElementText([string]$object_id, [string]$text)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(value) { this.value = value; this.dispatchEvent(new Event('input')); }"; arguments = @(@{ value = $text }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1052)
            LogWebDriverError $WebDriverErrorCodes.SET_ELEMENT_TEXT_ERROR "要素テキスト設定エラー: $($_.Exception.Message)"
            throw "要素へのテキスト入力に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素をクリック
    [void] ClickElement([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { this.click(); }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1054)
            LogWebDriverError $WebDriverErrorCodes.CLICK_ELEMENT_ERROR "要素クリックエラー: $($_.Exception.Message)"
            throw "要素のクリックに失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素の属性を取得
    [string] GetElementAttribute([string]$object_id, [string]$attribute)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($attribute))
            {
                throw "属性名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(name) { return this.getAttribute(name); }"; arguments = @(@{ value = $attribute }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1060)
            LogWebDriverError $WebDriverErrorCodes.GET_ELEMENT_ATTRIBUTE_ERROR "要素属性取得エラー: $($_.Exception.Message)"
            throw "要素の属性取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素の属性を設定
    [void] SetElementAttribute([string]$object_id, [string]$attribute, [string]$value)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($attribute))
            {
                throw "属性名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(name, value) { this.setAttribute(name, value); }"; arguments = @(@{ value = $attribute }, @{ value = $value }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1061)
            LogWebDriverError $WebDriverErrorCodes.SET_ELEMENT_ATTRIBUTE_ERROR "要素属性設定エラー: $($_.Exception.Message)"
            throw "要素の属性設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # アンカータグのhrefを取得
    [string] GetHrefFromAnchor([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $functionDeclaration = "function() { return this.href; }"
            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = $functionDeclaration; returnByValue = $true })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "JavaScript実行エラー: $($response.error.message)"
            }
            
            return $response.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1064)
            LogWebDriverError $WebDriverErrorCodes.GET_HREF_ERROR "href取得エラー: $($_.Exception.Message)"
            throw "hrefの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素のCSSプロパティを取得
    [string] GetElementCssProperty([string]$object_id, [string]$property)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($property))
            {
                throw "CSSプロパティ名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(name) { return window.getComputedStyle(this).getPropertyValue(name); }"; arguments = @(@{ value = $property }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1062)
            LogWebDriverError $WebDriverErrorCodes.GET_ELEMENT_CSS_PROPERTY_ERROR "CSSプロパティ取得エラー: $($_.Exception.Message)"
            throw "要素のCSSプロパティ取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素のCSSプロパティを設定
    [void] SetElementCssProperty([string]$object_id, [string]$property, [string]$value)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($property))
            {
                throw "CSSプロパティ名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(name, value) { this.style.setProperty(name, value); }"; arguments = @(@{ value = $property }, @{ value = $value }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1063)
            LogWebDriverError $WebDriverErrorCodes.SET_ELEMENT_CSS_PROPERTY_ERROR "CSSプロパティ設定エラー: $($_.Exception.Message)"
            throw "要素のCSSプロパティ設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # フォーム操作関連
    # ========================================

    # セレクトタグのオプションをインデックス番号から選択
    [void] SelectOptionByIndex([string]$object_id, [int]$index)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ($index -lt 0)
            {
                throw "インデックスは0以上である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(index) { this.selectedIndex = index; this.dispatchEvent(new Event('change')); }"; arguments = @(@{ value = $index }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1065)
            LogWebDriverError $WebDriverErrorCodes.SELECT_OPTION_BY_INDEX_ERROR "オプション選択エラー (インデックス): $($_.Exception.Message)"
            throw "インデックスによるオプション選択に失敗しました: $($_.Exception.Message)"
        }
    }

    # セレクトタグのオプションをテキストを指定して選択
    [void] SelectOptionByText([string]$object_id, [string]$text)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($text))
            {
                throw "選択するテキストが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(text) { for (let option of this.options) { if (option.text === text) { option.selected = true; this.dispatchEvent(new Event('change')); break; } } }"; arguments = @(@{ value = $text }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1066)
            LogWebDriverError $WebDriverErrorCodes.SELECT_OPTION_BY_TEXT_ERROR "オプション選択エラー (テキスト): $($_.Exception.Message)"
            throw "テキストによるオプション選択に失敗しました: $($_.Exception.Message)"
        }
    }

    # セレクトタグを未選択にする
    [void] DeselectAllOptions([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { for (let option of this.options) { option.selected = false; } this.dispatchEvent(new Event('change')); }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1067)
            LogWebDriverError $WebDriverErrorCodes.DESELECT_ALL_OPTIONS_ERROR "オプション未選択エラー: $($_.Exception.Message)"
            throw "オプションの未選択に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素に入力された値をクリア
    [void] ClearElement([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { this.value = ''; }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1053)
            LogWebDriverError $WebDriverErrorCodes.CLEAR_ELEMENT_ERROR "要素クリアエラー: $($_.Exception.Message)"
            throw "要素のクリアに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ウィンドウ操作関連
    # ========================================

    # ウィンドウサイズの変更
    [void] ResizeWindow([int]$width, [int]$height, [int]$windowHandle)
    {
        try
        {
            if ($width -le 0 -or $height -le 0)
            {
                throw "ウィンドウサイズは正の値である必要があります。"
            }

            if ($windowHandle -le 0)
            {
                throw "ウィンドウハンドルは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{width = $width; height = $height} })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ウィンドウリサイズエラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1071)
            LogWebDriverError $WebDriverErrorCodes.RESIZE_WINDOW_ERROR "ウィンドウリサイズエラー: $($_.Exception.Message)"
            throw "ウィンドウのリサイズに失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウサイズの変更(通常)
    [void] NormalWindow([int]$windowHandle)
    {
        try
        {
            if ($windowHandle -le 0)
            {
                throw "ウィンドウハンドルは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'normal'} })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ウィンドウ状態変更エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1072)
            LogWebDriverError $WebDriverErrorCodes.NORMAL_WINDOW_ERROR "ウィンドウ状態変更エラー (通常): $($_.Exception.Message)"
            throw "ウィンドウの通常状態への変更に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウサイズの変更(最大化)
    [void] MaximizeWindow([int]$windowHandle)
    {
        try
        {
            if ($windowHandle -le 0)
            {
                throw "ウィンドウハンドルは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'maximized'} })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ウィンドウ状態変更エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1073)
            LogWebDriverError $WebDriverErrorCodes.MAXIMIZE_WINDOW_ERROR "ウィンドウ状態変更エラー (最大化): $($_.Exception.Message)"
            throw "ウィンドウの最大化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウサイズの変更(最小化)
    [void] MinimizeWindow([int]$windowHandle)
    {
        try
        {
            if ($windowHandle -le 0)
            {
                throw "ウィンドウハンドルは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'minimized'} })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ウィンドウ状態変更エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1074)
            LogWebDriverError $WebDriverErrorCodes.MINIMIZE_WINDOW_ERROR "ウィンドウ状態変更エラー (最小化): $($_.Exception.Message)"
            throw "ウィンドウの最小化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウサイズの変更(フルスクリーン)
    [void] FullscreenWindow([int]$windowHandle)
    {
        try
        {
            if ($windowHandle -le 0)
            {
                throw "ウィンドウハンドルは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'fullscreen'} })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ウィンドウ状態変更エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1075)
            LogWebDriverError $WebDriverErrorCodes.FULLSCREEN_WINDOW_ERROR "ウィンドウ状態変更エラー (フルスクリーン): $($_.Exception.Message)"
            throw "ウィンドウのフルスクリーン化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウ位置の変更
    [void] MoveWindow([int]$x, [int]$y, [int]$windowHandle)
    {
        try
        {
            if ($windowHandle -le 0)
            {
                throw "ウィンドウハンドルは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{left = $x; top = $y} })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ウィンドウ移動エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1076)
            LogWebDriverError $WebDriverErrorCodes.MOVE_WINDOW_ERROR "ウィンドウ移動エラー: $($_.Exception.Message)"
            throw "ウィンドウの移動に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ナビゲーション関連（履歴操作）
    # ========================================

    # ブラウザの履歴中の前のページに戻る
    [void] GoBack()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Page.navigateToHistoryEntry', @{ entryId = -1 })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ブラウザ履歴移動エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ナビゲーション関連エラー (1012)
            LogWebDriverError $WebDriverErrorCodes.GO_BACK_ERROR "ブラウザ履歴移動エラー (戻る): $($_.Exception.Message)"
            throw "ブラウザの履歴を戻る操作に失敗しました: $($_.Exception.Message)"
        }
    }

    # ブラウザの履歴中の次のページに進む
    [void] GoForward()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Page.navigateToHistoryEntry', @{ entryId = 1 })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ブラウザ履歴移動エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ナビゲーション関連エラー (1013)
            LogWebDriverError $WebDriverErrorCodes.GO_FORWARD_ERROR "ブラウザ履歴移動エラー (進む): $($_.Exception.Message)"
            throw "ブラウザの履歴を進む操作に失敗しました: $($_.Exception.Message)"
        }
    }

    # ブラウザを更新
    [void] Refresh()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Page.reload', @{ ignoreCache = $true })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ブラウザ更新エラー: $($response.error.message)"
            }
        }
        catch
        {
            # ナビゲーション関連エラー (1014)
            LogWebDriverError $WebDriverErrorCodes.REFRESH_ERROR "ブラウザ更新エラー: $($_.Exception.Message)"
            throw "ブラウザの更新に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 情報取得関連
    # ========================================

    # URLを取得
    [string] GetUrl()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "URL取得エラー: $($response_json.error.message)"
            }
            
            if (-not $response_json.result.entries -or $response_json.result.entries.Count -eq 0)
            {
                throw "ナビゲーション履歴が空です。"
            }
            
            return $response_json.result.entries[0].url
        }
        catch
        {
            # 情報取得関連エラー (1091)
            LogWebDriverError $WebDriverErrorCodes.GET_URL_ERROR "URL取得エラー: $($_.Exception.Message)"
            throw "URLの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # タイトルを取得
    [string] GetTitle()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "タイトル取得エラー: $($response_json.error.message)"
            }
            
            if (-not $response_json.result.entries -or $response_json.result.entries.Count -eq 0)
            {
                throw "ナビゲーション履歴が空です。"
            }
            
            return $response_json.result.entries[0].title
        }
        catch
        {
            # 情報取得関連エラー (1092)
            LogWebDriverError $WebDriverErrorCodes.GET_TITLE_ERROR "タイトル取得エラー: $($_.Exception.Message)"
            throw "タイトルの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ソースコードを取得
    [string] GetSourceCode()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.documentElement.outerHTML" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "ソースコード取得エラー: $($response_json.error.message)"
            }
            
            return $response_json.result.result.value
        }
        catch
        {
            # 情報取得関連エラー (1093)
            LogWebDriverError $WebDriverErrorCodes.GET_SOURCE_CODE_ERROR "ソースコード取得エラー: $($_.Exception.Message)"
            throw "ソースコードの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウハンドルを取得
    [int] GetWindowHandle()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "ウィンドウハンドル取得エラー: $($response_json.error.message)"
            }
            
            if (-not $response_json.result.entries -or $response_json.result.entries.Count -eq 0)
            {
                throw "ナビゲーション履歴が空です。"
            }
            
            return $response_json.result.entries[0].windowId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1078)
            LogWebDriverError $WebDriverErrorCodes.GET_WINDOW_HANDLE_ERROR "ウィンドウハンドル取得エラー: $($_.Exception.Message)"
            throw "ウィンドウハンドルの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # 複数のウィンドウハンドルを取得
    [array] GetWindowHandles()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Target.getTargets', @{ })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "ウィンドウハンドル取得エラー: $($response_json.error.message)"
            }
            
            $window_handles = @()
            foreach ($target in $response_json.result.targetInfos)
            {
                if ($target.type -eq 'page')
                {
                    $window_handles += $target.targetId
                }
            }
            
            return $window_handles
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1079)
            LogWebDriverError $WebDriverErrorCodes.GET_WINDOW_HANDLES_ERROR "複数ウィンドウハンドル取得エラー: $($_.Exception.Message)"
            throw "複数のウィンドウハンドルの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウサイズを取得
    [hashtable] GetWindowSize()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "({width: window.innerWidth, height: window.innerHeight})" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "ウィンドウサイズ取得エラー: $($response_json.error.message)"
            }
            
            return @{
                width = $response_json.result.result.value.width
                height = $response_json.result.result.value.height
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1080)
            LogWebDriverError $WebDriverErrorCodes.GET_WINDOW_SIZE_ERROR "ウィンドウサイズ取得エラー: $($_.Exception.Message)"
            throw "ウィンドウサイズの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # スクリーンショット関連
    # ========================================

    # スクリーンショットを取得（フルページスクリーンショット、ビューポートスクリーンショット）
    [void] GetScreenshot([string]$type, [string]$save_path)
    {
        try
        {
            if ([string]::IsNullOrEmpty($type))
            {
                throw "スクリーンショットのタイプが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($save_path))
            {
                throw "保存パスが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # 保存先ディレクトリの確認と作成
            $save_dir = Split-Path $save_path -Parent
            if (-not (Test-Path $save_dir))
            {
                try
                {
                    New-Item -ItemType Directory -Path $save_dir -Force | Out-Null
                }
                catch
                {
                    throw "保存先ディレクトリの作成に失敗しました: $save_dir"
                }
            }

            switch ($type)
            {
                'fullPage'
                {
                    # ページ全体のスクリーンショット
                    try
                    {
                        $this.SendWebSocketMessage('Page.captureScreenshot', @{ format = 'png'; quality = 100; captureBeyondViewport = $true })
                        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

                        if ($response_json.error)
                        {
                            throw "スクリーンショット取得エラー: $($response_json.error.message)"
                        }

                        # 受信したBase64エンコードされた画像データをBase64デコードして保存
                        $image_data_base64 = $response_json.result.data

                        # Base64エンコードされた画像データをバイト配列に変換
                        $image_data_bytes = [System.Convert]::FromBase64String($image_data_base64)

                        # バイト配列をファイルに書き込む
                        [System.IO.File]::WriteAllBytes($save_path, $image_data_bytes)
                        
                        Write-Host "フルページスクリーンショットを保存しました: $save_path"
                    }
                    catch
                    {
                        throw "フルページスクリーンショットの取得に失敗しました: $($_.Exception.Message)"
                    }
                }
                'viewPort'
                {
                    try
                    {
                        $this.SendWebSocketMessage('Page.captureScreenshot', @{ format = 'png'; quality = 100; captureBeyondViewport = $false })
                        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

                        if ($response_json.error)
                        {
                            throw "スクリーンショット取得エラー: $($response_json.error.message)"
                        }

                        # 受信したBase64エンコードされた画像データをBase64デコードして保存
                        $image_data_base64 = $response_json.result.data

                        # Base64エンコードされた画像データをバイト配列に変換
                        $image_data_bytes = [System.Convert]::FromBase64String($image_data_base64)

                        # バイト配列をファイルに書き込む
                        [System.IO.File]::WriteAllBytes($save_path, $image_data_bytes)
                        
                        Write-Host "ビューポートスクリーンショットを保存しました: $save_path"
                    }
                    catch
                    {
                        throw "ビューポートスクリーンショットの取得に失敗しました: $($_.Exception.Message)"
                    }
                }
                default
                {
                    throw "サポートされていないスクリーンショットタイプです: $type (サポート: fullPage, viewPort)"
                }
            }
        }
        catch
        {
            # スクリーンショット関連エラー (1101)
            LogWebDriverError $WebDriverErrorCodes.SCREENSHOT_ERROR "スクリーンショット取得エラー: $($_.Exception.Message)"
            throw "スクリーンショットの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # スクリーンショットを取得（指定要素のスクリーンショット）
    [void] GetScreenshotObjectId([string]$object_id, [string]$save_path)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($save_path))
            {
                throw "保存パスが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # 保存先ディレクトリの確認と作成
            $save_dir = Split-Path $save_path -Parent
            if (-not (Test-Path $save_dir))
            {
                try
                {
                    New-Item -ItemType Directory -Path $save_dir -Force | Out-Null
                }
                catch
                {
                    throw "保存先ディレクトリの作成に失敗しました: $save_dir"
                }
            }

            # 要素の位置とサイズを取得
            $this.SendWebSocketMessage('DOM.getBoxModel', @{ objectId = $object_id })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

            if ($response.error)
            {
                throw "要素の位置とサイズ取得エラー: $($response.error.message)"
            }

            if (-not $response.result)
            {
                throw '要素の位置とサイズを取得できません。'
            }

            $content_box = $response.result.model.content

            # 座標をグループ化（x, yペアで分ける）
            $points = @()
            for ($i = 0; $i -lt $content_box.count; $i += 2)
            {
                $points += ,@($content_box[$i], $content_box[$i + 1])
            }

            # 最小x/y、最大x/yを使って矩形を算出（4点でも8点でも対応）
            $x_list = $points | ForEach-Object { $_[0] }
            $y_list = $points | ForEach-Object { $_[1] }

            $x_min = [Math]::Floor(($x_list | Measure-Object -Minimum).Minimum)
            $y_min = [Math]::Floor(($y_list | Measure-Object -Minimum).Minimum)
            $x_max = [Math]::Ceiling(($x_list | Measure-Object -Maximum).Maximum)
            $y_max = [Math]::Ceiling(($y_list | Measure-Object -Maximum).Maximum)

            $width  = $x_max - $x_min
            $height = $y_max - $y_min

            # スクリーンショットをClip付きで取得
            $this.SendWebSocketMessage('Page.captureScreenshot', @{ format = 'png'; quality = 100; clip = @{ x = $x_min; y = $y_min; width = $width; height = $height; scale = 1 }; captureBeyondViewport = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

            if ($response_json.error)
            {
                throw "スクリーンショット取得エラー: $($response_json.error.message)"
            }

            # 受信したBase64エンコードされた画像データをBase64デコードして保存
            $image_data_base64 = $response_json.result.data

            # Base64エンコードされた画像データをバイト配列に変換
            $image_data_bytes = [System.Convert]::FromBase64String($image_data_base64)

            # バイト配列をファイルに書き込む
            [System.IO.File]::WriteAllBytes($save_path, $image_data_bytes)
            
            Write-Host "要素スクリーンショットを保存しました: $save_path"
        }
        catch
        {
            # スクリーンショット関連エラー (1102)
            LogWebDriverError $WebDriverErrorCodes.SCREENSHOT_OBJECT_ERROR "要素スクリーンショット取得エラー: $($_.Exception.Message)"
            throw "要素のスクリーンショット取得に失敗しました: $($_.Exception.Message)"
        }

    }

    # スクリーンショットを取得（指定要素のスクリーンショット）
    [void] GetScreenshotObjectIds([string[]]$object_ids, [string]$save_path)
    {
        try
        {
            if ($object_ids.Count -eq 0)
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($save_path))
            {
                throw "保存パスが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # 保存先ディレクトリの確認と作成
            $save_dir = Split-Path $save_path -Parent
            if (-not (Test-Path $save_dir))
            {
                try
                {
                    New-Item -ItemType Directory -Path $save_dir -Force | Out-Null
                }
                catch
                {
                    throw "保存先ディレクトリの作成に失敗しました: $save_dir"
                }
            }

            $x_list = @()
            $y_list = @()

            foreach ($object_id in $object_ids)
            {
                # 要素の位置とサイズを取得
                $this.SendWebSocketMessage('DOM.getBoxModel', @{ objectId = $object_id })
                $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

                if (-not $response.result.model.content)
                {
                    throw "ボックスモデルの取得に失敗しました（objectId: $object_id）"
                }

                $content_box = $response.result.model.content
                for ($i = 0; $i -lt $content_box.Count; $i += 2)
                {
                    $x_list += $content_box[$i]
                    $y_list += $content_box[$i + 1]
                }
            }

            # 全要素の合成矩形を計算
            $x_min = [Math]::Floor(($x_list | Measure-Object -Minimum).Minimum)
            $y_min = [Math]::Floor(($y_list | Measure-Object -Minimum).Minimum)
            $x_max = [Math]::Ceiling(($x_list | Measure-Object -Maximum).Maximum)
            $y_max = [Math]::Ceiling(($y_list | Measure-Object -Maximum).Maximum)

            $width  = $x_max - $x_min
            $height = $y_max - $y_min

            # スクリーンショットをClip付きで取得
            $this.SendWebSocketMessage('Page.captureScreenshot', @{ format = 'png'; quality = 100; clip = @{ x = $x_min; y = $y_min; width = $width; height = $height; scale = 1 }; captureBeyondViewport = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

            if ($response_json.error)
            {
                throw "スクリーンショット取得エラー: $($response_json.error.message)"
            }

            # 受信したBase64エンコードされた画像データをBase64デコードして保存
            $image_data_base64 = $response_json.result.data

            # Base64エンコードされた画像データをバイト配列に変換
            $image_data_bytes = [System.Convert]::FromBase64String($image_data_base64)

            # バイト配列をファイルに書き込む
            [System.IO.File]::WriteAllBytes($save_path, $image_data_bytes)

            Write-Host "複数要素スクリーンショットを保存しました: $save_path"
        }
        catch
        {
            # スクリーンショット関連エラー (1103)
            LogWebDriverError $WebDriverErrorCodes.SCREENSHOT_OBJECTS_ERROR "複数要素スクリーンショット取得エラー: $($_.Exception.Message)"
            throw "複数要素のスクリーンショット取得に失敗しました: $($_.Exception.Message)"
        }
    }
        
    # ========================================
    # 待機機能
    # ========================================

    # 要素が表示されるまで待機
    [void] WaitForElementVisible([string]$selector, [int]$timeout_seconds = 30)
    {
        try
        {
            if ([string]::IsNullOrEmpty($selector))
            {
                throw "CSSセレクタが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $timeout = [datetime]::Now.AddSeconds($timeout_seconds)
            
            while ([datetime]::Now -lt $timeout)
            {
                try
                {
                    $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.querySelector('$selector') && document.querySelector('$selector').offsetParent !== null"; returnByValue = $true })
                    $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                    
                    if ($response_json.error)
                    {
                        throw "JavaScript実行エラー: $($response_json.error.message)"
                    }
                    
                    if ($response_json.result.result.value -eq $true)
                    {
                        Write-Host "要素が表示されました: $selector"
                        return
                    }
                }
                catch
                {
                    Write-Host "要素表示確認エラー: $($_.Exception.Message)"
                }
                
                Start-Sleep -Milliseconds 500
            }
            
            throw "要素の表示を待機中にタイムアウトしました: $selector"
        }
        catch
        {
            # 要素検索関連エラー (1039)
            LogWebDriverError $WebDriverErrorCodes.WAIT_FOR_ELEMENT_VISIBLE_ERROR "要素表示待機エラー: $($_.Exception.Message)"
            throw "要素の表示待機に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素がクリック可能になるまで待機
    [void] WaitForElementClickable([string]$selector, [int]$timeout_seconds = 30)
    {
        try
        {
            if ([string]::IsNullOrEmpty($selector))
            {
                throw "CSSセレクタが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $timeout = [datetime]::Now.AddSeconds($timeout_seconds)
            
            while ([datetime]::Now -lt $timeout)
            {
                try
                {
                    $expression = @"
                    (function() {
                        const element = document.querySelector('$selector');
                        if (!element) return false;
                        if (element.offsetParent === null) return false;
                        if (element.disabled) return false;
                        if (getComputedStyle(element).pointerEvents === 'none') return false;
                        return true;
                    })()
"@
                    $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression; returnByValue = $true })
                    $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                    
                    if ($response_json.error)
                    {
                        throw "JavaScript実行エラー: $($response_json.error.message)"
                    }
                    
                    if ($response_json.result.result.value -eq $true)
                    {
                        Write-Host "要素がクリック可能になりました: $selector"
                        return
                    }
                }
                catch
                {
                    Write-Host "要素クリック可能性確認エラー: $($_.Exception.Message)"
                }
                
                Start-Sleep -Milliseconds 500
            }
            
            throw "要素のクリック可能性を待機中にタイムアウトしました: $selector"
        }
        catch
        {
            # 要素検索関連エラー (1040)
            LogWebDriverError $WebDriverErrorCodes.WAIT_FOR_ELEMENT_CLICKABLE_ERROR "要素クリック可能性待機エラー: $($_.Exception.Message)"
            throw "要素のクリック可能性待機に失敗しました: $($_.Exception.Message)"
        }
    }

    # ページロード完了まで待機
    [void] WaitForPageLoad([int]$timeout_seconds = 60)
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $timeout = [datetime]::Now.AddSeconds($timeout_seconds)
            
            while ([datetime]::Now -lt $timeout)
            {
                try
                {
                    $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.readyState"; returnByValue = $true })
                    $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                    
                    if ($response_json.error)
                    {
                        throw "JavaScript実行エラー: $($response_json.error.message)"
                    }
                    
                    if ($response_json.result.result.value -eq 'complete')
                    {
                        Write-Host "ページロードが完了しました"
                        return
                    }
                }
                catch
                {
                    Write-Host "ページロード状態確認エラー: $($_.Exception.Message)"
                }
                
                Start-Sleep -Milliseconds 500
            }
            
            throw "ページロード完了を待機中にタイムアウトしました"
        }
        catch
        {
            # ナビゲーション関連エラー (1015)
            LogWebDriverError $WebDriverErrorCodes.WAIT_FOR_PAGE_LOAD_ERROR "ページロード待機エラー: $($_.Exception.Message)"
            throw "ページロードの待機に失敗しました: $($_.Exception.Message)"
        }
    }

    # カスタム条件で待機
    [void] WaitForCondition([string]$javascript_condition, [int]$timeout_seconds = 30)
    {
        try
        {
            if ([string]::IsNullOrEmpty($javascript_condition))
            {
                throw "JavaScript条件が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $timeout = [datetime]::Now.AddSeconds($timeout_seconds)
            
            while ([datetime]::Now -lt $timeout)
            {
                try
                {
                    $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $javascript_condition; returnByValue = $true })
                    $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                    
                    if ($response_json.error)
                    {
                        throw "JavaScript実行エラー: $($response_json.error.message)"
                    }
                    
                    if ($response_json.result.result.value -eq $true)
                    {
                        Write-Host "カスタム条件が満たされました: $javascript_condition"
                        return
                    }
                }
                catch
                {
                    Write-Host "カスタム条件確認エラー: $($_.Exception.Message)"
                }
                
                Start-Sleep -Milliseconds 500
            }
            
            throw "カスタム条件の待機中にタイムアウトしました: $javascript_condition"
        }
        catch
        {
            # 要素検索関連エラー (1041)
            LogWebDriverError $WebDriverErrorCodes.WAIT_FOR_CONDITION_ERROR "カスタム条件待機エラー: $($_.Exception.Message)"
            throw "カスタム条件の待機に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # キーボード操作機能
    # ========================================

    # キーボード入力
    [void] SendKeys([string]$object_id, [string]$keys)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($keys))
            {
                throw "入力するキーが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(keys) { this.focus(); this.value = keys; this.dispatchEvent(new Event('input')); this.dispatchEvent(new Event('change')); }"; arguments = @(@{ value = $keys }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1058)
            LogWebDriverError $WebDriverErrorCodes.SEND_KEYS_ERROR "キーボード入力エラー: $($_.Exception.Message)"
            throw "キーボード入力に失敗しました: $($_.Exception.Message)"
        }
    }

    # 特殊キーの送信
    [void] SendSpecialKey([string]$object_id, [string]$special_key)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($special_key))
            {
                throw "特殊キーが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $key_map = @{
                'Enter' = 'Enter'
                'Tab' = 'Tab'
                'Escape' = 'Escape'
                'Backspace' = 'Backspace'
                'Delete' = 'Delete'
                'ArrowUp' = 'ArrowUp'
                'ArrowDown' = 'ArrowDown'
                'ArrowLeft' = 'ArrowLeft'
                'ArrowRight' = 'ArrowRight'
                'Home' = 'Home'
                'End' = 'End'
                'PageUp' = 'PageUp'
                'PageDown' = 'PageDown'
                'F1' = 'F1'
                'F2' = 'F2'
                'F3' = 'F3'
                'F4' = 'F4'
                'F5' = 'F5'
                'F6' = 'F6'
                'F7' = 'F7'
                'F8' = 'F8'
                'F9' = 'F9'
                'F10' = 'F10'
                'F11' = 'F11'
                'F12' = 'F12'
            }

            if (-not $key_map.ContainsKey($special_key))
            {
                throw "サポートされていない特殊キーです: $special_key"
            }

            $key_code = $key_map[$special_key]
            
            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(key) { this.focus(); this.dispatchEvent(new KeyboardEvent('keydown', {key: key})); this.dispatchEvent(new KeyboardEvent('keyup', {key: key})); }"; arguments = @(@{ value = $key_code }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1059)
            LogWebDriverError $WebDriverErrorCodes.SEND_SPECIAL_KEY_ERROR "特殊キー送信エラー: $($_.Exception.Message)"
            throw "特殊キーの送信に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # マウス操作機能
    # ========================================

    # マウスホバー
    [void] MouseHover([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { this.dispatchEvent(new MouseEvent('mouseenter')); this.dispatchEvent(new MouseEvent('mouseover')); }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1057)
            LogWebDriverError $WebDriverErrorCodes.MOUSE_HOVER_ERROR "マウスホバーエラー: $($_.Exception.Message)"
            throw "マウスホバーに失敗しました: $($_.Exception.Message)"
        }
    }

    # マウスダブルクリック
    [void] DoubleClick([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { this.dispatchEvent(new MouseEvent('dblclick')); }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1055)
            LogWebDriverError $WebDriverErrorCodes.DOUBLE_CLICK_ERROR "ダブルクリックエラー: $($_.Exception.Message)"
            throw "ダブルクリックに失敗しました: $($_.Exception.Message)"
        }
    }

    # マウス右クリック
    [void] RightClick([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { this.dispatchEvent(new MouseEvent('contextmenu')); }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1056)
            LogWebDriverError $WebDriverErrorCodes.RIGHT_CLICK_ERROR "右クリックエラー: $($_.Exception.Message)"
            throw "右クリックに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # フォーム操作機能
    # ========================================

    # チェックボックスの選択/解除
    [void] SetCheckbox([string]$object_id, [bool]$checked)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function(checked) { this.checked = checked; this.dispatchEvent(new Event('change')); }"; arguments = @(@{ value = $checked }) })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1068)
            LogWebDriverError $WebDriverErrorCodes.SET_CHECKBOX_ERROR "チェックボックス設定エラー: $($_.Exception.Message)"
            throw "チェックボックスの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ラジオボタンの選択
    [void] SelectRadioButton([string]$object_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $object_id; functionDeclaration = "function() { this.checked = true; this.dispatchEvent(new Event('change')); }" })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1069)
            LogWebDriverError $WebDriverErrorCodes.SELECT_RADIO_BUTTON_ERROR "ラジオボタン選択エラー: $($_.Exception.Message)"
            throw "ラジオボタンの選択に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ファイルアップロード機能
    # ========================================

    # ファイルアップロード
    [void] UploadFile([string]$object_id, [string]$file_path)
    {
        try
        {
            if ([string]::IsNullOrEmpty($object_id))
            {
                throw "オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($file_path))
            {
                throw "ファイルパスが指定されていません。"
            }

            if (-not (Test-Path $file_path))
            {
                throw "指定されたファイルが見つかりません: $file_path"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # ファイルの内容をBase64エンコード
            $file_bytes = [System.IO.File]::ReadAllBytes($file_path)
            $file_base64 = [System.Convert]::ToBase64String($file_bytes)
            $file_name = Split-Path $file_path -Leaf

            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ 
                objectId = $object_id; 
                functionDeclaration = "function(fileName, fileData) { 
                    const file = new File([Uint8Array.from(atob(fileData), c => c.charCodeAt(0))], fileName);
                    const dataTransfer = new DataTransfer();
                    dataTransfer.items.add(file);
                    this.files = dataTransfer.files;
                    this.dispatchEvent(new Event('change'));
                }"; 
                arguments = @(@{ value = $file_name }, @{ value = $file_base64 }) 
            })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # 要素操作関連エラー (1070)
            LogWebDriverError $WebDriverErrorCodes.UPLOAD_FILE_ERROR "ファイルアップロードエラー: $($_.Exception.Message)"
            throw "ファイルアップロードに失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # JavaScript実行機能
    # ========================================

    # JavaScript実行（戻り値あり）
    [object] ExecuteScript([string]$script)
    {
        try
        {
            if ([string]::IsNullOrEmpty($script))
            {
                throw "JavaScriptスクリプトが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $script; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            return $response_json.result.result.value
        }
        catch
        {
            # JavaScript実行関連エラー (1111)
            LogWebDriverError $WebDriverErrorCodes.EXECUTE_SCRIPT_ERROR "JavaScript実行エラー: $($_.Exception.Message)"
            throw "JavaScriptの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # JavaScript実行（戻り値なし）
    [void] ExecuteScriptAsync([string]$script)
    {
        try
        {
            if ([string]::IsNullOrEmpty($script))
            {
                throw "JavaScriptスクリプトが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $script })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # JavaScript実行関連エラー (1112)
            LogWebDriverError $WebDriverErrorCodes.EXECUTE_SCRIPT_ASYNC_ERROR "JavaScript非同期実行エラー: $($_.Exception.Message)"
            throw "JavaScriptの非同期実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # クッキー操作機能
    # ========================================

    # クッキーの取得
    [string] GetCookie([string]$name)
    {
        try
        {
            if ([string]::IsNullOrEmpty($name))
            {
                throw "クッキー名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.cookie.split('; ').find(row => row.startsWith('$name='))?.split('=')[1]"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            return $response_json.result.result.value
        }
        catch
        {
            # ストレージ操作関連エラー (1121)
            LogWebDriverError $WebDriverErrorCodes.GET_COOKIE_ERROR "クッキー取得エラー: $($_.Exception.Message)"
            throw "クッキーの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # クッキーの設定
    [void] SetCookie([string]$name, [string]$value, [string]$domain = "", [string]$path = "/", [int]$expires_days = 30)
    {
        try
        {
            if ([string]::IsNullOrEmpty($name))
            {
                throw "クッキー名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expires_date = [datetime]::Now.AddDays($expires_days).ToString("ddd, dd MMM yyyy HH:mm:ss") + " GMT"
            $cookie_string = "$name=$value; expires=$expires_date; path=$path"
            
            if (-not [string]::IsNullOrEmpty($domain))
            {
                $cookie_string += "; domain=$domain"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.cookie = '$cookie_string'"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1122)
            LogWebDriverError $WebDriverErrorCodes.SET_COOKIE_ERROR "クッキー設定エラー: $($_.Exception.Message)"
            throw "クッキーの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # クッキーの削除
    [void] DeleteCookie([string]$name)
    {
        try
        {
            if ([string]::IsNullOrEmpty($name))
            {
                throw "クッキー名が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.cookie = '$name=; expires=Thu, 01 Jan 1970 00:00:00 GMT; path=/'"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1123)
            LogWebDriverError $WebDriverErrorCodes.DELETE_COOKIE_ERROR "クッキー削除エラー: $($_.Exception.Message)"
            throw "クッキーの削除に失敗しました: $($_.Exception.Message)"
        }
    }

    # すべてのクッキーを削除
    [void] ClearAllCookies()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.cookie.split(';').forEach(function(c) { document.cookie = c.replace(/^ +/, '').replace(/=.*/, '=;expires=' + new Date().toUTCString() + ';path=/'); });"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1124)
            LogWebDriverError $WebDriverErrorCodes.CLEAR_ALL_COOKIES_ERROR "全クッキー削除エラー: $($_.Exception.Message)"
            throw "全クッキーの削除に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # ローカルストレージ操作機能
    # ========================================

    # ローカルストレージの値を取得
    [string] GetLocalStorage([string]$key)
    {
        try
        {
            if ([string]::IsNullOrEmpty($key))
            {
                throw "キーが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "localStorage.getItem('$key')"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
            
            return $response_json.result.result.value
        }
        catch
        {
            # ストレージ操作関連エラー (1125)
            LogWebDriverError $WebDriverErrorCodes.GET_LOCAL_STORAGE_ERROR "ローカルストレージ取得エラー: $($_.Exception.Message)"
            throw "ローカルストレージの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ローカルストレージに値を設定
    [void] SetLocalStorage([string]$key, [string]$value)
    {
        try
        {
            if ([string]::IsNullOrEmpty($key))
            {
                throw "キーが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "localStorage.setItem('$key', '$value')"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1126)
            LogWebDriverError $WebDriverErrorCodes.SET_LOCAL_STORAGE_ERROR "ローカルストレージ設定エラー: $($_.Exception.Message)"
            throw "ローカルストレージの設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # ローカルストレージから値を削除
    [void] RemoveLocalStorage([string]$key)
    {
        try
        {
            if ([string]::IsNullOrEmpty($key))
            {
                throw "キーが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "localStorage.removeItem('$key')"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1127)
            LogWebDriverError $WebDriverErrorCodes.REMOVE_LOCAL_STORAGE_ERROR "ローカルストレージ削除エラー: $($_.Exception.Message)"
            throw "ローカルストレージの削除に失敗しました: $($_.Exception.Message)"
        }
    }

    # ローカルストレージをクリア
    [void] ClearLocalStorage()
    {
        try
        {
            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "localStorage.clear()"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "JavaScript実行エラー: $($response_json.error.message)"
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1128)
            LogWebDriverError $WebDriverErrorCodes.CLEAR_LOCAL_STORAGE_ERROR "ローカルストレージクリアエラー: $($_.Exception.Message)"
            throw "ローカルストレージのクリアに失敗しました: $($_.Exception.Message)"
        }
    }
}


