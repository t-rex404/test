class WebDriver
{
    [int]$browser_exe_process_id
    [System.Net.WebSockets.ClientWebSocket]$web_socket
    [int]$message_id
    [bool]$is_initialized

    # ログファイルパス（共有可能）
    static [string]$NormalLogFile = ".\WebDriver_$($env:USERNAME)_Normal.log"
    static [string]$ErrorLogFile = ".\WebDriver_$($env:USERNAME)_Error.log"

    # ========================================
    # ログユーティリティ
    # ========================================

    # 情報ログを出力
    [void] LogInfo([string]$message)
    {
        if ($global:Common)
        {
            try
            {
                $global:Common.WriteLog($message, "INFO")
            }
            catch
            {
                Write-Host "正常ログ出力に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
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
                $global:Common.HandleError($errorCode, $message, "WebDriver")
            }
            catch
            {
                Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message" | Out-File -Append -FilePath ([WebDriver]::ErrorLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        Write-Host $message -ForegroundColor Red
    }

    WebDriver()
    {
        try
        {
            $this.message_id = 0
            $this.is_initialized = $true
            $this.browser_exe_process_id = 0
            $this.web_socket = $null

            Write-Host "WebDriverの初期化が完了しました。" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("WebDriverの初期化が完了しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            # 初期化・接続関連エラー
            $this.LogError("WebDriverError_1001", "WebDriver初期化エラー: $($_.Exception.Message)")

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
            $argument_list = '--new-window ' + $url + ' --remote-debugging-port=9222 --disable-popup-blocking --no-first-run --disable-fre --user-data-dir=' + $browser_user_data_dir
                # 引数の意味
                # --new-window                 : 新しいウィンドウを開く。
                # --remote-debugging-port=9222 : デバッグ用 WebSocket をポート 9222 で有効化。
                # --disable-popup-blocking     : パップアップを無効化。
                # --no-first-run               : 最初の起動を無効化。
                # --disable-fre                : フリを無効化。
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

            # 正常ログ出力
            $this.LogInfo("ブラウザが正常に起動しました。プロセスID: $($this.browser_exe_process_id)")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            # 初期化・接続関連エラー
            $this.LogError("WebDriverError_1002", "ブラウザ起動エラー: $($_.Exception.Message)")

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
                    Write-Host "タブ情報を取得中... ($($i+1)/$retry_count)" -ForegroundColor Yellow
                    
                    $tabs = Invoke-RestMethod -Uri 'http://localhost:9222/json' -ErrorAction Stop -TimeoutSec 10
                    
                    if (-not $tabs)
                    {
                        throw 'タブ情報を取得できません。'
                    }

                    #$tab = $null
                    ##「about:blank」を選択
                    #foreach ($tab in $tabs)
                    #{
                    #    if ($tab.type -eq 'page' -and $tab.title -eq 'about:blank')
                    #    {
                    #        return $tab
                    #    }
                    #}
                    Write-Host "取得したタブ数: $($tabs.Count)" -ForegroundColor Cyan
                    
                    # まず "about:blank" の page を優先し、なければ最初の page、さらに無ければ先頭要素を返す
                    $tab = $tabs | Where-Object { $_.type -eq 'page' -and $_.title -eq 'about:blank' } | Select-Object -First 1
                    if (-not $tab) {
                        Write-Host "about:blankタブが見つからないため、最初のpageタブを検索中..." -ForegroundColor Yellow
                        $tab = $tabs | Where-Object { $_.type -eq 'page' } | Select-Object -First 1
                    }
                    if (-not $tab) {
                        Write-Host "pageタブが見つからないため、最初のタブを選択中..." -ForegroundColor Yellow
                        $tab = $tabs | Select-Object -First 1
                    }
                    
                    if ($tab) { 
                        Write-Host "選択されたタブ: type=$($tab.type), title=$($tab.title), id=$($tab.id)" -ForegroundColor Green

                        # webSocketDebuggerUrlの存在確認
                        if ([string]::IsNullOrEmpty($tab.webSocketDebuggerUrl))
                        {
                            Write-Host "警告: webSocketDebuggerUrlが空です。タブ情報を再取得します。" -ForegroundColor Yellow
                            if ($i -lt $retry_count - 1)
                            {
                                Start-Sleep -Seconds $retry_delay
                                continue
                            }
                            else
                            {
                                throw 'webSocketDebuggerUrlが取得できません。'
                            }
                        }

                        # 正常ログ出力
                        $this.LogInfo("タブ情報を取得しました: type=$($tab.type), title=$($tab.title), id=$($tab.id)")

                        return $tab 
                    }
                    
                    if ($i -lt $retry_count - 1)
                    {
                        Write-Host "条件に合致するタブが見つかりません。再試行します... ($($i+1)/$retry_count)" -ForegroundColor Yellow
                        Start-Sleep -Seconds $retry_delay
                    }
                    else
                    {
                        throw '条件に合致するタブが見つかりません。'
                    }
                }
                catch
                {
                    if ($i -lt $retry_count - 1)
                    {
                        Write-Host "タブ情報の取得に失敗しました。再試行します... ($($i+1)/$retry_count)" -ForegroundColor Yellow
                        Write-Host "エラー詳細: $($_.Exception.Message)" -ForegroundColor Red
                        Start-Sleep -Seconds $retry_delay
                    }
                    else
                    {
                        Write-Host "最大再試行回数に達しました。最後のエラーを再スローします。" -ForegroundColor Red
                        throw
                    }
                }
            }
            throw "タブ情報の取得に失敗しました。"
        }
        catch
        {
            # 初期化・接続関連エラー
            $this.LogError("WebDriverError_1003", "タブ情報取得エラー: $($_.Exception.Message)")
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

            Write-Host "WebSocket接続を開始します: $web_socket_debugger_url" -ForegroundColor Yellow

            $retry_count = 3
            $retry_delay = 2  # 秒
            
            for ($i = 0; $i -lt $retry_count; $i++)
            {
                $cancellationTokenSource = $null
                try
                {
                    Write-Host "WebSocket接続試行中... ($($i+1)/$retry_count)" -ForegroundColor Yellow
                    
                    # WebSocket接続の準備
                    $this.web_socket = [System.Net.WebSockets.ClientWebSocket]::new()
                    $uri = [System.Uri]::new($web_socket_debugger_url)
                    
                    # 接続タイムアウトを設定
                    $cancellationTokenSource = [System.Threading.CancellationTokenSource]::new()
                    $cancellationTokenSource.CancelAfter([TimeSpan]::FromSeconds(10))
                    
                    Write-Host "WebSocket接続を確立中..." -ForegroundColor Yellow
                    $this.web_socket.ConnectAsync($uri, $cancellationTokenSource.Token).Wait()
                    
                    if ($this.web_socket.State -eq [System.Net.WebSockets.WebSocketState]::Open)
                    {
                        $this.is_initialized = $true
                        Write-Host "WebSocket接続が確立されました。" -ForegroundColor Green

                        # 正常ログ出力
                        $this.LogInfo("WebSocket接続が確立されました: $web_socket_debugger_url")

                        return
                    }
                    else
                    {
                        throw "WebSocket接続が確立されませんでした。状態: $($this.web_socket.State)"
                    }
                }
                catch
                {
                    Write-Host "WebSocket接続エラー（試行 $($i+1)/$retry_count）: $($_.Exception.Message)" -ForegroundColor Red
                    
                    if ($this.web_socket)
                    {
                        try
                        {
                            $this.web_socket.Dispose()
                        }
                        catch
                        {
                            Write-Host "WebSocketの破棄中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                        $this.web_socket = $null
                    }
                    
                    if ($i -lt $retry_count - 1)
                    {
                        Write-Host "WebSocket接続に失敗しました。再試行します... ($($i+1)/$retry_count)" -ForegroundColor Yellow
                        Start-Sleep -Seconds $retry_delay
                    }
                    else
                    {
                        Write-Host "最大再試行回数に達しました。" -ForegroundColor Red
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
            # 初期化・接続関連エラー
            $this.LogError("WebDriverError_1004", "WebSocket接続エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("WebSocketメッセージを送信しました: $method")
        }
        catch
        {
            # 初期化・接続関連エラー
            $this.LogError("WebDriverError_1005", "WebSocket メッセージ送信エラー: $($_.Exception.Message)")
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
                            # 正常ログ出力
                            $this.LogInfo("WebSocketメッセージを受信しました: id=$($response_json_object.id)")
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
                    finally
                    {
                        if ($cancellationTokenSource)
                        {
                            $cancellationTokenSource.Dispose()
                        }
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
                }
            }
        }
        catch
        {
            # 初期化・接続関連エラー
            $this.LogError("WebDriverError_1006", "WebSocket メッセージ受信エラー: $($_.Exception.Message)")
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
            if (-not [System.Uri]::IsWellFormedUriString($url, [System.UriKind]::RelativeOrAbsolute))
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

            # 正常ログ出力
            $this.LogInfo("ページ遷移が完了しました: $url")
        }
        catch
        {
            # ナビゲーション関連エラー
            $this.LogError("WebDriverError_1101", "ページ遷移エラー: $($_.Exception.Message)")
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

                        # 正常ログ出力
                        $this.LogInfo("広告が正常に読み込まれました")

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
            # ナビゲーション関連エラー
            $this.LogError("WebDriverError_1102", "広告読み込み待機エラー: $($_.Exception.Message)")
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

                # 正常ログ出力
                $this.LogInfo("ウィンドウを正常に閉じました")
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1201", "ウィンドウを閉じるエラー: $($_.Exception.Message)")
            Write-Host "ウィンドウを閉じる際にエラーが発生しました: $($_.Exception.Message)"
        }
    }

    # WebSocketを閉じる
    [void] Dispose()
    {
        $cancellationTokenSource = $null
        try
        {
            # WebSocket接続を閉じる
            if ($this.web_socket -and $this.web_socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) 
            {
                try
                {
                    $cancellationTokenSource = [System.Threading.CancellationTokenSource]::new()
                    $cancellationTokenSource.CancelAfter([TimeSpan]::FromSeconds(5))
                    
                    $this.web_socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, 'Closing', $cancellationTokenSource.Token).Wait()
                    Write-Host "WebSocket接続を正常に閉じました。" -ForegroundColor Green

                    # 正常ログ出力
                    $this.LogInfo("WebSocket接続を正常に閉じました")
                }
                catch
                {
                    Write-Host "WebSocket接続を閉じることが出来ませんでした: $($_.Exception.Message)"
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

                        # 正常ログ出力
                        $this.LogInfo("ブラウザプロセスを正常に終了しました。プロセスID: $($this.browser_exe_process_id)")
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

            # 正常ログ出力
            $this.LogInfo("WebDriverのリソースを正常に解放しました")
        }
        catch
        {
            # 初期化・接続関連エラー
            $this.LogError("WebDriverError_1007", "Disposeエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ターゲット検出を有効化しました")

            return $response.result.targetId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1202", "ターゲット発見エラー: $($_.Exception.Message)")
            throw "ターゲットの発見に失敗しました: $($_.Exception.Message)"
        }
    }
    
    # 現在開かれているすべてのターゲット（タブやページ）の情報を取得
    # ターゲットには、タブのID、URL、タイトル、タイプなどが含まれる
    [pscustomobject] GetAvailableTabs()
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

            # 正常ログ出力
            $this.LogInfo("利用可能なタブ情報を取得しました")

            # PSCustomObject をそのまま返す（変換しない）
            return [pscustomobject]$response_json.result
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1203", "タブ情報取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("タブをアクティブにしました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1204", "タブアクティブ化エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("タブを閉じました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1205", "タブ切断エラー: $($_.Exception.Message)")
            throw "タブの切断に失敗しました: $($_.Exception.Message)"
        }
    }

    # セッション管理（attach/detach）
    # ターゲットにアタッチして sessionId を取得
    [string] AttachToTargetAndGetSessionId([string]$target_id, [bool]$flatten = $true)
    {
        try
        {
            if ([string]::IsNullOrEmpty($target_id))
            {
                throw "タブID（targetId）が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Target.attachToTarget', @{ targetId = $target_id; flatten = $flatten })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

            if ($response.error)
            {
                throw "ターゲットアタッチエラー: $($response.error.message)"
            }

            if ((-not $response.result) -or (-not $response.result.sessionId))
            {
                throw 'attach の応答に sessionId が含まれていません。'
            }

            # 正常ログ出力
            $this.LogInfo("ターゲットにアタッチしてセッションIDを取得しました")

            return [string]$response.result.sessionId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1206", "ターゲットアタッチ（sessionId取得）エラー: $($_.Exception.Message)")
            throw "ターゲットへのアタッチ（sessionId取得）に失敗しました: $($_.Exception.Message)"
        }
    }

    # 現在のタブにアタッチして sessionId を取得
    [string] AttachToCurrentTabAndGetSessionId([bool]$flatten = $true)
    {
        try
        {
            $tab = $this.GetTabInfomation()
            if (-not $tab -or [string]::IsNullOrEmpty($tab.id))
            {
                throw '現在タブの取得に失敗しました。'
            }

            $result = $this.AttachToTargetAndGetSessionId($tab.id, $flatten)

            # 正常ログ出力
            $this.LogInfo("現在のタブにアタッチしてセッションIDを取得しました")

            return $result
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1207", "現在タブアタッチ（sessionId取得）エラー: $($_.Exception.Message)")
            throw "現在タブへのアタッチ（sessionId取得）に失敗しました: $($_.Exception.Message)"
        }
    }

    # セッション（attach）をデタッチして解放
    [void] DetachSession([string]$session_id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($session_id))
            {
                throw "sessionId が指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $this.SendWebSocketMessage('Target.detachFromTarget', @{ sessionId = $session_id })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

            if ($response.error)
            {
                throw "セッションデタッチエラー: $($response.error.message)"
            }

            # 正常ログ出力
            $this.LogInfo("セッションをデタッチしました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1208", "セッションデタッチエラー: $($_.Exception.Message)")
            throw "セッションのデタッチに失敗しました: $($_.Exception.Message)"
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

            # 正常ログ出力
            $this.LogInfo("ページイベントを有効化しました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1209", "ページイベント有効化エラー: $($_.Exception.Message)")
            throw "ページイベントの有効化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ブラウザの表示倍率を変更
    [void] SetZoomLevel([double]$zoom_level)
    {
        try
        {
            if ($zoom_level -le 0)
            {
                throw "ズームレベルは正の値である必要があります。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # ズームレベルをパーセンテージに変換（例：1.0 = 100%, 1.5 = 150%）
            $zoom_percentage = [math]::Round($zoom_level * 100, 2)
            
            # DevTools Protocolを使用してズームレベルを設定
            $this.SendWebSocketMessage('Emulation.setPageScaleFactor', @{ pageScaleFactor = $zoom_level })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response.error)
            {
                throw "ズームレベル設定エラー: $($response.error.message)"
            }

            Write-Host "ブラウザの表示倍率を $zoom_percentage% に設定しました。" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("ブラウザの表示倍率を $zoom_percentage% に正常に設定しました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1210", "ブラウザ表示倍率変更エラー: $($_.Exception.Message)")
            throw "ブラウザの表示倍率変更に失敗しました: $($_.Exception.Message)"
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
            
            if ($null -ne $response_json.result.result.objectId)
            {
                # 正常ログ出力
                $this.LogInfo("要素を正常に検索しました。セレクタ: $selector")

                return @{ nodeId = $response_json.result.result.objectId; selector = $selector }
            }
            else
            {
                throw "CSSセレクタで要素を取得できません。セレクタ：$selector"
            }
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1301", "要素検索エラー (CSS): $($_.Exception.Message)")
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
                # 正常ログ出力
                $this.LogInfo("複数要素検索結果を返します (CSS)。セレクタ: $selector")
                return @{ nodeId = $response_json.result.result.objectIds; selector = $selector }
            }
            else
            {
                throw "CSSセレクタで複数の要素を取得できません。セレクタ：$selector"
            }
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1302", "複数要素検索エラー (CSS): $($_.Exception.Message)")
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
                # 正常ログ出力
                $this.LogInfo("要素を検索しました ($query_type): $element")
                return @{ nodeId = $response_json.result.result.objectId; query_type = $query_type; element = $element } 
            }
            else
            {
                throw "$query_type による要素取得に失敗しました。element：$element"
            }
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1303", "要素検索エラー ($query_type): $($_.Exception.Message)")
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
                # 正常ログ出力
                $this.LogInfo("複数要素数を取得しました ($query_type): $element, 件数: $($response_json.result.result.value)")
                return @{ count = $response_json.result.result.value; query_type = $query_type; element = $element }
            }
            else
            {
                throw "$query_type による複数の要素取得に失敗しました。element：$element"
            }
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1304", "複数要素検索エラー ($query_type): $($_.Exception.Message)")
            throw "$query_type による複数要素検索に失敗しました: $($_.Exception.Message)"
        }
    }
    


    # ========================================
    # 単体要素検索メソッド
    # ========================================

    # XPathで単体要素を検索
    [hashtable] FindElementByXPath([string]$xpath)
    {
        try
        {
            if ([string]::IsNullOrEmpty($xpath))
            {
                throw "XPathが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.evaluate('$xpath', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
            return $this.FindElementGeneric($expression, 'XPath', $xpath)
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1305", "XPath単体要素検索エラー: $($_.Exception.Message)")
            throw "XPathでの単体要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # class属性で単体要素を検索
    [hashtable] FindElementByClassName([string]$class_name, [int]$index = 0)
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

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.getElementsByClassName('$class_name')[$index]"
            return $this.FindElementGeneric($expression, 'ClassName', $class_name)
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1306", "ClassName単体要素検索エラー: $($_.Exception.Message)")
            throw "ClassNameでの単体要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # name属性で単体要素を検索
    [hashtable] FindElementByName([string]$name, [int]$index = 0)
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

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.getElementsByName('$name')[$index]"
            return $this.FindElementGeneric($expression, 'Name', $name)
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1307", "Name単体要素検索エラー: $($_.Exception.Message)")
            throw "Nameでの単体要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # id属性で単体要素を検索
    [hashtable] FindElementById([string]$id)
    {
        try
        {
            if ([string]::IsNullOrEmpty($id))
            {
                throw "IDが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.getElementById('$id')"
            return $this.FindElementGeneric($expression, 'Id', $id)
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1308", "Id単体要素検索エラー: $($_.Exception.Message)")
            throw "Idでの単体要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # tag名で単体要素を検索
    [hashtable] FindElementByTagName([string]$tag_name, [int]$index = 0)
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

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            $expression = "document.getElementsByTagName('$tag_name')[$index]"
            return $this.FindElementGeneric($expression, 'TagName', $tag_name)
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1309", "TagName単体要素検索エラー: $($_.Exception.Message)")
            throw "TagNameでの単体要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 複数要素検索メソッド
    # ========================================

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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1310", "複数要素検索エラー (ClassName): $($_.Exception.Message)")
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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1311", "複数要素検索エラー (Name): $($_.Exception.Message)")
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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1312", "複数要素検索エラー (TagName): $($_.Exception.Message)")
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
                # 正常ログ出力
                $this.LogInfo("要素の存在を確認しました")

                return $true
            }
            else
            {
                return $false
            }
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1313", "要素存在確認エラー ($query_type): $($_.Exception.Message)")
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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1314", "XPath要素存在確認エラー: $($_.Exception.Message)")
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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1315", "ClassName要素存在確認エラー: $($_.Exception.Message)")
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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1316", "Id要素存在確認エラー: $($_.Exception.Message)")
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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1317", "Name要素存在確認エラー: $($_.Exception.Message)")
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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1318", "TagName要素存在確認エラー: $($_.Exception.Message)")
            throw "TagNameでの要素存在確認に失敗しました: $($_.Exception.Message)"
        }
    }

    # 親オブジェクト内の子オブジェクト単数要素検索（CSSセレクタ）
    [hashtable] FindChildElement([string]$parent_object_id, [string]$child_selector)
    {
        try
        {
            if ([string]::IsNullOrEmpty($parent_object_id))
            {
                throw "親オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($child_selector))
            {
                throw "子要素のセレクタが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # 親要素内で子要素を検索するJavaScriptを実行
            $expression = "function(parentId, selector) { 
                const parent = this; 
                const child = parent.querySelector(selector); 
                return child ? child : null; 
            }"
            
            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ 
                objectId = $parent_object_id; 
                functionDeclaration = $expression; 
                arguments = @(@{ value = $child_selector })
            })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "子要素検索エラー: $($response_json.error.message)"
            }
            
            if ($null -ne $response_json.result.result.value)
            {
                # 正常ログ出力
                $this.LogInfo("子要素を検索しました")

                return @{ nodeId = $response_json.result.result.objectId; selector = $child_selector; parentId = $parent_object_id }
            }
            else
            {
                throw "親要素内で子要素を取得できません。セレクタ：$child_selector"
            }
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1319", "親オブジェクト内子要素単数検索エラー: $($_.Exception.Message)")
            throw "親オブジェクト内の子要素検索に失敗しました: $($_.Exception.Message)"
        }
    }

    # 親オブジェクト内の子オブジェクト複数要素検索（CSSセレクタ）
    [array] FindChildElements([string]$parent_object_id, [string]$child_selector)
    {
        try
        {
            if ([string]::IsNullOrEmpty($parent_object_id))
            {
                throw "親オブジェクトIDが指定されていません。"
            }

            if ([string]::IsNullOrEmpty($child_selector))
            {
                throw "子要素のセレクタが指定されていません。"
            }

            if (-not $this.is_initialized)
            {
                throw "WebDriverが初期化されていません。"
            }

            # 親要素内で子要素を複数検索するJavaScriptを実行
            $expression = "function(parentId, selector) { 
                const parent = this; 
                const children = parent.querySelectorAll(selector); 
                return Array.from(children); 
            }"
            
            $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ 
                objectId = $parent_object_id; 
                functionDeclaration = $expression; 
                arguments = @(@{ value = $child_selector })
            })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "子要素複数検索エラー: $($response_json.error.message)"
            }
            
            if ($response_json.result.result.value -and $response_json.result.result.value.Count -gt 0)
            {
                $element_list = @()
                for ($i = 0; $i -lt $response_json.result.result.value.Count; $i++)
                {
                    try
                    {
                        # 各子要素のobjectIdを取得
                        $child_expression = "function(index) { return this[index]; }"
                        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ 
                            objectId = $response_json.result.result.objectId; 
                            functionDeclaration = $child_expression; 
                            arguments = @(@{ value = $i })
                        })
                        $child_response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                        
                        if ($child_response.result.result.objectId)
                        {
                            $element_list += @{ nodeId = $child_response.result.result.objectId; selector = $child_selector; parentId = $parent_object_id; index = $i }
                        }
                    }
                    catch
                    {
                        Write-Host "インデックス $i の子要素取得に失敗しましたが、処理を続行します: $($_.Exception.Message)"
                    }
                }

                # 正常ログ出力
                $this.LogInfo("子要素を複数検索しました")

                return $element_list
            }
            else
            {
                return @()
            }
        }
        catch
        {
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1320", "親オブジェクト内子要素複数検索エラー: $($_.Exception.Message)")
            throw "親オブジェクト内の子要素複数検索に失敗しました: $($_.Exception.Message)"
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

            # 正常ログ出力
            $this.LogInfo("要素からテキストを正常に取得しました。オブジェクトID: $object_id")

            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1401", "要素テキスト取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素にテキストを正常に設定しました。オブジェクトID: $object_id, テキスト: $text")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1402", "要素テキスト設定エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素を正常にクリックしました。オブジェクトID: $object_id")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1403", "要素クリックエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素の属性を正常に取得しました。オブジェクトID: $object_id, 属性: $attribute")

            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1404", "要素属性取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素の属性を正常に設定しました。オブジェクトID: $object_id, 属性: $attribute, 値: $value")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1405", "要素属性設定エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("アンカータグのhrefを正常に取得しました。オブジェクトID: $object_id")

            return $response.result.result.value
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1406", "href取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素のCSSプロパティを正常に取得しました。オブジェクトID: $object_id, プロパティ: $property")

            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1407", "CSSプロパティ取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素のCSSプロパティを正常に設定しました。オブジェクトID: $object_id, プロパティ: $property, 値: $value")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1408", "CSSプロパティ設定エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("インデックスによるオプション選択が正常に完了しました。オブジェクトID: $object_id, インデックス: $index")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1409", "オプション選択エラー (インデックス): $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("テキストによるオプション選択が正常に完了しました。オブジェクトID: $object_id, テキスト: $text")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1410", "オプション選択エラー (テキスト): $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("全オプションの選択解除が正常に完了しました。オブジェクトID: $object_id")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1411", "オプション未選択エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素の値を正常にクリアしました。オブジェクトID: $object_id")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1412", "要素クリアエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ウィンドウサイズを正常に変更しました。ハンドル: $windowHandle, 幅: $width, 高さ: $height")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1211", "ウィンドウリサイズエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ウィンドウを正常状態に変更しました。ハンドル: $windowHandle")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1212", "ウィンドウ状態変更エラー (通常): $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ウィンドウを最大化しました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1213", "ウィンドウ状態変更エラー (最大化): $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ウィンドウを最小化しました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1214", "ウィンドウ状態変更エラー (最小化): $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ウィンドウをフルスクリーンにしました")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1215", "ウィンドウ状態変更エラー (フルスクリーン): $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ウィンドウを正常に移動しました。ハンドル: $windowHandle, X: $x, Y: $y")
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1216", "ウィンドウ移動エラー: $($_.Exception.Message)")
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

            # 履歴を取得し、1つ前の entryId に移動
            $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $history = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            if ($history.error) { throw "履歴取得エラー: $($history.error.message)" }

            $currentIndex = $history.result.currentIndex
            $entries      = $history.result.entries
            if ($null -eq $currentIndex -or $null -eq $entries -or $entries.Count -eq 0) { return }
            if ($currentIndex -le 0) { return }

            $targetEntryId = $entries[$currentIndex - 1].id
            $this.SendWebSocketMessage('Page.navigateToHistoryEntry', @{ entryId = $targetEntryId })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            if ($response.error) { throw "ブラウザ履歴移動エラー: $($response.error.message)" }

            # 正常ログ出力
            $this.LogInfo("ブラウザ履歴を前のページに戻しました")
        }
        catch
        {
            # ナビゲーション関連エラー
            $this.LogError("WebDriverError_1103", "ブラウザ履歴移動エラー (戻る): $($_.Exception.Message)")
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

            # 履歴を取得し、1つ先の entryId に移動
            $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $history = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            if ($history.error) { throw "履歴取得エラー: $($history.error.message)" }

            $currentIndex = $history.result.currentIndex
            $entries      = $history.result.entries
            if ($null -eq $currentIndex -or $null -eq $entries -or $entries.Count -eq 0) { return }
            if ($currentIndex -ge ($entries.Count - 1)) { return }

            $targetEntryId = $entries[$currentIndex + 1].id
            $this.SendWebSocketMessage('Page.navigateToHistoryEntry', @{ entryId = $targetEntryId })
            $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            if ($response.error) { throw "ブラウザ履歴移動エラー: $($response.error.message)" }

            # 正常ログ出力
            $this.LogInfo("ブラウザ履歴を次のページに進みました")
        }
        catch
        {
            # ナビゲーション関連エラー
            $this.LogError("WebDriverError_1104", "ブラウザ履歴移動エラー (進む): $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ページを更新しました")
        }
        catch
        {
            # ナビゲーション関連エラー
            $this.LogError("WebDriverError_1105", "ブラウザ更新エラー: $($_.Exception.Message)")
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

            #$this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.URL"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "URL取得エラー: $($response_json.error.message)"
            }
            
            #if (-not $response_json.result.entries -or $response_json.result.entries.Count -eq 0)
            #{
            #    throw "ナビゲーション履歴が空です。"
            #}
            if ($null -eq $response_json.result.result.value)
            {
                throw "URLが取得できませんでした。"
            }

            # 正常ログ出力
            $this.LogInfo("現在のURLを取得しました")

            #return $response_json.result.entries[0].url
            return $response_json.result.result.value
        }
        catch
        {
            # 情報取得関連エラー
            $this.LogError("WebDriverError_1501", "URL取得エラー: $($_.Exception.Message)")
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

            #$this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.title"; returnByValue = $true })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "タイトル取得エラー: $($response_json.error.message)"
            }
            
            #if (-not $response_json.result.entries -or $response_json.result.entries.Count -eq 0)
            #{
            #    throw "ナビゲーション履歴が空です。"
            #}
            if ($null -eq $response_json.result.result.value)
            {
                throw "タイトルが取得できませんでした。"
            }

            # 正常ログ出力
            $this.LogInfo("ページタイトルを取得しました")

            #return $response_json.result.entries[0].title
            return $response_json.result.result.value
        }
        catch
        {
            # 情報取得関連エラー
            $this.LogError("WebDriverError_1502", "タイトル取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ページのソースコードを取得しました")

            return $response_json.result.result.value
        }
        catch
        {
            # 情報取得関連エラー
            $this.LogError("WebDriverError_1503", "ソースコード取得エラー: $($_.Exception.Message)")
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

            # 現在のターゲットの windowId を取得
            #$this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
            $this.SendWebSocketMessage('Browser.getWindowForTarget', @{ })
            $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
            
            if ($response_json.error)
            {
                throw "ウィンドウハンドル取得エラー: $($response_json.error.message)"
            }

            # 正常ログ出力
            $this.LogInfo("ウィンドウハンドルを取得しました")

            #if (-not $response_json.result.entries -or $response_json.result.entries.Count -eq 0)
            #{
            #    throw "ナビゲーション履歴が空です。"
            #}

            #return $response_json.result.entries[0].windowId
            return $response_json.result.windowId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1217", "ウィンドウハンドル取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("全ウィンドウハンドルを取得しました")

            return $window_handles
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1218", "複数ウィンドウハンドル取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ウィンドウサイズを取得しました")

            return @{
                width = $response_json.result.result.value.width
                height = $response_json.result.result.value.height
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー
            $this.LogError("WebDriverError_1219", "ウィンドウサイズ取得エラー: $($_.Exception.Message)")
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

                        # 正常ログ出力
                        $this.LogInfo("フルページスクリーンショットを正常に保存しました: $save_path")
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

                        # 正常ログ出力
                        $this.LogInfo("ビューポートスクリーンショットを正常に保存しました: $save_path")
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
            # スクリーンショット関連エラー
            $this.LogError("WebDriverError_1601", "スクリーンショット取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("要素スクリーンショットを正常に保存しました: $save_path")
        }
        catch
        {
            # スクリーンショット関連エラー
            $this.LogError("WebDriverError_1602", "要素スクリーンショット取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("複数要素スクリーンショットを正常に保存しました: $save_path")
        }
        catch
        {
            # スクリーンショット関連エラー
            $this.LogError("WebDriverError_1603", "複数要素スクリーンショット取得エラー: $($_.Exception.Message)")
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

                        # 正常ログ出力
                        $this.LogInfo("要素が正常に表示されました: $selector")

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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1321", "要素表示待機エラー: $($_.Exception.Message)")
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

                        # 正常ログ出力
                        $this.LogInfo("要素が正常にクリック可能になりました: $selector")

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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1322", "要素クリック可能性待機エラー: $($_.Exception.Message)")
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

                        # 正常ログ出力
                        $this.LogInfo("ページロードが正常に完了しました")

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
            # ナビゲーション関連エラー
            $this.LogError("WebDriverError_1106", "ページロード待機エラー: $($_.Exception.Message)")
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

                        # 正常ログ出力
                        $this.LogInfo("カスタム条件が正常に満たされました: $javascript_condition")

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
            # 要素検索関連エラー
            $this.LogError("WebDriverError_1323", "カスタム条件待機エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("キーボード入力を送信しました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1413", "キーボード入力エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("特殊キーを送信しました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1414", "特殊キー送信エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("マウスホバーを実行しました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1415", "マウスホバーエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ダブルクリックを実行しました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1416", "ダブルクリックエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("右クリックを実行しました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1417", "右クリックエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("チェックボックスを設定しました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1418", "チェックボックス設定エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ラジオボタンを選択しました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1419", "ラジオボタン選択エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ファイルをアップロードしました")
        }
        catch
        {
            # 要素操作関連エラー
            $this.LogError("WebDriverError_1420", "ファイルアップロードエラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("JavaScriptスクリプトを正常に実行しました。スクリプト: $script")

            return $response_json.result.result.value
        }
        catch
        {
            # JavaScript実行関連エラー
            $this.LogError("WebDriverError_1701", "JavaScript実行エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("JavaScriptスクリプトを非同期で実行しました")
        }
        catch
        {
            # JavaScript実行関連エラー
            $this.LogError("WebDriverError_1702", "JavaScript非同期実行エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("クッキーを取得しました")

            return $response_json.result.result.value
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1801", "クッキー取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("クッキーを設定しました")
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1802", "クッキー設定エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("クッキーを削除しました")
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1803", "クッキー削除エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("全てのクッキーを削除しました")
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1804", "全クッキー削除エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ローカルストレージから値を取得しました")

            return $response_json.result.result.value
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1805", "ローカルストレージ取得エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ローカルストレージに値を設定しました")
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1806", "ローカルストレージ設定エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ローカルストレージから値を削除しました")
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1807", "ローカルストレージ削除エラー: $($_.Exception.Message)")
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

            # 正常ログ出力
            $this.LogInfo("ローカルストレージをクリアしました")
        }
        catch
        {
            # ストレージ操作関連エラー
            $this.LogError("WebDriverError_1808", "ローカルストレージクリアエラー: $($_.Exception.Message)")
            throw "ローカルストレージのクリアに失敗しました: $($_.Exception.Message)"
        }
    }
}






