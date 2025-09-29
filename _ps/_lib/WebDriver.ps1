class WebDriver
{
    [int]$browser_exe_process_id
    [System.Net.WebSockets.ClientWebSocket]$web_socket
    [int]$message_id
    [bool]$is_initialized

    # ログファイルパス（共有可能）
    static [string]$NormalLogFile = ".\WebDriver_$($env:USERNAME)_Normal.log"
    static [string]$ErrorLogFile = ".\WebDriver_$($env:USERNAME)_Error.log"

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
            if ($global:Common)
            {
                $global:Common.WriteLog("WebDriverの初期化が完了しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] WebDriverの初期化が完了しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            # 初期化・接続関連エラー (1001)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0001", "WebDriver初期化エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WebDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

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
            if ($global:Common)
            {
                $global:Common.WriteLog("ブラウザが正常に起動しました。プロセスID: $($this.browser_exe_process_id)", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ブラウザが正常に起動しました。プロセスID: $($this.browser_exe_process_id)" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            # 初期化・接続関連エラー (1002)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0010", "ブラウザ起動エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile)
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ブラウザの起動に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }

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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("タブ情報を取得しました: type=$($tab.type), title=$($tab.title), id=$($tab.id)", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] タブ情報を取得しました: type=$($tab.type), title=$($tab.title), id=$($tab.id)" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }

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
            # 初期化・接続関連エラー (1003)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0011", "タブ情報取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "タブ情報の取得に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("WebSocket接続が確立されました: $web_socket_debugger_url", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] WebSocket接続が確立されました: $web_socket_debugger_url" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }

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
            # 初期化・接続関連エラー (1004)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0012", "WebSocket接続エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WebSocket接続エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("WebSocketメッセージを送信しました: $method", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] WebSocketメッセージを送信しました: $method" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 初期化・接続関連エラー (1005)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0013", "WebSocket メッセージ送信エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WebSocket メッセージ送信エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                            if ($global:Common)
                            {
                                $global:Common.WriteLog("WebSocketメッセージを受信しました: id=$($response_json_object.id)", "INFO")
                                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] WebSocketメッセージを受信しました: id=$($response_json_object.id)" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                            }
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
            # 初期化・接続関連エラー (1006)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0014", "WebSocket メッセージ受信エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "WebSocket メッセージ受信エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ページ遷移が完了しました: $url", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ページ遷移が完了しました: $url" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ナビゲーション関連エラー (1011)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0020", "ページ遷移エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ページ遷移エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("広告が正常に読み込まれました", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 広告が正常に読み込まれました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }

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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0021", "広告読み込み待機エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "広告読み込み待機エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("ウィンドウを正常に閉じました", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウを正常に閉じました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1077)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0022", "ウィンドウを閉じるエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウを閉じるエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                    if ($global:Common)
                    {
                        $global:Common.WriteLog("WebSocket接続を正常に閉じました", "INFO")
                        "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] WebSocket接続を正常に閉じました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                    }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("ブラウザプロセスを正常に終了しました。プロセスID: $($this.browser_exe_process_id)", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ブラウザプロセスを正常に終了しました。プロセスID: $($this.browser_exe_process_id)" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("WebDriverのリソースを正常に解放しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] WebDriverのリソースを正常に解放しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 初期化・接続関連エラー (1007)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0090", "Disposeエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Disposeエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ターゲット検出を有効化しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ターゲット検出を有効化しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response.result.targetId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1081)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0015", "ターゲット発見エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ターゲット発見エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("利用可能なタブ情報を取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 利用可能なタブ情報を取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            # PSCustomObject をそのまま返す（変換しない）
            return [pscustomobject]$response_json.result
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1082)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0016", "タブ情報取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "タブ情報取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("タブをアクティブにしました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] タブをアクティブにしました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1083)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0017", "タブアクティブ化エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "タブアクティブ化エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("タブを閉じました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] タブを閉じました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1084)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0018", "タブ切断エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "タブ切断エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ターゲットにアタッチしてセッションIDを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ターゲットにアタッチしてセッションIDを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return [string]$response.result.sessionId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1086)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0019", "ターゲットアタッチ（sessionId取得）エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ターゲットアタッチ（sessionId取得）エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("現在のタブにアタッチしてセッションIDを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 現在のタブにアタッチしてセッションIDを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $result
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1086)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0015", "現在タブアタッチ（sessionId取得）エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "現在タブアタッチ（sessionId取得）エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("セッションをデタッチしました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] セッションをデタッチしました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1087)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0016", "セッションデタッチエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "セッションデタッチエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ページイベントを有効化しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ページイベントを有効化しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1085)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0017", "ページイベント有効化エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ページイベント有効化エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ブラウザの表示倍率を $zoom_percentage% に正常に設定しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ブラウザの表示倍率を $zoom_percentage% に正常に設定しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1072)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0018", "ブラウザ表示倍率変更エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ブラウザ表示倍率変更エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("要素を正常に検索しました。セレクタ: $selector", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素を正常に検索しました。セレクタ: $selector" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }

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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0040", "要素検索エラー (CSS): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素検索エラー (CSS): $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("複数要素検索結果を返します (CSS)。セレクタ: $selector", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 複数要素検索結果を返します (CSS)。セレクタ: $selector" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0041", "複数要素検索エラー (CSS): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "複数要素検索エラー (CSS): $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("要素を検索しました ($query_type): $element", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素を検索しました ($query_type): $element" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0042", "要素検索エラー ($query_type): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素検索エラー ($query_type): $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("複数要素数を取得しました ($query_type): $element, 件数: $($response_json.result.result.value)", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 複数要素数を取得しました ($query_type): $element, 件数: $($response_json.result.result.value)" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0043", "複数要素検索エラー ($query_type): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "複数要素検索エラー ($query_type): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            # 要素検索関連エラー (1023)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0044", "XPath単体要素検索エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "XPath単体要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            # 要素検索関連エラー (1024)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0045", "ClassName単体要素検索エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ClassName単体要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            # 要素検索関連エラー (1025)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0046", "Name単体要素検索エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Name単体要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            # 要素検索関連エラー (1026)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0047", "Id単体要素検索エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Id単体要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            # 要素検索関連エラー (1027)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0048", "TagName単体要素検索エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TagName単体要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            # 要素検索関連エラー (1030)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0049", "複数要素検索エラー (ClassName): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "複数要素検索エラー (ClassName): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0050", "複数要素検索エラー (Name): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "複数要素検索エラー (Name): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0051", "複数要素検索エラー (TagName): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "複数要素検索エラー (TagName): $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("要素の存在を確認しました", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素の存在を確認しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }

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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0052", "要素存在確認エラー ($query_type): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素存在確認エラー ($query_type): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0053", "XPath要素存在確認エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "XPath要素存在確認エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0054", "ClassName要素存在確認エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ClassName要素存在確認エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0055", "Id要素存在確認エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Id要素存在確認エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0056", "Name要素存在確認エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Name要素存在確認エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0057", "TagName要素存在確認エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "TagName要素存在確認エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("子要素を検索しました", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 子要素を検索しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }

                return @{ nodeId = $response_json.result.result.objectId; selector = $child_selector; parentId = $parent_object_id }
            }
            else
            {
                throw "親要素内で子要素を取得できません。セレクタ：$child_selector"
            }
        }
        catch
        {
            # 要素検索関連エラー (1042)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0058", "親オブジェクト内子要素単数検索エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "親オブジェクト内子要素単数検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                if ($global:Common)
                {
                    $global:Common.WriteLog("子要素を複数検索しました", "INFO")
                    "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 子要素を複数検索しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                }

                return $element_list
            }
            else
            {
                return @()
            }
        }
        catch
        {
            # 要素検索関連エラー (1043)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0059", "親オブジェクト内子要素複数検索エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "親オブジェクト内子要素複数検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素からテキストを正常に取得しました。オブジェクトID: $object_id", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素からテキストを正常に取得しました。オブジェクトID: $object_id" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1051)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0060", "要素テキスト取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素テキスト取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素にテキストを正常に設定しました。オブジェクトID: $object_id, テキスト: $text", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素にテキストを正常に設定しました。オブジェクトID: $object_id, テキスト: $text" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1052)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0061", "要素テキスト設定エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素テキスト設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素を正常にクリックしました。オブジェクトID: $object_id", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素を正常にクリックしました。オブジェクトID: $object_id" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1054)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0062", "要素クリックエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素クリックエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素の属性を正常に取得しました。オブジェクトID: $object_id, 属性: $attribute", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素の属性を正常に取得しました。オブジェクトID: $object_id, 属性: $attribute" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1060)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0063", "要素属性取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素属性取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素の属性を正常に設定しました。オブジェクトID: $object_id, 属性: $attribute, 値: $value", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素の属性を正常に設定しました。オブジェクトID: $object_id, 属性: $attribute, 値: $value" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1061)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0064", "要素属性設定エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素属性設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("アンカータグのhrefを正常に取得しました。オブジェクトID: $object_id", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] アンカータグのhrefを正常に取得しました。オブジェクトID: $object_id" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1064)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0065", "href取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "href取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素のCSSプロパティを正常に取得しました。オブジェクトID: $object_id, プロパティ: $property", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素のCSSプロパティを正常に取得しました。オブジェクトID: $object_id, プロパティ: $property" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response_json.result.result.value
        }
        catch
        {
            # 要素操作関連エラー (1062)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0066", "CSSプロパティ取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "CSSプロパティ取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素のCSSプロパティを正常に設定しました。オブジェクトID: $object_id, プロパティ: $property, 値: $value", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素のCSSプロパティを正常に設定しました。オブジェクトID: $object_id, プロパティ: $property, 値: $value" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1063)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0067", "CSSプロパティ設定エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "CSSプロパティ設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("インデックスによるオプション選択が正常に完了しました。オブジェクトID: $object_id, インデックス: $index", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] インデックスによるオプション選択が正常に完了しました。オブジェクトID: $object_id, インデックス: $index" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1065)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0068", "オプション選択エラー (インデックス): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "オプション選択エラー (インデックス): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("テキストによるオプション選択が正常に完了しました。オブジェクトID: $object_id, テキスト: $text", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] テキストによるオプション選択が正常に完了しました。オブジェクトID: $object_id, テキスト: $text" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1066)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0069", "オプション選択エラー (テキスト): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "オプション選択エラー (テキスト): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("全オプションの選択解除が正常に完了しました。オブジェクトID: $object_id", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 全オプションの選択解除が正常に完了しました。オブジェクトID: $object_id" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1067)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0070", "オプション未選択エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "オプション未選択エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素の値を正常にクリアしました。オブジェクトID: $object_id", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素の値を正常にクリアしました。オブジェクトID: $object_id" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1053)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0071", "要素クリアエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素クリアエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウサイズを正常に変更しました。ハンドル: $windowHandle, 幅: $width, 高さ: $height", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウサイズを正常に変更しました。ハンドル: $windowHandle, 幅: $width, 高さ: $height" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1071)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0072", "ウィンドウリサイズエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウリサイズエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウを正常状態に変更しました。ハンドル: $windowHandle", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウを正常状態に変更しました。ハンドル: $windowHandle" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1072)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0073", "ウィンドウ状態変更エラー (通常): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウ状態変更エラー (通常): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウを最大化しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウを最大化しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1073)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0074", "ウィンドウ状態変更エラー (最大化): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウ状態変更エラー (最大化): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウを最小化しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウを最小化しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1074)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0075", "ウィンドウ状態変更エラー (最小化): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウ状態変更エラー (最小化): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウをフルスクリーンにしました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウをフルスクリーンにしました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1075)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0076", "ウィンドウ状態変更エラー (フルスクリーン): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウ状態変更エラー (フルスクリーン): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウを正常に移動しました。ハンドル: $windowHandle, X: $x, Y: $y", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウを正常に移動しました。ハンドル: $windowHandle, X: $x, Y: $y" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1076)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0023", "ウィンドウ移動エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウ移動エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ブラウザ履歴を前のページに戻しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ブラウザ履歴を前のページに戻しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ナビゲーション関連エラー (1012)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0024", "ブラウザ履歴移動エラー (戻る): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ブラウザ履歴移動エラー (戻る): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ブラウザ履歴を次のページに進みました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ブラウザ履歴を次のページに進みました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ナビゲーション関連エラー (1013)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0025", "ブラウザ履歴移動エラー (進む): $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ブラウザ履歴移動エラー (進む): $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ページを更新しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ページを更新しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ナビゲーション関連エラー (1014)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0026", "ブラウザ更新エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ブラウザ更新エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("現在のURLを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 現在のURLを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            #return $response_json.result.entries[0].url
            return $response_json.result.result.value
        }
        catch
        {
            # 情報取得関連エラー (1091)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0027", "URL取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "URL取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ページタイトルを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ページタイトルを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            #return $response_json.result.entries[0].title
            return $response_json.result.result.value
        }
        catch
        {
            # 情報取得関連エラー (1092)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0028", "タイトル取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "タイトル取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ページのソースコードを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ページのソースコードを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response_json.result.result.value
        }
        catch
        {
            # 情報取得関連エラー (1093)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0029", "ソースコード取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ソースコード取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウハンドルを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウハンドルを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            #if (-not $response_json.result.entries -or $response_json.result.entries.Count -eq 0)
            #{
            #    throw "ナビゲーション履歴が空です。"
            #}

            #return $response_json.result.entries[0].windowId
            return $response_json.result.windowId
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1078)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0079", "ウィンドウハンドル取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウハンドル取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("全ウィンドウハンドルを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 全ウィンドウハンドルを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $window_handles
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1079)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0075", "複数ウィンドウハンドル取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "複数ウィンドウハンドル取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ウィンドウサイズを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ウィンドウサイズを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return @{
                width = $response_json.result.result.value.width
                height = $response_json.result.result.value.height
            }
        }
        catch
        {
            # ウィンドウ・タブ操作関連エラー (1080)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0076", "ウィンドウサイズ取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウサイズ取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("フルページスクリーンショットを正常に保存しました: $save_path", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] フルページスクリーンショットを正常に保存しました: $save_path" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("ビューポートスクリーンショットを正常に保存しました: $save_path", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ビューポートスクリーンショットを正常に保存しました: $save_path" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }
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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0080", "スクリーンショット取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "スクリーンショット取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("要素スクリーンショットを正常に保存しました: $save_path", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素スクリーンショットを正常に保存しました: $save_path" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # スクリーンショット関連エラー (1102)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0081", "要素スクリーンショット取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素スクリーンショット取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("複数要素スクリーンショットを正常に保存しました: $save_path", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 複数要素スクリーンショットを正常に保存しました: $save_path" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # スクリーンショット関連エラー (1103)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0082", "複数要素スクリーンショット取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "複数要素スクリーンショット取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("要素が正常に表示されました: $selector", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素が正常に表示されました: $selector" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }

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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0031", "要素表示待機エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素表示待機エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("要素が正常にクリック可能になりました: $selector", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 要素が正常にクリック可能になりました: $selector" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }

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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0032", "要素クリック可能性待機エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素クリック可能性待機エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("ページロードが正常に完了しました", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ページロードが正常に完了しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }

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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0030", "ページロード待機エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ページロード待機エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
                        if ($global:Common)
                        {
                            $global:Common.WriteLog("カスタム条件が正常に満たされました: $javascript_condition", "INFO")
                            "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] カスタム条件が正常に満たされました: $javascript_condition" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
                        }

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
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0033", "カスタム条件待機エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "カスタム条件待機エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("キーボード入力を送信しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] キーボード入力を送信しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1058)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0037", "キーボード入力エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "キーボード入力エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("特殊キーを送信しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 特殊キーを送信しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1059)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0038", "特殊キー送信エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "特殊キー送信エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("マウスホバーを実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] マウスホバーを実行しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1057)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0036", "マウスホバーエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "マウスホバーエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ダブルクリックを実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ダブルクリックを実行しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1055)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0034", "ダブルクリックエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ダブルクリックエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("右クリックを実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 右クリックを実行しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1056)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0035", "右クリックエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "右クリックエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("チェックボックスを設定しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] チェックボックスを設定しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1068)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0039", "チェックボックス設定エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "チェックボックス設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ラジオボタンを選択しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ラジオボタンを選択しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1069)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0077", "ラジオボタン選択エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ラジオボタン選択エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ファイルをアップロードしました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ファイルをアップロードしました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # 要素操作関連エラー (1070)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0078", "ファイルアップロードエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
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
            if ($global:Common)
            {
                $global:Common.WriteLog("JavaScriptスクリプトを正常に実行しました。スクリプト: $script", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] JavaScriptスクリプトを正常に実行しました。スクリプト: $script" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response_json.result.result.value
        }
        catch
        {
            # JavaScript実行関連エラー (1111)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0083", "JavaScript実行エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "JavaScript実行エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("JavaScriptスクリプトを非同期で実行しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] JavaScriptスクリプトを非同期で実行しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # JavaScript実行関連エラー (1112)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0084", "JavaScript非同期実行エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "JavaScript非同期実行エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("クッキーを取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] クッキーを取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response_json.result.result.value
        }
        catch
        {
            # ストレージ操作関連エラー (1121)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0085", "クッキー取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "クッキー取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("クッキーを設定しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] クッキーを設定しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1122)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0086", "クッキー設定エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "クッキー設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("クッキーを削除しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] クッキーを削除しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1123)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0087", "クッキー削除エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "クッキー削除エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("全てのクッキーを削除しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] 全てのクッキーを削除しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1124)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0088", "全クッキー削除エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "全クッキー削除エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ローカルストレージから値を取得しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ローカルストレージから値を取得しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            return $response_json.result.result.value
        }
        catch
        {
            # ストレージ操作関連エラー (1125)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0087", "ローカルストレージ取得エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ローカルストレージ取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ローカルストレージに値を設定しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ローカルストレージに値を設定しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1126)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0088", "ローカルストレージ設定エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ローカルストレージ設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ローカルストレージから値を削除しました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ローカルストレージから値を削除しました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1127)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0089", "ローカルストレージ削除エラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ローカルストレージ削除エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
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
            if ($global:Common)
            {
                $global:Common.WriteLog("ローカルストレージをクリアしました", "INFO")
                "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] ローカルストレージをクリアしました" | Out-File -Append -FilePath ([WebDriver]::NormalLogFile) -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
        catch
        {
            # ストレージ操作関連エラー (1128)
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("WebDriverError_0089", "ローカルストレージクリアエラー: $($_.Exception.Message)", "WebDriver", [WebDriver]::ErrorLogFile) | Out-Null
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ローカルストレージクリアエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            throw "ローカルストレージのクリアに失敗しました: $($_.Exception.Message)"
        }
    }
}






