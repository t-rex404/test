class WebDriver
{
    [int]$browser_exe_process_id
    [System.Net.WebSockets.ClientWebSocket]$web_socket
    [int]$message_id

    WebDriver()
    {
        $this.message_id = 0
    }

    # ブラウザ起動
    [void] StartBrowser([string]$browser_exe_path, [string]$browser_user_data_dir)
    {
        $url = 'about:blank'
        $argument_list = '--new-window ' + $url + ' --remote-debugging-port=9222 --disable-popup-blocking --no-first-run --disable-fre --user-data-dir=' + $browser_user_data_dir
        # ブラウザをデバッグモードで開く
        $browser_exe_process = Start-Process -FilePath $browser_exe_path -ArgumentList $argument_list -PassThru
            # 引数の意味
            # --remote-debugging-port=9222 : デバッグ用 WebSocket をポート 9222 で有効化。
            # --disable-popup-blocking     : パップアップを無効化。
            # --no-first-run               : 最初の起動を無効化。
            # --disable-fre                : フリを無効化。
            # --user-data-dir              : ユーザーデータを指定(別のプロファイルを指定)（既存のブラウザに影響を与えないようにするため）。
        $this.browser_exe_process_id = $browser_exe_process.Id
    }

    # タブ情報を取得
    [Object] GetTabInfomation()
    {
        # デバッグ対象のWebSocket URLを取得(タブ情報を取得)
        $tabs = Invoke-RestMethod -Uri 'http://localhost:9222/json' -Erroraction Stop
        if (-not $tabs) { throw 'タブ情報を取得できません。' }

        $tab = $null
        # 「about:blank」を選択
        foreach ($tab in $tabs)
        {
            if ($tab.title -like 'about:blank' -and $tab.type -like 'page')
            {
                return $tab
                break
            }
        }
        throw 'タブ情報を取得できません。'
    }
    
    # WebSocket接続
    [void] GetWebSocketInfomation([string]$web_socket_debugger_url)
    {
        $retry_count = 3
        $retry_delay = 2  # 秒
        for ($i = 0; $i -lt $retry_count; $i++)
        {
            try
            {
                # WebSocket接続の準備
                $this.web_socket = [System.Net.WebSockets.ClientWebSocket]::new()
                $uri = [System.Uri]::new($web_socket_debugger_url)
                $this.web_socket.ConnectAsync($uri, [System.Threading.CancellationToken]::None).Wait()
                break
            }
            catch
            {
                if ($i -lt $retry_count - 1)
                {
                    Write-Host "WebSocket接続に失敗しました。再試行します... ($($i+1)/$retry_count)"
                    Start-Sleep -Seconds $retry_delay
                }
                else
                {
                    throw 'WebSocket接続に失敗しました。'
                }
            }
        }
    }

    # WebSocketメッセージ送信
    [void] SendWebSocketMessage([string]$method, [hashtable]$params)
    {
        try
        {
            if ($this.web_socket.State -ne [System.Net.WebSockets.WebSocketState]::Open)
            {
                throw 'WebSocketが切断されています。Edgeの状態を確認してください。'
            }
            
            # WebSocketメッセージ送信メッセージ作成
            $this.message_id++
            $message = @{
                id = $this.message_id
                method = $method
                params = $params
            } | ConvertTo-Json -Depth 10 -Compress
            #Write-Host '>>>>>>送信内容:'$message
            # WebSocketメッセージ送信
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($message)
            $buffer = [System.ArraySegment[byte]]::new($bytes)
            $this.web_socket.SendAsync($buffer, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [System.Threading.CancellationToken]::None).Wait()
        }
        catch
        {
            #Write-Host 'SendWebSocketMessage_catch'
            LogError 1004 "WebSocket メッセージ送信エラー: $($_.Exception.Message)"
            #throw 'WebSocketメッセージ送信に失敗しました。エラーメッセージ：' + $_
            throw
        }
    }

    # WebSocketメッセージ受信
    [string] ReceiveWebSocketMessage()
    {
        try
        {
            if ($this.web_socket.State -ne [System.Net.WebSockets.WebSocketState]::Open)
            {
                #throw 'WebSocketは開いていません。'
                throw 'WebSocketが切断されています。Edgeの状態を確認してください。'
            }
            $timeout = [datetime]::Now.AddSeconds(10)

            # WebSocketメッセージ受信
            $responce_result = $false
            $response_json   = $null
            while(-not($responce_result))
            {
                if ([datetime]::Now -gt $timeout)
                {
                    #LogError 9999, 'WebSocketメッセージの受信がタイムアウトしました。'
                    throw 'WebSocketメッセージの受信がタイムアウトしました。'
                }
                #$buffer_size = 1024
                #$buffer_size = 2048
                $buffer_size = 4096
                #$buffer_size = 8192
                #$buffer_size = 16384
                #$buffer_size = 32768
                #$buffer_size = 65536

                $buffer = New-Object byte[] $buffer_size
                $message_stream = New-Object System.IO.MemoryStream
                do
                {
                    $segment = [System.ArraySegment[byte]]::new($buffer)
                    $receive_task = $this.web_socket.ReceiveAsync($segment, [System.Threading.CancellationToken]::None)
                    $receive_task.Wait()
                    $result = $receive_task.Result
                    $message_stream.Write($buffer, 0, $result.Count)
                } while (-not $result.EndOfMessage)
                    
                # UTF8文字列に変換
                $message_stream.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
                $reader = New-Object System.IO.StreamReader($message_stream, [System.Text.Encoding]::UTF8)
                $response_json = $reader.ReadToEnd()
                #Write-Host '<<<<<<受信内容:'$response_json

                try
                {
                    $response_json_object = ConvertFrom-Json $response_json
                    # 受信時には、送信に対する受信だけでなく、ブラウザからのイベント通知も受信する可能性がある
                    # イベント通知の場合、response_jsonにidが含まれないため、idの有無で送信に対する受信なのかイベントの通知なのかを判断する
                    # 注意：送信に成功しても、もし送信したJsonに不備がある場合、response_jsonにidが含まれない
                    if ($response_json_object.id -eq $this.message_id)
                    {
                        return $response_json
                    }
                }
                catch
                {
                    write-host '受信内容をJSON化できませんでした。'$response_json
                }
            }
        }
        catch
        {
            Write-Host 'ReceiveWebSocketMessage_catch'
            $response_json = $null
            LogError 1005 "WebSocket メッセージ受信エラー: $($_.Exception.Message)"
            #throw 'WebSocketメッセージの受信に失敗しました。エラーメッセージ：' + $_
            throw
        }
        return $response_json
    }

    # ページ遷移
    [void] Navigate([string]$url)
    {
        # ページ遷移リクエスト送信
        $this.SendWebSocketMessage('Page.navigate', @{ url = $url })
        $this.ReceiveWebSocketMessage() | Out-Null

        # ページロード完了待機
        $timeout = [datetime]::Now.AddSeconds(60)
        try
        {
            while ($true)
            {
                if ([datetime]::Now -gt $timeout)
                {
                    throw 'ページロード完了待機タイムアウト'
                }
                $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.readyState;"; returnByValue = $false; })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                if ($response_json.result.result.value -eq 'complete')
                {
                    break
                }
                Start-Sleep -Milliseconds 500
            }
        }
        catch
        {
            throw 'ページロード完了待機に失敗しました。エラーメッセージ：' + $_
        }

        # ページイベントの有効化
        $this.EnablePageEvents()

        # 利用可能なタブ情報を取得
        $this.DiscoverTargets()

        # JavaScript実行環境を有効化し、`document` オブジェクトにアクセスできるようにする
        # 結果の返却は不要であるため、returnByValueをfalseに設定
        $this.SendWebSocketMessage('Runtime.enable', @{ expression = 'document;'; returnByValue = $false; })
        $this.ReceiveWebSocketMessage() | Out-Null

        # 現在のページのDOMツリーを取得し、documentノード直下の要素まで取得
        # depth = 1 は、最上位の子要素（例えば<html>や<body>）までを対象とする
        $this.SendWebSocketMessage('DOM.getDocument', @{ depth = 1; })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウを閉じる
    [void] CloseWindow()
    {
        $this.SendWebSocketMessage('Browser.close', @{ })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # WebSocketを閉じる
    [void] Dispose()
    {
        if ($this.web_socket -and $this.web_socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
            $this.web_socket.CloseAsync('NormalClosure', 'Closing', [System.Threading.CancellationToken]::None).Wait()
            $this.web_socket.Dispose()
        }
        if ($this.edge_exe_process_id) {
            Stop-Process -Id $this.edge_exe_process_id -Force
        }
    }

    # 新しいタブやページが開かれた際に、そのターゲット情報を自動で検出できるようにする
    # discover = $true でターゲットの発見機能を有効にする
    [string] DiscoverTargets()
    {
        $this.SendWebSocketMessage('Target.setDiscoverTargets', @{ discover = $true })
        return ($this.ReceiveWebSocketMessage() | ConvertFrom-Json).result.targetId
    }
    
    # 現在開かれているすべてのターゲット（タブやページ）の情報を取得
    # ターゲットには、タブのID、URL、タイトル、タイプなどが含まれる
    [hashtable] GetAvailableTabs()
    {
        $this.SendWebSocketMessage('Target.getTargets', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result
    }

    # タブをアクティブにする
    [void] SetActiveTab($tab_id)
    {
        $this.SendWebSocketMessage('Target.attachToTarget', @{ targetId = $tab_id })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # タブを閉じる
    [void] CloseTab($tab_id)
    {
        $this.SendWebSocketMessage('Target.detachFromTarget', @{ targetId = $tab_id })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ページイベント有効化
    [void] EnablePageEvents()
    {
        $this.SendWebSocketMessage('Page.enable', @{ })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # CSSセレクタで要素を検索
    [hashtable] FindElement([string]$selector)
    {
        $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "document.querySelector('$selector')" })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        if ($response_json.result.result.objectId -eq $null)
        {
            return @{ nodeId = $response_json.result.result.objectId; selector = $selector }
        }
        else
        {
            throw 'CSSセレクタで要素を取得できません。セレクタ：' + $selector
        }
    }

    # CSSセレクタで複数の要素を検索
    [array] FindElements([string]$selector)
    {
        $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = "...document.querySelectorAll('$selector').map(e => e.outerHTML)" })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        if ($response_json.result.result.objectIds.Count -eq 0)
        {
            return @{ nodeId = $response_json.result.result.objectIds; selector = $selector }
        }
        else
        {
            throw 'CSSセレクタで複数の要素を取得できません。セレクタ：' + $selector
        }
    }

    # JavaScriptで要素を検索
    [hashtable] FindElementGeneric([string]$expression, [string]$query_type, [string]$element)
    {
        $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        if ($response_json.result.result.objectId)
        {
            return @{ nodeId = $response_json.result.result.objectId; query_type = $query_type; element = $element } 
        }
        else
        {
            throw "$query_type による要素取得に失敗しました。element：$element"
        }
    }

    # XPathで要素を検索
    [hashtable] FindElementByXPath([string]$xpath)
    {
        $expression = "document.evaluate('$xpath', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
        return $this.FindElementGeneric($expression, 'XPath', $xpath)
    }

    # class属性で要素を検索
    [hashtable] FindElementByClassName([string]$class_name, [int]$index)
    {
        $expression = "document.getElementsByClassName('$class_name')[$index]"
        return $this.FindElementGeneric($expression, 'ClassName', $class_name)
    }

    # id属性で要素を検索
    [hashtable] FindElementById([string]$id)
    {
        $expression = "document.getElementById('$id')"
        return $this.FindElementGeneric($expression, 'Id', $id)
    }

    # name属性で要素を検索
    [hashtable] FindElementByName([string]$name, [int]$index)
    {
        $expression = "document.getElementsByName('$name')[$index]"
        return $this.FindElementGeneric($expression, 'Name', $name)
    }

    # tag名で要素を検索
    [hashtable] FindElementByTagName([string]$tag_name, [int]$index)
    {
        $expression = "document.getElementsByTagName('$tag_name')[$index]"
        return $this.FindElementGeneric($expression, 'TagName', $tag_name)
    }

    # JavaScriptで複数の要素を検索
    [hashtable] FindElementsGeneric([string]$expression, [string]$query_type, [string]$element)
    {
        $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        if ($response_json.result.result.value -gt 0)
        {
            return @{ count = $response_json.result.result.value; query_type = $query_type; element = $element }
        }
        else
        {
            throw "$query_type による複数の要素取得に失敗しました。element：$element"
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
        $expression = "document.getElementsByClassName('$class_name').length"
        $element_count = $this.FindElementsGeneric($expression, 'ClassName', $class_name).count

        $element_list = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $element_count; $i++)
        {
            $element_list.add($this.FindElementByClassName($class_name, $i))
        }
        return $element_list
    }

    # name属性で複数の要素を検索
    [array] FindElementsByName([string]$name)
    {
        $expression = "document.getElementsByName('$name').length"
        $element_count = $this.FindElementsGeneric($expression, 'Name', $name).count

        $element_list = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $element_count; $i++)
        {
            $element_list.add($this.FindElementByName($name, $i))
        }
        return $element_list
    }

    # tag名で複数の要素を検索
    [array] FindElementsByTagName([string]$tag_name)
    {
        $expression = "document.getElementsByTagName('$tag_name').length"
        $element_count = $this.FindElementsGeneric($expression, 'TagName', $tag_name).count        
        
        $element_list = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $element_count; $i++)
        {
            $element_list.add($this.FindElementByTagName($tag_name, $i))
        }
        return $element_list
    }

    <#↓↓↓エレメントの存在確認作成中↓↓↓#>
    # JavaScriptで要素有無を検索
    [bool] IsExistsElementGeneric([string]$expression, [string]$query_type, [string]$element)
    {
        $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        if ($response_json.result.result.objectId)
        {
            return $true
        }
        else
        {
            return $false
        }
    }

    # XPathで要素有無を検索
    [bool] IsExistsElementByXPath([string]$xpath)
    {
        $expression = "document.evaluate('$xpath', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
        return $this.IsExistsElementGeneric($expression, 'XPath', $xpath)
    }

    # class属性で要素有無を検索
    [bool] IsExistsElementByClassName([string]$class_name, [int]$index)
    {
        $expression = "document.getElementsByClassName('$class_name')[$index]"
        return $this.IsExistsElementGeneric($expression, 'ClassName', $class_name)
    }

    # id属性で要素有無を検索
    [bool] IsExistsElementById([string]$id)
    {
        $expression = "document.getElementById('$id')"
        return $this.IsExistsElementGeneric($expression, 'Id', $id)
    }

    # name属性で要素有無を検索
    [bool] IsExistsElementByName([string]$name, [int]$index)
    {
        $expression = "document.getElementsByName('$name')[$index]"
        return $this.IsExistsElementGeneric($expression, 'Name', $name)
    }

    # tag名で要素有無を検索
    [bool] IsExistsElementByTagName([string]$tag_name, [int]$index)
    {
        $expression = "document.getElementsByTagName('$tag_name')[$index]"
        return $this.IsExistsElementGeneric($expression, 'TagName', $tag_name)
    }

    <#↑↑↑作成中↑↑↑#>





    # 要素のテキストを取得
    [string] GetElementText([hashtable]$element)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function() { return this.textContent; }" })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.result.value
    }

    # 要素にテキストを入力
    [void] SetElementText([hashtable]$element, [string]$text)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function(value) { this.value = value; this.dispatchEvent(new Event('input')); }"; arguments = @(@{ value = $text }) })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # 要素をクリック
    [void] ClickElement([hashtable]$element)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function() { this.click(); }" })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # 要素の属性を取得
    [string] GetElementAttribute([hashtable]$element, [string]$attribute)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function(name) { return this.getAttribute(name); }"; arguments = @(@{ value = $attribute }) })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.result.value
    }

    # 要素の属性を設定
    [void] SetElementAttribute([hashtable]$element, [string]$attribute, [string]$value)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function(name, value) { this.setAttribute(name, value); }"; arguments = @(@{ value = $attribute }, @{ value = $value }) })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # アンカータグのhrefを取得
    [string] GetHrefFromAnchor([hashtable]$element)
    {
    $functionDeclaration = "function() { return this.href; }"
    $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = $functionDeclaration; returnByValue = $true })
    $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
    return $response.result.result.value
    }

    # 要素のCSSプロパティを取得
    [string] GetElementCssProperty([hashtable]$element, [string]$property)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function(name) { return window.getComputedStyle(this).getPropertyValue(name); }"; arguments = @(@{ value = $property }) })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.result.value
    }

    # 要素のCSSプロパティを設定
    [void] SetElementCssProperty([hashtable]$element, [string]$property, [string]$value)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function(name, value) { this.style.setProperty(name, value); }"; arguments = @(@{ value = $property }, @{ value = $value }) })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # セレクトタグのオプションをインデックス番号から選択
    [void] SelectOptionByIndex([hashtable]$element, [int]$index)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function(index) { this.selectedIndex = index; this.dispatchEvent(new Event('change')); }"; arguments = @(@{ value = $index }) })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # セレクトタグのオプションをテキストを指定して選択
    [void] SelectOptionByText([hashtable]$element, [string]$text)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function(text) { for (let option of this.options) { if (option.text === text) { option.selected = true; this.dispatchEvent(new Event('change')); break; } } }"; arguments = @(@{ value = $text }) })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # セレクトタグを未選択にする
    [void] DeselectAllOptions([hashtable]$element)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function() { for (let option of this.options) { option.selected = false; } this.dispatchEvent(new Event('change')); }" })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # 要素に入力された値をクリア
    [void] ClearElement([hashtable]$element)
    {
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function() { this.value = ''; }" })
        $this.ReceiveWebSocketMessage() | Out-Null
    }


    # ウィンドウサイズの変更
    [void] ResizeWindow([int]$width, [int]$height, [int]$windowHandle)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{width = $width; height = $height} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(通常)
    [void] NormalWindow([int]$windowHandle)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'normal'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(最大化)
    [void] MaximizeWindow([int]$windowHandle)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'maximized'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(最小化)
    [void] MinimizeWindow([int]$windowHandle)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'minimized'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(フルスクリーン)
    [void] FullscreenWindow([int]$windowHandle)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{windowState = 'fullscreen'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウ位置の変更
    [void] MoveWindow([int]$x, [int]$y, [int]$windowHandle)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = $windowHandle; bounds = @{left = $x; top = $y} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ブラウザの履歴中の前のページに戻る
    [void] GoBack()
    {
        $this.SendWebSocketMessage('Page.navigateToHistoryEntry', @{ entryId = -1 })
        $this.ReceiveWebSocketMessage() | Out-Null
    }
    # ブラウザの履歴中の次のページに進む
    [void] GoForward()
    {
        $this.SendWebSocketMessage('Page.navigateToHistoryEntry', @{ entryId = 1 })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ブラウザを更新
    [void] Refresh()
    {
        $this.SendWebSocketMessage('Page.reload', @{ ignoreCache = $true })
        $this.ReceiveWebSocketMessage() | Out-Null
    }
    # URLを取得
    [string] GetUrl()
    {
        $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.entries[0].url
    }

    # タイトルを取得
    [string] GetTitle()
    {
        $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.entries[0].title
    }

    # ソースコードを取得
    [string] GetSourceCode()
    {
        $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.entries[0].url
    }

    # ウィンドウハンドルを取得
    [int] GetWindowHandle()
    {
        $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.entries[0].windowId
    }

    # 複数のウィンドウハンドルを取得
    [array] GetWindowHandles()
    {
        $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.entries[0].windowId
    }

    # ウィンドウサイズを取得
    [hashtable] GetWindowSize()
    {
        $this.SendWebSocketMessage('Page.getNavigationHistory', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.entries[0].windowId
    }

    # 作成中
    # スクリーンショットを取得（フルスクリーンショット、アクティブウィンドウスクリーンショット、フルページスクリーンショット、ビューポートスクリーンショット、指定要素のスクリーンショット）
    [void] GetScreenshot([string]$type, [string]$save_path)
    {
        switch ($type)
        {
            <#
            'fullScreen'
            {
                # フルスクリーンキャプチャはOSベースで取得
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing
                $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
                $bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
                $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
                $graphics.CopyFromScreen(0, 0, 0, 0, $screen.Size)
                $bitmap.Save($save_path, [System.Drawing.Imaging.ImageFormat]::Png)
                Write-Host 'フルスクリーンショットを保存しました：' $save_path

                # リソースを解放
                $bitmap.Dispose()
                
                #$this.SendWebSocketMessage('Page.captureScreenshot', @{ }) 
                #$response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                #return $response_json.result.data
            }
            #>
            <#
            'active'
            {
                Add-Type @"
                using System;
                using System.Runtime.InteropServices;
                using System.Drawing;

                public class Capture
                {
                    [DllImport("user32.dll")]
                    public static extern IntPtr GetForegroundWindow();

                    [DllImport("user32.dll")]
                    public static extern bool GetWindowRect(IntPtr hwnd, ref RECT rect);

                    [DllImport("user32.dll")]
                    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, int nFlags);

                    [StructLayout(LayoutKind.Sequential)]
                    public struct RECT
                    {
                        public int Left;
                        public int Top;
                        public int Right;
                        public int Bottom;
                    }

                    public static Bitmap CaptureActiveWindow()
                    {
                        IntPtr hwnd = GetForegroundWindow();
                        RECT rect = new RECT();
                        GetWindowRect(hwnd, ref rect);

                        int width = rect.Right - rect.Left;
                        int height = rect.Bottom - rect.Top;

                        Bitmap bmp = new Bitmap(width, height);
                        Graphics g = Graphics.FromImage(bmp);
                        IntPtr hdc = g.GetHdc();
                        PrintWindow(hwnd, hdc, 0);
                        g.ReleaseHdc(hdc);
                        g.Dispose();

                        return bmp;
                    }
                }
"@

                # アクティブウィンドウのスクリーンショットをキャプチャ
                $bitmap = [Capture]::CaptureActiveWindow()

                # スクリーンショットをファイルに保存
                $bitmap.Save($save_path, [System.Drawing.Imaging.ImageFormat]::Png)

                # リソースを解放
                $bitmap.Dispose()

            }
            #>
            'fullPage'
            {
                # ページ全体のスクリーンショット

                $this.SendWebSocketMessage('Page.captureScreenshot', @{ format = 'png'; quality = 100; captureBeyondViewport = $true })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

                # 受信したBase64エンコードされた画像データをBase64デコードして保存
                $image_data_base64 = $response_json.result.data

                # Base64エンコードされた画像データをバイト配列に変換
                $image_data_bytes = [System.Convert]::FromBase64String($image_data_base64)

                # バイト配列をファイルに書き込む
                [System.IO.File]::WriteAllBytes($save_path, $image_data_bytes)
            }
            'viewPort'
            {
                $this.SendWebSocketMessage('Page.captureScreenshot', @{ format = 'png'; quality = 100; captureBeyondViewport = $false })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

                # 受信したBase64エンコードされた画像データをBase64デコードして保存
                $image_data_base64 = $response_json.result.data

                # Base64エンコードされた画像データをバイト配列に変換
                $image_data_bytes = [System.Convert]::FromBase64String($image_data_base64)

                # バイト配列をファイルに書き込む
                [System.IO.File]::WriteAllBytes($save_path, $image_data_bytes)
            }
            <#
            'element'
            {
                $this.SendWebSocketMessage('Page.captureScreenshot', @{ })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                return $response_json.result.data
            }
            #>
            default
            {
                Write-Error 'スクリーンショットのタイプを指定してください。'
            }
        }
        
    }

    # スクリーンショットを取得（指定要素のスクリーンショット）
    [void] GetScreenshotObjectId([hashtable]$element, [string]$save_path)
    {
        # スクロールして要素を可視化（必要に応じて）
        $this.SendWebSocketMessage('Runtime.callFunctionOn', @{ objectId = $element.nodeId; functionDeclaration = "function() { this.scrollIntoViewIfNeeded(true); }" })
        $this.ReceiveWebSocketMessage() | Out-Null

        # 要素の位置とサイズを取得
        $this.SendWebSocketMessage('DOM.getBoxModel', @{ objectId = $element.nodeId })
        $response = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

        if (-not $response.result)
        {
            throw '要素の位置とサイズを取得できません。'
        }

        $content_box = $response.result.model.content
        <#
        $x = [Math]::Min($content_box[0], $content_box[2], $content_box[4], $content_box[6])
        $y = [Math]::Min($content_box[1], $content_box[3], $content_box[5], $content_box[7])
        $width = [Math]::Abs($content_box[2] - $content_box[0])
        $height = [Math]::Abs($content_box[3] - $content_box[1])
        #>

        # 座標をグループ化（x, yペアで分ける）
        $points = @()
        for ($i = 0; $i -lt $content_box.count; $i += 2)
        {
            $points += ,@($content_box[$i], $content_box[$i + 1])
        }

        # 最小x/y、最大x/yを使って矩形を算出（4点でも8点でも対応）
        $x_list = $points | ForEach-Object { $_[0] }
        $y_list = $points | ForEach-Object { $_[1] }

        $x = [Math]::Floor(($x_list | Measure-Object -Minimum).Minimum)
        $y = [Math]::Floor(($y_list | Measure-Object -Minimum).Minimum)
        $x_max = [Math]::Ceiling(($x_list | Measure-Object -Maximum).Maximum)
        $y_max = [Math]::Ceiling(($y_list | Measure-Object -Maximum).Maximum)

        $width  = $x_max - $x
        $height = $y_max - $y

        # スクリーンショットをClip付きで取得
        $this.SendWebSocketMessage('Page.captureScreenshot', @{ format = 'png'; quality = 100; clip = @{ x = $x; y = $y; width = $width; height = $height; scale = 1 } })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json

        # 受信したBase64エンコードされた画像データをBase64デコードして保存
        $image_data_base64 = $response_json.result.data

        # Base64エンコードされた画像データをバイト配列に変換
        $image_data_bytes = [System.Convert]::FromBase64String($image_data_base64)

        # バイト配列をファイルに書き込む
        [System.IO.File]::WriteAllBytes($save_path, $image_data_bytes)
    }
    # 
    # 
    # 
    # 
    # 
    # 
    # 
    # 
    # 
    # 
    # 
        
}

function LogError([string]$error_code, [string]$message)
{
    #$log_file = 'C:\temp\EdgeDriver_Error.log'
    $log_file = '.\EdgeDriver_Error.log'
    $error_message = '[' + $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss') + '], ERROR_CODE:' + $error_code + ', ERROR_MESSAGE:' + $message
    #Add-Content -Path $log_file -Value $error_message
    $error_message | Out-File -Append -FilePath $log_file
    Write-Error $error_message
}
