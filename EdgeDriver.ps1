class EdgeDriver
{
    [int]$edge_exe_process_id
    [System.Net.WebSockets.ClientWebSocket]$web_socket
    [int]$message_id

    EdgeDriver()
    {
        # Edgeの実行ファイルのパスを取得
        $edge_exe_reg_key = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\'
        try
        {
            # Edgeの実行ファイルのパスを取得
            $edge_exe_path = Get-ItemPropertyValue -Path $edge_exe_reg_key -Name '(default)'
            if (-not $edge_exe_path)
            {
                throw 'Edge実行ファイルが見つかりませんでした。Edgeのインストール状況を確認してください。'
            }
        }
        catch
        {
            throw 'Edgeのパスが見つかりませんでした。エラーメッセージ：' + $_
        }

        # Edgeをデバッグモードで開く
        $edge_exe_process = Start-Process -FilePath $edge_exe_path -ArgumentList '--remote-debugging-port=9222 --disable-popup-blocking --no-first-run --disable-fre --user-data-dir=C:\temp\UserDataDirectoryForEdge\' -PassThru
            # 引数の意味
            # --remote-debugging-port=9222 : デバッグ用 WebSocket をポート 9222 で有効化。
            # --disable-popup-blocking     : パップアップを無効化。
            # --no-first-run               : 最初の起動を無効化。
            # --disable-fre                : フリを無効化。
            # --user-data-dir              : ユーザーデータを指定。
        $this.edge_exe_process_id = $edge_exe_process.Id

        # デバッグ対象のWebSocket URLを取得(タブ情報を取得)
        $tabs = Invoke-RestMethod -Uri 'http://localhost:9222/json' -Erroraction Stop
        if (-not $tabs) { throw 'タブ情報を取得できません。' }

        $web_socket_debugger_url = ''
        $web_socket_target_id    = ''

        # 「新しいタブ」を選択
        foreach ($tab in $tabs)
        {
            if ($tab.title -like '新しいタブ' -and $tab.type -like 'page')
            {
                $web_socket_debugger_url = $tab.webSocketDebuggerUrl
                $web_socket_target_id    = $tab.id
                break
            }
        }

        if (-not $web_socket_debugger_url -or -not $web_socket_target_id)
        {
            throw 'タブ情報を取得できません。'
        }

        # WebSocket接続の準備
        $this.web_socket = [System.Net.WebSockets.ClientWebSocket]::new()
        $uri = [System.Uri]::new($web_socket_debugger_url)
        $this.web_socket.ConnectAsync($uri, [System.Threading.CancellationToken]::None).Wait()

        $this.message_id = 0

        # 利用可能なタブ情報を取得
        #$this.SendWebSocketMessage('Target.setDiscoverTargets', @{ discover = $true })
        #$this.ReceiveWebSocketMessage()
        $this.DiscoverTargets()

        # タブをアクティブにする
        #$this.SendWebSocketMessage('Target.attachToTarget', @{targetId = $web_socket_target_id})
        #$this.ReceiveWebSocketMessage()
        $this.SetActiveTab($web_socket_target_id)

        # デバッグモードを有効化
        $this.SendWebSocketMessage('Emulation.setEmulatedMedia', @{ media = 'screen' })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # WebSocketメッセージ送信
    [void] SendWebSocketMessage([string]$method, [hashtable]$params = @{})
    {
        try
        {
            #write-host 'SendWebSocketMessage_try'
            # WebSocketメッセージ送信メッセージ作成
            $this.message_id++
            $message = @{
                id = $this.message_id
                method = $method
                params = $params
            } | ConvertTo-Json -Depth 10 -Compress
            write-host '>>>>>>送信内容:'$message
            # WebSocketメッセージ送信
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($message)
            $buffer = [System.ArraySegment[byte]]::new($bytes)
            $this.web_socket.SendAsync($buffer, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [System.Threading.CancellationToken]::None).Wait()
        }
        catch
        {
            Write-Host 'SendWebSocketMessage_catch'
            throw 'WebSocketメッセージ送信に失敗しました。エラーメッセージ：' + $_
        }
    }

    # WebSocketメッセージ受信
    [string] ReceiveWebSocketMessage()
    {
        try
        {
            #write-host 'ReceiveWebSocketMessage_try'
            $timeout = [datetime]::Now.AddSeconds(10)
            # WebSocketメッセージ受信
            $responce_result = $false
            $response_json   = $null
            while(-not($responce_result))
            {
                if ([datetime]::Now -gt $timeout)
                {
                    throw 'WebSocketメッセージの受信がタイムアウトしました。'
                }
                $buffer = New-Object byte[] 4096
                $segment = [System.ArraySegment[byte]]::new($buffer)
                $receive_task = $this.web_socket.ReceiveAsync($segment, [System.Threading.CancellationToken]::None)
                $receive_task.Wait()
                $response_json = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $receive_task.Result.Count)
                write-host '<<<<<<受信内容:'$response_json
                $response_json_object = ConvertFrom-Json $response_json
                # 受信時には、送信に対する受信だけでなく、ブラウザからのイベント通知も受信する可能性がある
                # イベント通知の場合、response_jsonにidが含まれないため、idの有無で送信に対する受信なのかイベントの通知なのかを判断する
                # 注意：送信に成功しても、もし送信したJsonに不備がある場合、response_jsonにidが含まれない
                if ($response_json_object.id -eq $this.message_id)
                {
                    return $response_json
                }
            }
        }
        catch
        {
            Write-Host 'ReceiveWebSocketMessage_catch'
            $response_json = $null
            throw 'WebSocketメッセージの受信に失敗しました。エラーメッセージ：' + $_
        }
        return $response_json
    }

    # ページ遷移
    [void] Navigate([string]$url)
    {
        # ページ遷移リクエスト送信
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
                    throw 'ページロード完了待機タイムアウト'
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
            throw 'ページロード完了待機に失敗しました。エラーメッセージ：' + $_
        }

        # ページイベントの有効化
        $this.EnablePageEvents()

        # 利用可能なタブ情報を取得
        $this.DiscoverTargets()

        $this.SendWebSocketMessage('Runtime.enable', @{ expression = 'document;'; returnByValue = $false; })
        $this.ReceiveWebSocketMessage() | Out-Null

        $this.SendWebSocketMessage('DOM.getDocument', @{ depth = 1; })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウを閉じる
    [void] CloseWindow()
    {
        $this.SendWebSocketMessage('Browser.close', @{})
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # WebSocketを閉じる
    [void] Dispose() {
        if ($this.web_socket -and $this.web_socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
            $this.web_socket.CloseAsync('NormalClosure', 'Closing', [System.Threading.CancellationToken]::None).Wait()
            $this.web_socket.Dispose()
        }
        if ($this.edge_exe_process_id) {
            Stop-Process -Id $this.edge_exe_process_id -Force
        }
    }    
    
    # 利用可能なタブ情報を取得
    [hashtable] GetAvailableTabs()
    {
        $this.SendWebSocketMessage('Target.getTargets', @{})
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result
    }

    # 利用可能なタブ情報を取得
    [void] DiscoverTargets()
    {
        $this.SendWebSocketMessage('Target.setDiscoverTargets', @{ discover = $true })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # タブをアクティブにする
    [void] SetActiveTab($tab_id)
    {
        $this.SendWebSocketMessage('Target.attachToTarget', @{targetId = $tab_id})
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # タブを閉じる
    [void] CloseTab($tab_id)
    {
        $this.SendWebSocketMessage('Target.detachFromTarget', @{targetId = $tab_id})
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ページイベント有効化
    [void] EnablePageEvents() {
        $this.SendWebSocketMessage('Page.enable', @{})
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
        return = $this.FindElementGeneric($expression, 'XPath', $xpath)
    }

    # class属性で要素を検索
    [hashtable] FindElementByClassName([string]$class_name)
    {
        $expression = "document.getElementsByClassName('$class_name')"
        return = $this.FindElementGeneric($expression, 'ClassName', $class_name)
    }

    # id属性で要素を検索
    [hashtable] FindElementById([string]$id)
    {
        $expression = "document.getElementById('$id')"
        return = $this.FindElementGeneric($expression, 'Id', $id)
    }

    # name属性で要素を検索
    [hashtable] FindElementByName([string]$name)
    {
        $expression = "document.getElementsByName('$name')"
        return = $this.FindElementGeneric($expression, 'Name', $name)
    }

    # tag名で要素を検索
    [hashtable] FindElementByTagName([string]$tag_name)
    {
        $expression = "document.getElementsByTagName('$tag_name')"
        return = $this.FindElementGeneric($expression, 'TagName', $tag_name)
    }

    # JavaScriptで複数の要素を検索
    [hashtable] FindElementsGeneric([string]$expression, [string]$query_type, [string]$element)
    {
        $this.SendWebSocketMessage('Runtime.evaluate', @{ expression = $expression })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        if ($response_json.result.result.objectIds.Count -eq 0)
        {
            return @{ nodeId = $response_json.result.result.objectIds; query_type = $query_type; element = $element }
        }
        else
        {
            throw "$query_type による複数の要素取得に失敗しました。element：$element"
        }
    }
    
    # XPathで複数の要素を検索
    [array] FindElementsByXPath([string]$xpath)
    {
        $expression = "...document.evaluate('$xpath', document, null, XPathResult.ANY_TYPE, null).iterateNext().outerHTML"
        return $this.FindElementsGeneric($expression, 'XPath', $xpath)
    }
    # class属性で複数の要素を検索
    [array] FindElementsByClassName([string]$class_name)
    {
        $expression = "...document.getElementsByClassName('$class_name').map(e => e.outerHTML)"
        return $this.FindElementsGeneric($expression, 'ClassName', $class_name)
    }

    # name属性で複数の要素を検索
    [array] FindElementsByName([string]$name)
    {
        $expression = "...document.getElementsByName('$name').map(e => e.outerHTML)"
        return $this.FindElementsGeneric($expression, 'Name', $name)
    }

    # tag名で複数の要素を検索
    [array] FindElementsByTagName([string]$tag_name)
    {
        $expression = "...document.getElementsByTagName('$tag_name').map(e => e.outerHTML)"
        return $this.FindElementsGeneric($expression, 'TagName', $tag_name)
    }

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
    [void] ResizeWindow([int]$width, [int]$height)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = 1; bounds = @{width = $width; height = $height} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(通常)
    [void] NormalWindow()
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = 1; bounds = @{windowState = 'normal'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(最大化)
    [void] MaximizeWindow()
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = 1; bounds = @{windowState = 'maximized'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(最小化)
    [void] MinimizeWindow()
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = 1; bounds = @{windowState = 'minimized'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウサイズの変更(フルスクリーン)
    [void] FullscreenWindow()
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = 1; bounds = @{windowState = 'fullscreen'} })
        $this.ReceiveWebSocketMessage() | Out-Null
    }

    # ウィンドウ位置の変更
    [void] MoveWindow([int]$x, [int]$y)
    {
        $this.SendWebSocketMessage('Browser.setWindowBounds', @{ windowId = 1; bounds = @{left = $x; top = $y} })
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
    [string] GetScreenshot($type = 'fullPage', $element = $null, $save_path = '.\screenshot.png')
    {
        switch ($type)
        {
            'fullPage'
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


                #$this.SendWebSocketMessage('Page.captureScreenshot', @{ }) 
                #$response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                #return $response_json.result.data
            }
            'active'
            {
                $this.SendWebSocketMessage('Page.captureScreenshot', @{ })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                return $response_json.result.data
            }
            'fullPage'
            {
                # ページ全体のスクリーンショット

                $this.SendWebSocketMessage('Page.captureScreenshot', @{ })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                return $response_json.result.data
            }
            'viewPort'
            {
                $this.SendWebSocketMessage('Page.captureScreenshot', @{ })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                return $response_json.result.data
            }
            'element'
            {
                $this.SendWebSocketMessage('Page.captureScreenshot', @{ })
                $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
                return $response_json.result.data
            }
            condition {  }
            Default {}
        }
        $this.SendWebSocketMessage('Page.captureScreenshot', @{ })
        $response_json = $this.ReceiveWebSocketMessage() | ConvertFrom-Json
        return $response_json.result.data
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
    # 
    function Get-ActiveWindowScreenshot ([string]$OutputPath = "ActiveWindowScreenshot.png")
    {
    
        # Windows APIの定義
        Add-Type -Namespace Win32 -Name GDI32 -MemberDefinition @"
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    
        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    
        public struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    "@
    
        Add-Type -Namespace Win32 -Name Drawing -MemberDefinition @"
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
    
        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
    
        [DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);
    
        [DllImport("gdi32.dll")]
        public static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, uint dwRop);
    
        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);
    
        [DllImport("gdi32.dll")]
        public static extern bool DeleteDC(IntPtr hdc);
    
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);
    
        [DllImport("user32.dll")]
        public static extern bool ReleaseDC(IntPtr hwnd, IntPtr hdc);
    
        public const int SRCCOPY = 0x00CC0020;
    "@
    
        # アクティブウィンドウのハンドルを取得
        $hWnd = [Win32.GDI32]::GetForegroundWindow()
    
        if (-not $hWnd) {
            Write-Error "アクティブウィンドウが見つかりませんでした。"
            return
        }
    
        # アクティブウィンドウの位置とサイズを取得
        $rect = New-Object Win32.GDI32+RECT
        [Win32.GDI32]::GetWindowRect($hWnd, [ref]$rect) | Out-Null
    
        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top
    
        if ($width -le 0 -or $height -le 0) {
            Write-Error "ウィンドウサイズが無効です。"
            return
        }
    
        # デバイスコンテキストを取得
        $windowDC = [Win32.Drawing]::GetDC($hWnd)
        $memoryDC = [Win32.Drawing]::CreateCompatibleDC($windowDC)
        $bitmap = [Win32.Drawing]::CreateCompatibleBitmap($windowDC, $width, $height)
        [Win32.Drawing]::SelectObject($memoryDC, $bitmap) | Out-Null
    
        # ウィンドウの内容をキャプチャ
        [Win32.Drawing]::BitBlt($memoryDC, 0, 0, $width, $height, $windowDC, 0, 0, [Win32.Drawing]::SRCCOPY)
    
        # スクリーンショットをファイルに保存
        $image = [System.Drawing.Image]::FromHbitmap($bitmap)
        $image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $image.Dispose()
    
        # リソースを解放
        [Win32.Drawing]::DeleteObject($bitmap)
        [Win32.Drawing]::DeleteDC($memoryDC)
        [Win32.Drawing]::ReleaseDC($hWnd, $windowDC)
    
        Write-Host "アクティブウィンドウのスクリーンショットを保存しました: $OutputPath"
    }
        
}
