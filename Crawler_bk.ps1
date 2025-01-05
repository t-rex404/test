class CDPBrowser {
    [System.Net.WebSockets.ClientWebSocket]$webSocket
    [string]$webSocketDebuggerUrl
    [int]$messageId

    CDPBrowser([string]$debuggerUrl) {
        $this.webSocket = [System.Net.WebSockets.ClientWebSocket]::new()
        $this.webSocketDebuggerUrl = $debuggerUrl
        $this.messageId = 0
        $uri = [System.Uri]::new($this.webSocketDebuggerUrl)
        $this.webSocket.ConnectAsync($uri, [System.Threading.CancellationToken]::None).Wait()
    }

    # WebSocketメッセージ送信
    [void]Send([string]$method, [hashtable]$params = @{}) {
        $this.messageId++
        $message = @{
            id = $this.messageId
            method = $method
            params = $params
        } | ConvertTo-Json -Depth 10
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($message)
        $buffer = [System.ArraySegment[byte]]::new($bytes)
        $this.webSocket.SendAsync($buffer, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [System.Threading.CancellationToken]::None).Wait()
    }

    # WebSocket応答受信
    [string]Receive() {
        $buffer = New-Object byte[] 4096
        $segment = [System.ArraySegment[byte]]::new($buffer)
        $result = $this.webSocket.ReceiveAsync($segment, [System.Threading.CancellationToken]::None).Result
        return [System.Text.Encoding]::UTF8.GetString($buffer, 0, $result.Count)
    }

    # ページ遷移
    [void]NavigateTo([string]$url) {
        $this.Send("Page.navigate", @{ url = $url })
        $this.Receive() | Out-Null
    }

    # 要素を検索
    [string]FindElement([string]$selector) {
        $this.Send("Runtime.evaluate", @{
            expression = "document.querySelector('$selector')"
        })
        $response = $this.Receive() | ConvertFrom-Json
        return $response.result.result.objectId
    }

    # 要素をクリック
    [void]Click([string]$objectId) {
        $this.Send("Runtime.callFunctionOn", @{
            objectId = $objectId
            functionDeclaration = "function() { this.click(); }"
        })
        $this.Receive() | Out-Null
    }

    # テキストを送信
    [void]SendKeys([string]$objectId, [string]$text) {
        $this.Send("Runtime.callFunctionOn", @{
            objectId = $objectId
            functionDeclaration = "function(text) { this.value = text; }"
            arguments = @(@{ value = $text })
        })
        $this.Receive() | Out-Null
    }

    # スクリーンショットを取得
    [void]TakeScreenshot([string]$filePath) {
        $this.Send("Page.captureScreenshot")
        $response = $this.Receive() | ConvertFrom-Json
        $imageData = $response.result.data
        [System.IO.File]::WriteAllBytes($filePath, [Convert]::FromBase64String($imageData))
    }

    # WebSocketを閉じる
    [void]Close() {
        $this.webSocket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "Close", [System.Threading.CancellationToken]::None).Wait()
    }
}
