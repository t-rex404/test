class CDPAutomation {
    [System.Net.WebSockets.ClientWebSocket]$WebSocket
    [string]$WebSocketDebuggerUrl
    [int]$MessageId

    CDPAutomation([string]$debuggerUrl) {
        $this.webSocket = [System.Net.WebSockets.ClientWebSocket]::new()
        $this.webSocketDebuggerUrl = $debuggerUrl
        $this.messageId = 0
        $uri = [System.Uri]::new($this.webSocketDebuggerUrl)
        $this.webSocket.ConnectAsync($uri, [System.Threading.CancellationToken]::None).Wait()
    }

    [void] Connect() {
        $uri = [System.Uri]$this.WebSocketDebuggerUrl
        $this.WebSocket.ConnectAsync($uri, [Threading.CancellationToken]::None).Wait()
    }

    [void] Disconnect() {
        $this.WebSocket.CloseAsync('NormalClosure', 'Closing', [Threading.CancellationToken]::None).Wait()
    }

    [string] SendCommand([string]$method, [hashtable]$params = @{}) {
        # WebSocketメッセージ送信
        $this.MessageId++
        $message = @{
            id = $this.MessageId
            method = $method
            params = $params
        } | ConvertTo-Json -Depth 10
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($message)
        $buffer = [System.ArraySegment[byte]]::new($bytes)
        $this.webSocket.SendAsync($buffer, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [System.Threading.CancellationToken]::None).Wait()

        # WebSocket応答受信
        $responseBuffer = New-Object byte[] 4096
        $responseSegment = [System.ArraySegment[byte]]::new($responseBuffer)
        $receiveTask = $this.WebSocket.ReceiveAsync($responseSegment, [Threading.CancellationToken]::None)
        $receiveTask.Wait()
        $responseJson = [System.Text.Encoding]::UTF8.GetString($responseBuffer, 0, $receiveTask.Result.Count)
        return $responseJson
    }

    [void] Navigate([string]$url) {
        $this.SendCommand("Page.navigate", @{ url = $url })
    }

    [void] EnablePageEvents() {
        $this.SendCommand("Page.enable", @{})
    }

    [hashtable] FindElementById([string]$id) {
        $selector = "#$id"
        return $this.FindElement($selector)
    }

    [hashtable] FindElementByName([string]$name) {
        $selector = "[name='$name']"
        return $this.FindElement($selector)
    }

    [hashtable] FindElementByXPath([string]$xpath) {
        $response = $this.SendCommand("DOM.getDocument", @{})
        $rootNodeId = ($response | ConvertFrom-Json).result.root.nodeId

        $searchResponse = $this.SendCommand("DOM.performSearch", @{
            query = $xpath
        })

        $searchResults = ($searchResponse | ConvertFrom-Json).result.nodeIds
        if ($searchResults.Count -gt 0) {
            return @{ nodeId = $searchResults[0]; xpath = $xpath }
        } else {
            return $null
        }
    }

    [array] FindElementsByTag([string]$tagName) {
        $selector = $tagName
        return $this.FindElements($selector)
    }

    [array] FindElementsByClassName([string]$className) {
        $selector = ".$className"
        return $this.FindElements($selector)
    }

    [array] FindElementsByXPath([string]$xpath) {
        $response = $this.SendCommand("DOM.getDocument", @{})
        $rootNodeId = ($response | ConvertFrom-Json).result.root.nodeId

        $searchResponse = $this.SendCommand("DOM.performSearch", @{
            query = $xpath
        })

        $searchResults = ($searchResponse | ConvertFrom-Json).result.nodeIds
        if ($searchResults.Count -gt 0) {
            return $searchResults | ForEach-Object { @{ nodeId = $_; xpath = $xpath } }
        } else {
            return @()
        }
    }
}
