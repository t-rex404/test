# クラスファイルをインポート
. ".\_ps\_lib\Crawler.ps1"

# Microsoft Edgeを起動し、デバッガーポートを指定 (9222番ポートに設定)
Start-Process "msedge.exe" -ArgumentList "--remote-debugging-port=9222 --disable-popup-blocking --no-first-run --disable-fre --user-data-dir=C:\temp\test\"
Start-Sleep -Seconds 3 # 起動待機

# WebSocket Debugger URLを取得
$debuggerUrlJson = Invoke-RestMethod -Uri "http://localhost:9222/json"
$webSocketDebuggerUrl = $debuggerUrlJson[0].webSocketDebuggerUrl

# CDPAutomationクラスのインスタンスを作成
$browser = [CDPAutomation]::new($webSocketDebuggerUrl)
#$browser.Connect()

## ページを有効化
#$browser.EnablePageEvents()

# 指定URLに移動
$browser.Navigate("https://www.google.co.jp/")
Write-Host "Googleへ移動しました。"
Start-Sleep -Seconds 3 # ページロード待機

## IDで検索
#$elementById = $browser.FindElementById("example-id")
#if ($elementById) {
#    Write-Host "要素が見つかりました (ID):" $elementById.nodeId
#}

# XPathで検索
$elementByXPath = $browser.FindElementByXPath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[1]/div[2]/textarea")
if ($elementByXPath) {
    Write-Host "要素が見つかりました (XPath):" $elementByXPath.nodeId
}

# タグ名で複数要素を検索
$elementsByTag = $browser.FindElementsByTag("p")
Write-Host "タグ名で見つかった要素の数:" $elementsByTag.Count

# クラス名で複数要素を検索
$elementsByClassName = $browser.FindElementsByClassName("example-class")
Write-Host "クラス名で見つかった要素の数:" $elementsByClassName.Count

# 接続を切断
$browser.Disconnect()

# Microsoft Edgeを終了
Stop-Process -Name "msedge"
