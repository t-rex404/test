# aaa
Write-Host "powershell test aaa"

$hoge = (Get-Location).Path
Write-Host "hoge:$hoge"
$hogegoge = $hoge + "\_ps¥_lib¥Crawler.ps1"
Write-Host "hogegoge:$hogegoge"
Write-Host "hogegoge:C:\000_common\PowerShell\prototype\_ps\_lib\Crawler.ps1"

# 外部ファイルのインクルード
. "C:\000_common\PowerShell\prototype\_ps\_lib\Crawler.ps1"
#$TestObject = New-Object CDPBrowser

# Edgeをデバッグモードで起動
Start-Process "msedge.exe" -ArgumentList "--remote-debugging-port=9222 --disable-popup-blocking --no-first-run --disable-fre --user-data-dir=C:\temp\test\"
Start-Sleep -Seconds 5

# デバッグ対象のWebSocket URLを取得
$tabs = Invoke-RestMethod -Uri "http://localhost:9222/json"
$tab = $tabs[0]  # 最初のタブを選択
$debuggerUrl = $tab.webSocketDebuggerUrl
Write-Host "debuggerUrl:$debuggerUrl"
# CDPBrowserインスタンスを作成
$browser = [CDPBrowser]::new($debuggerUrl)

# 1. ページを開く
$browser.NavigateTo("https://www.google.co.jp/")
Start-Sleep -Seconds 5

# 2. 要素を検索
# $elementId = $browser.FindElement("h1")
$elementId = $browser.FindElement("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]")
Write-Host "elementId:$elementId"
# 3. 要素をクリック（例としてボタン）
$browser.Click($elementId)
# 3. テキストを入力
# $browser.SendKeys($elementId, "test")
Write-Output "テキストを入力しました。"
Start-Sleep -Seconds 5

# 4. スクリーンショットを撮る
$browser.TakeScreenshot("screenshot.png")
Write-Output "スクリーンショットを保存しました。"

# 終了処理
$browser.Close()
Stop-Process -Name "msedge" -Force
