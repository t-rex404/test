$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-StepResult {
    param(
        [string]$Name,
        [bool]$Succeeded,
        [string]$Message
    )
    $status = if ($Succeeded) { 'OK' } else { 'NG' }
    Write-Host ("[STEP] {0,-45} : {1} - {2}" -f $Name, $status, $Message)
}

function Invoke-TestStep {
    param(
        [string]$Name,
        [scriptblock]$ScriptBlock
    )
    try {
        $result = & $ScriptBlock
        Write-StepResult -Name $Name -Succeeded $true -Message "Success"
        return $result
    } catch {
        Write-StepResult -Name $Name -Succeeded $false -Message $_.Exception.Message
    }
}

# スクリプトの基準パス
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$LibDir    = Join-Path $ScriptDir '_lib'

# 依存スクリプト読込（Common -> WebDriver -> EdgeDriver の順）
. (Join-Path $LibDir 'Common.ps1')
. (Join-Path $LibDir 'WebDriver.ps1')
. (Join-Path $LibDir 'EdgeDriver.ps1')

# 出力ディレクトリ
$OutDir = Join-Path $RepoRoot '_out'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

# テスト用ページ
$samplePath = Join-Path $RepoRoot 'sample.html'
$sampleUrl  = [System.Uri] $samplePath

# ドライバー生成（EdgeDriver コンストラクタ内で WebDriver 初期化が行われる）
$driver = $null
try {
    $driver = [EdgeDriver]::new()
} catch {
    Write-Error "EdgeDriver 初期化に失敗しました: $($_.Exception.Message)"
    throw
}

try {
    Write-Host "=== WebDriver 全メソッド機能テスト開始 ===" -ForegroundColor Cyan
    
    # ========================================
    # 1. 初期化・接続関連メソッド
    # ========================================
    Write-Host "`n--- 1. 初期化・接続関連メソッド ---" -ForegroundColor Yellow
    
    # ブラウザ起動（EdgeDriverコンストラクタで既に実行済み）
    Invoke-TestStep -Name 'StartBrowser (EdgeDriver経由)' -ScriptBlock { Write-Host "EdgeDriver経由でブラウザが起動済み" }
    
    # WebSocket情報取得
    #Invoke-TestStep -Name 'GetWebSocketInfomation' -ScriptBlock { $driver.GetWebSocketInfomation($driver.web_socket_debugger_url) }
    
    # ターゲット発見
    Invoke-TestStep -Name 'DiscoverTargets' -ScriptBlock { $driver.DiscoverTargets() }
    
    # ページイベント有効化
    Invoke-TestStep -Name 'EnablePageEvents' -ScriptBlock { $driver.EnablePageEvents() }
    
    # ========================================
    # 2. ページ遷移・ナビゲーション関連メソッド
    # ========================================
    Write-Host "`n--- 2. ページ遷移・ナビゲーション関連メソッド ---" -ForegroundColor Yellow
    
    # ページ遷移（file:/// のローカル HTML）
    Invoke-TestStep -Name 'Navigate(sample.html)' -ScriptBlock { $driver.Navigate($sampleUrl.AbsoluteUri) }
    
    # ページロード待機
    Invoke-TestStep -Name 'WaitForPageLoad' -ScriptBlock { $driver.WaitForPageLoad(10) }
    
    # カスタム条件待機
    Invoke-TestStep -Name 'WaitForCondition(document.readyState===complete)' -ScriptBlock { $driver.WaitForCondition('document.readyState === "complete"', 5) }
    
    # 広告読み込み待機（事前に要素を挿入）
    Invoke-TestStep -Name 'ExecuteScript(insert .ad element)' -ScriptBlock { $driver.ExecuteScript("var d=document.createElement('div');d.className='ad';document.body.appendChild(d);") }
    Invoke-TestStep -Name 'WaitForAdToLoad' -ScriptBlock { $driver.WaitForAdToLoad() }
    
    # ========================================
    # 3. 要素検索関連メソッド（単体）
    # ========================================
    Write-Host "`n--- 3. 要素検索関連メソッド（単体） ---" -ForegroundColor Yellow
    
    # 基本要素検索
    $textbox  = Invoke-TestStep -Name 'FindElement(#textbox)'  -ScriptBlock { $driver.FindElement('#textbox') }
    $dropdown = Invoke-TestStep -Name 'FindElement(#dropdown)' -ScriptBlock { $driver.FindElement('#dropdown') }
    $btn      = Invoke-TestStep -Name 'FindElement(#normalButton)' -ScriptBlock { $driver.FindElement('#normalButton') }
    $cbx1     = Invoke-TestStep -Name 'FindElement(#checkbox1)' -ScriptBlock { $driver.FindElement('#checkbox1') }
    $rad1     = Invoke-TestStep -Name 'FindElement(#radio1)' -ScriptBlock { $driver.FindElement('#radio1') }
    $fileIn   = Invoke-TestStep -Name 'FindElement(#file)' -ScriptBlock { $driver.FindElement('#file') }
    
    # 汎用要素検索
    Invoke-TestStep -Name 'FindElementGeneric(document.getElementById)' -ScriptBlock { $driver.FindElementGeneric("document.getElementById('textbox')", 'JS', 'textbox') }
    Invoke-TestStep -Name 'FindElementGeneric(getElementsByTagName)' -ScriptBlock { $driver.FindElementGeneric("document.getElementsByTagName('input')[0]", 'TagName', 'input') }
    
    # 特定タイプ要素検索
    Invoke-TestStep -Name 'FindElementByXPath(//*[@id="textbox"])' -ScriptBlock { $driver.FindElementByXPath('//*[@id="textbox"]') }
    Invoke-TestStep -Name 'FindElementByClassName(input-text,0)' -ScriptBlock { $driver.FindElementByClassName('input-text', 0) }
    Invoke-TestStep -Name 'FindElementByName(textbox,0)' -ScriptBlock { $driver.FindElementByName('textbox', 0) }
    Invoke-TestStep -Name 'FindElementById(textbox)' -ScriptBlock { $driver.FindElementById('textbox') }
    Invoke-TestStep -Name 'FindElementByTagName(input,0)' -ScriptBlock { $driver.FindElementByTagName('input', 0) }
    
    # ========================================
    # 4. 要素検索関連メソッド（複数）
    # ========================================
    Write-Host "`n--- 4. 要素検索関連メソッド（複数） ---" -ForegroundColor Yellow
    
    # 複数要素検索
    Invoke-TestStep -Name 'FindElements(.input-checkbox)' -ScriptBlock { $driver.FindElements('.input-checkbox') }
    Invoke-TestStep -Name 'FindElementsGeneric(getElementsByTagName length)' -ScriptBlock { $driver.FindElementsGeneric("document.getElementsByTagName('input').length", 'TagNameCount', 'input') }
    Invoke-TestStep -Name 'FindElementsByClassName(input-checkbox)' -ScriptBlock { $driver.FindElementsByClassName('input-checkbox') }
    Invoke-TestStep -Name 'FindElementsByName(checkbox1)' -ScriptBlock { $driver.FindElementsByName('checkbox1') }
    Invoke-TestStep -Name 'FindElementsByTagName(input)' -ScriptBlock { $driver.FindElementsByTagName('input') }
    
    # ========================================
    # 5. 要素存在確認関連メソッド
    # ========================================
    Write-Host "`n--- 5. 要素存在確認関連メソッド ---" -ForegroundColor Yellow
    
    # 存在確認（各種）
    Invoke-TestStep -Name 'IsExistsElementGeneric(document.getElementById("textbox"))' -ScriptBlock { $driver.IsExistsElementGeneric("document.getElementById('textbox')", 'JS', 'textbox') }
    Invoke-TestStep -Name 'IsExistsElementByXPath(//*[@id="textbox"])' -ScriptBlock { $driver.IsExistsElementByXPath('//*[@id="textbox"]') }
    Invoke-TestStep -Name 'IsExistsElementByClassName(input-text,0)' -ScriptBlock { $driver.IsExistsElementByClassName('input-text', 0) }
    Invoke-TestStep -Name 'IsExistsElementById(textbox)' -ScriptBlock { $driver.IsExistsElementById('textbox') }
    Invoke-TestStep -Name 'IsExistsElementByName(textbox,0)' -ScriptBlock { $driver.IsExistsElementByName('textbox', 0) }
    Invoke-TestStep -Name 'IsExistsElementByTagName(input,0)' -ScriptBlock { $driver.IsExistsElementByTagName('input', 0) }
    
    # ========================================
    # 6. 要素操作関連メソッド
    # ========================================
    Write-Host "`n--- 6. 要素操作関連メソッド ---" -ForegroundColor Yellow
    
    if ($textbox -and $textbox.nodeId) {
        # テキスト操作
        Invoke-TestStep -Name 'SetElementText(textbox,"Hello World")' -ScriptBlock { $driver.SetElementText($textbox.nodeId, 'Hello World') }
        Invoke-TestStep -Name 'GetElementText(textbox)' -ScriptBlock { $driver.GetElementText($textbox.nodeId) }
        
        # 属性操作
        Invoke-TestStep -Name 'GetElementAttribute(textbox,id)' -ScriptBlock { $driver.GetElementAttribute($textbox.nodeId, 'id') }
        Invoke-TestStep -Name 'SetElementAttribute(textbox,data-test=abc)' -ScriptBlock { $driver.SetElementAttribute($textbox.nodeId, 'data-test', 'abc') }
        
        # CSS操作
        Invoke-TestStep -Name 'GetElementCssProperty(textbox,display)' -ScriptBlock { $driver.GetElementCssProperty($textbox.nodeId, 'display') }
        Invoke-TestStep -Name 'SetElementCssProperty(textbox,backgroundColor=red)' -ScriptBlock { $driver.SetElementCssProperty($textbox.nodeId, 'backgroundColor', 'red') }
        
        # 要素クリア
        Invoke-TestStep -Name 'ClearElement(textbox)' -ScriptBlock { $driver.ClearElement($textbox.nodeId) }
        
        # キーボード入力
        Invoke-TestStep -Name 'SendKeys(textbox,"Abc123")' -ScriptBlock { $driver.SendKeys($textbox.nodeId, 'Abc123') }
        Invoke-TestStep -Name 'SendSpecialKey(textbox,Enter)' -ScriptBlock { $driver.SendSpecialKey($textbox.nodeId, 'Enter') }
    }
    
    # ========================================
    # 7. マウス操作関連メソッド
    # ========================================
    Write-Host "`n--- 7. マウス操作関連メソッド ---" -ForegroundColor Yellow
    
    if ($btn -and $btn.nodeId) {
        Invoke-TestStep -Name 'ClickElement(normalButton)' -ScriptBlock { $driver.ClickElement($btn.nodeId) }
        Invoke-TestStep -Name 'MouseHover(normalButton)' -ScriptBlock { $driver.MouseHover($btn.nodeId) }
        Invoke-TestStep -Name 'DoubleClick(normalButton)' -ScriptBlock { $driver.DoubleClick($btn.nodeId) }
        Invoke-TestStep -Name 'RightClick(normalButton)' -ScriptBlock { $driver.RightClick($btn.nodeId) }
    }
    
    # ========================================
    # 8. フォーム要素操作関連メソッド
    # ========================================
    Write-Host "`n--- 8. フォーム要素操作関連メソッド ---" -ForegroundColor Yellow
    
    if ($dropdown -and $dropdown.nodeId) {
        # セレクト操作
        Invoke-TestStep -Name 'SelectOptionByIndex(dropdown,1)' -ScriptBlock { $driver.SelectOptionByIndex($dropdown.nodeId, 1) }
        Invoke-TestStep -Name 'SelectOptionByText(dropdown,選択肢3)' -ScriptBlock { $driver.SelectOptionByText($dropdown.nodeId, '選択肢3') }
        # multi 属性付与してから全解除
        Invoke-TestStep -Name 'ExecuteScript(set dropdown multiple)' -ScriptBlock { $driver.ExecuteScript("document.getElementById('dropdown').setAttribute('multiple','multiple')") }
        Invoke-TestStep -Name 'DeselectAllOptions(dropdown)' -ScriptBlock { $driver.DeselectAllOptions($dropdown.nodeId) }
    }
    
    # チェック/ラジオ
    Invoke-TestStep -Name 'SetCheckbox(checkbox1,true)' -ScriptBlock { if ($cbx1 -and $cbx1.nodeId) { $driver.SetCheckbox($cbx1.nodeId, $true) } }
    Invoke-TestStep -Name 'SelectRadioButton(radio1)' -ScriptBlock { if ($rad1 -and $rad1.nodeId) { $driver.SelectRadioButton($rad1.nodeId) } }
    
    # ファイルアップロード
    $uploadTarget = $samplePath
    Invoke-TestStep -Name 'UploadFile(file input, sample.html)' -ScriptBlock { if ($fileIn -and $fileIn.nodeId) { $driver.UploadFile($fileIn.nodeId, $uploadTarget) } }
    
    # ========================================
    # 9. 待機関連メソッド
    # ========================================
    Write-Host "`n--- 9. 待機関連メソッド ---" -ForegroundColor Yellow
    
    Invoke-TestStep -Name 'WaitForElementVisible(#textbox)' -ScriptBlock { $driver.WaitForElementVisible('#textbox', 5) }
    Invoke-TestStep -Name 'WaitForElementClickable(#normalButton)' -ScriptBlock { $driver.WaitForElementClickable('#normalButton', 5) }
    
    # ========================================
    # 10. ウィンドウ制御関連メソッド
    # ========================================
    Write-Host "`n--- 10. ウィンドウ制御関連メソッド ---" -ForegroundColor Yellow
    
    $handle = Invoke-TestStep -Name 'GetWindowHandle()' -ScriptBlock { $driver.GetWindowHandle() }
    if ($handle) {
        Invoke-TestStep -Name 'ResizeWindow(1024x768)' -ScriptBlock { $driver.ResizeWindow(1024, 768, $handle) }
        Invoke-TestStep -Name 'NormalWindow()'         -ScriptBlock { $driver.NormalWindow($handle) }
        Invoke-TestStep -Name 'MaximizeWindow()'       -ScriptBlock { $driver.MaximizeWindow($handle) }
        Invoke-TestStep -Name 'NormalWindow()'         -ScriptBlock { $driver.NormalWindow($handle) }
        Invoke-TestStep -Name 'MinimizeWindow()'       -ScriptBlock { $driver.MinimizeWindow($handle) }
        Invoke-TestStep -Name 'NormalWindow()'         -ScriptBlock { $driver.NormalWindow($handle) }
        Invoke-TestStep -Name 'FullscreenWindow()'     -ScriptBlock { $driver.FullscreenWindow($handle) }
        Invoke-TestStep -Name 'NormalWindow()'         -ScriptBlock { $driver.NormalWindow($handle) }
        Invoke-TestStep -Name 'MoveWindow(10,10)'      -ScriptBlock { $driver.MoveWindow(10, 10, $handle) }
    }
    Invoke-TestStep -Name 'GetWindowHandles()' -ScriptBlock { $driver.GetWindowHandles() }
    Invoke-TestStep -Name 'GetWindowSize()'    -ScriptBlock { $driver.GetWindowSize() }
    
    # ========================================
    # 11. ナビゲーション・情報取得関連メソッド
    # ========================================
    Write-Host "`n--- 11. ナビゲーション・情報取得関連メソッド ---" -ForegroundColor Yellow
    
    Invoke-TestStep -Name 'GetUrl()'        -ScriptBlock { $driver.GetUrl() }
    Invoke-TestStep -Name 'GetTitle()'      -ScriptBlock { $driver.GetTitle() }
    Invoke-TestStep -Name 'GetSourceCode()' -ScriptBlock { $driver.GetSourceCode() }
    
    # ========================================
    # 12. スクリーンショット関連メソッド
    # ========================================
    Write-Host "`n--- 12. スクリーンショット関連メソッド ---" -ForegroundColor Yellow
    
    $ss1 = Join-Path $OutDir 'screenshot_viewport.png'
    $ss2 = Join-Path $OutDir 'screenshot_fullpage.png'
    Invoke-TestStep -Name 'GetScreenshot(viewPort)' -ScriptBlock { $driver.GetScreenshot('viewPort', $ss1) }
    Invoke-TestStep -Name 'GetScreenshot(fullPage)' -ScriptBlock { $driver.GetScreenshot('fullPage', $ss2) }
    
    if ($textbox -and $textbox.nodeId) {
        $ssEl1 = Join-Path $OutDir 'screenshot_element_textbox.png'
        Invoke-TestStep -Name 'GetScreenshotObjectId(textbox)' -ScriptBlock { $driver.GetScreenshotObjectId($textbox.nodeId, $ssEl1) }
    }
    
    if ($textbox -and $dropdown -and $textbox.nodeId -and $dropdown.nodeId) {
        $ssEls = Join-Path $OutDir 'screenshot_elements_textbox_dropdown.png'
        Invoke-TestStep -Name 'GetScreenshotObjectIds([textbox,dropdown])' -ScriptBlock { $driver.GetScreenshotObjectIds(@($textbox.nodeId, $dropdown.nodeId), $ssEls) }
    }
    
    # ========================================
    # 13. JavaScript実行関連メソッド
    # ========================================
    Write-Host "`n--- 13. JavaScript実行関連メソッド ---" -ForegroundColor Yellow
    
    Invoke-TestStep -Name 'ExecuteScript(1+1)'       -ScriptBlock { $driver.ExecuteScript('1+1') }
    Invoke-TestStep -Name 'ExecuteScriptAsync(setTimeout...)' -ScriptBlock { $driver.ExecuteScriptAsync('setTimeout(()=>true,100)') }
    
    # ========================================
    # 14. WebSocket通信関連メソッド
    # ========================================
    Write-Host "`n--- 14. WebSocket通信関連メソッド ---" -ForegroundColor Yellow
    
    # 低レベルAPI（Send/Receive）を明示的に実行
    Invoke-TestStep -Name 'SendWebSocketMessage(Runtime.evaluate 2+3)' -ScriptBlock { $driver.SendWebSocketMessage('Runtime.evaluate', @{ expression = '2+3'; returnByValue = $true }) }
    Invoke-TestStep -Name 'ReceiveWebSocketMessage()' -ScriptBlock { $driver.ReceiveWebSocketMessage() }
    
    # ========================================
    # 15. タブ管理関連メソッド
    # ========================================
    Write-Host "`n--- 15. タブ管理関連メソッド ---" -ForegroundColor Yellow
    
    Invoke-TestStep -Name 'GetTabInfomation()' -ScriptBlock { $driver.GetTabInfomation() }
    Invoke-TestStep -Name 'GetAvailableTabs()' -ScriptBlock { $driver.GetAvailableTabs() }
    
    # ========================================
    # 16. セッション管理関連メソッド
    # ========================================
    Write-Host "`n--- 16. セッション管理関連メソッド ---" -ForegroundColor Yellow
    
    $currentTab = $driver.GetTabInfomation()
    if ($currentTab -and $currentTab.id) {
        Invoke-TestStep -Name 'AttachToCurrentTabAndGetSessionId' -ScriptBlock { $driver.AttachToCurrentTabAndGetSessionId() }
        Invoke-TestStep -Name 'AttachToTargetAndGetSessionId' -ScriptBlock { $driver.AttachToTargetAndGetSessionId($currentTab.id) }
    }
    
    # ========================================
    # 17. アンカー要素関連メソッド
    # ========================================
    Write-Host "`n--- 17. アンカー要素関連メソッド ---" -ForegroundColor Yellow
    
    # アンカー要素を動的追加して href 取得をテスト
    Invoke-TestStep -Name 'ExecuteScript(insert <a id=alink>)' -ScriptBlock { $driver.ExecuteScript("var a=document.createElement('a');a.id='alink';a.href='https://example.com';a.text='link';document.body.appendChild(a);") }
    $alink = Invoke-TestStep -Name 'FindElement(#alink)' -ScriptBlock { $driver.FindElement('#alink') }
    Invoke-TestStep -Name 'GetHrefFromAnchor(#alink)' -ScriptBlock { if ($alink -and $alink.nodeId) { $driver.GetHrefFromAnchor($alink.nodeId) } }
    
    # ========================================
    # 18. Cookie/LocalStorage関連メソッド
    # ========================================
    Write-Host "`n--- 18. Cookie/LocalStorage関連メソッド ---" -ForegroundColor Yellow
    
    # Cookie/LocalStorage は https ドメインで
    Invoke-TestStep -Name 'Navigate(https://example.com)' -ScriptBlock { $driver.Navigate('https://example.com') }
    Invoke-TestStep -Name 'SetCookie(test,123,example.com)' -ScriptBlock { $driver.SetCookie('test','123','example.com','/', 1) }
    Invoke-TestStep -Name 'GetCookie(test)' -ScriptBlock { $driver.GetCookie('test') }
    Invoke-TestStep -Name 'DeleteCookie(test)' -ScriptBlock { $driver.DeleteCookie('test') }
    Invoke-TestStep -Name 'ClearAllCookies()' -ScriptBlock { $driver.ClearAllCookies() }
    
    Invoke-TestStep -Name 'SetLocalStorage(k,v)' -ScriptBlock { $driver.SetLocalStorage('k','v') }
    Invoke-TestStep -Name 'GetLocalStorage(k)'   -ScriptBlock { $driver.GetLocalStorage('k') }
    Invoke-TestStep -Name 'RemoveLocalStorage(k)' -ScriptBlock { $driver.RemoveLocalStorage('k') }
    Invoke-TestStep -Name 'ClearLocalStorage()'  -ScriptBlock { $driver.ClearLocalStorage() }
    
    # ========================================
    # 19. ブラウザ履歴関連メソッド
    # ========================================
    Write-Host "`n--- 19. ブラウザ履歴関連メソッド ---" -ForegroundColor Yellow
    
    Invoke-TestStep -Name 'Navigate(sample.html again)' -ScriptBlock { $driver.Navigate($sampleUrl.AbsoluteUri) }
    Invoke-TestStep -Name 'GoBack()'    -ScriptBlock { $driver.GoBack() }
    Invoke-TestStep -Name 'GoForward()' -ScriptBlock { $driver.GoForward() }
    Invoke-TestStep -Name 'Refresh()'   -ScriptBlock { $driver.Refresh() }
    
    # ========================================
    # 20. タブ操作関連メソッド（最後に実行）
    # ========================================
    Write-Host "`n--- 20. タブ操作関連メソッド ---" -ForegroundColor Yellow
    
    $existsOpenTab = Invoke-TestStep -Name 'Exists(#openTabButton)' -ScriptBlock { $driver.IsExistsElementGeneric("document.getElementById('openTabButton')", 'JS', 'openTabButton') }
    if ($existsOpenTab -eq $true) {
        $openTabBtn = Invoke-TestStep -Name 'FindElement(#openTabButton)' -ScriptBlock { $driver.FindElement('#openTabButton') }
        if ($openTabBtn -and $openTabBtn.nodeId) {
            Invoke-TestStep -Name 'ClickElement(openTabButton)' -ScriptBlock { $driver.ClickElement($openTabBtn.nodeId) }
            Start-Sleep -Seconds 2
            $current = $driver.GetTabInfomation()
            $tabs = Invoke-TestStep -Name 'GetAvailableTabs()' -ScriptBlock { $driver.GetAvailableTabs() }
            if ($tabs -and $tabs.targetInfos) {
                $newTab = $tabs.targetInfos | Where-Object { $_.type -eq 'page' -and $_.url -like '*example.com*' -and $_.targetId -ne $current.id } | Select-Object -First 1
                if ($newTab) {
                    Invoke-TestStep -Name "SetActiveTab($($newTab.targetId))" -ScriptBlock { $driver.SetActiveTab($newTab.targetId) }
                    Invoke-TestStep -Name "CloseTab($($newTab.targetId))"    -ScriptBlock { $driver.CloseTab($newTab.targetId) }
                } else {
                    Write-Host '新しい example.com タブが検出できなかったため、タブ切替をスキップします。'
                }
            }
        }
    } else {
        Write-Host 'openTabButton が存在しないため、タブ操作をスキップします。'
    }
    
    # ========================================
    # 21. リソース解放関連メソッド
    # ========================================
    Write-Host "`n--- 21. リソース解放関連メソッド ---" -ForegroundColor Yellow
    
    # 最後にブラウザを閉じる
    Invoke-TestStep -Name 'CloseWindow()' -ScriptBlock { $driver.CloseWindow() }
    
    Write-Host "`n=== WebDriver 全メソッド機能テスト完了 ===" -ForegroundColor Cyan
}
finally {
    if ($driver) {
        Invoke-TestStep -Name 'Dispose()' -ScriptBlock { $driver.Dispose() }
    }
}

Write-Host "テスト完了: EdgeDriver (親WebDriverメソッド) の全機能を実行しました。" -ForegroundColor Green


