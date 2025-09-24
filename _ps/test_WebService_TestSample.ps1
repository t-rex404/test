# ========================================
# Webサービステスト用サンプルスクリプト
# EdgeDriver, WordDriver, OracleDriverを統合した実践的なテスト例
# ========================================

# エラーハンドリング設定
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ========================================
# 設定セクション（実環境に合わせて変更してください）
# ========================================
$TestConfig = @{
    # テスト対象Webサービス設定
    WebService = @{
        BaseUrl = "https://your-webapp.example.com"  # テスト対象のWebサービスURL
        LoginUrl = "/login"
        DashboardUrl = "/dashboard"
        DataEntryUrl = "/data-entry"
        ReportUrl = "/reports"
        TestAccount = @{
            Username = "test_user@example.com"
            Password = "SecurePassword123!"
        }
    }

    # Oracle Database設定
    Database = @{
        Server = "db-server.example.com"
        Port = "1521"
        ServiceName = "TESTDB"
        Username = "test_automation"
        Password = "DBPassword123!"
        Schema = "TEST_SCHEMA"
    }

    # テストレポート設定
    Report = @{
        OutputDirectory = ".\_test_results"
        ScreenshotDirectory = ".\_test_results\screenshots"
        ReportPrefix = "WebServiceTest"
    }

    # リトライ設定
    Retry = @{
        MaxAttempts = 3
        WaitSeconds = 2
    }

    # タイムアウト設定
    Timeout = @{
        PageLoad = 30
        ElementWait = 10
        DatabaseQuery = 60
    }
}

# ========================================
# 共通関数定義
# ========================================

function Write-StepResult {
    param(
        [string]$Name,
        [bool]$Succeeded,
        [string]$Message,
        [string]$Details = ""
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $status = if ($Succeeded) { '[OK]' } else { '[NG]' }
    $statusColor = if ($Succeeded) { 'Green' } else { 'Red' }

    Write-Host "$timestamp $status " -NoNewline -ForegroundColor $statusColor
    Write-Host "$Name - $Message" -ForegroundColor White

    if (-not [string]::IsNullOrEmpty($Details)) {
        Write-Host "         詳細: $Details" -ForegroundColor Gray
    }
}

function Invoke-TestStep {
    param(
        [string]$Name,
        [scriptblock]$ScriptBlock,
        [bool]$ContinueOnError = $false,
        [int]$RetryCount = 1
    )

    $attempt = 1
    while ($attempt -le $RetryCount) {
        try {
            $result = & $ScriptBlock
            Write-StepResult -Name $Name -Succeeded $true -Message "成功"
            return @{
                Success = $true
                Result = $result
                ErrorMessage = $null
            }
        } catch {
            if ($attempt -eq $RetryCount) {
                Write-StepResult -Name $Name -Succeeded $false -Message "失敗" -Details $_.Exception.Message
                if (-not $ContinueOnError) {
                    throw
                }
                return @{
                    Success = $false
                    Result = $null
                    ErrorMessage = $_.Exception.Message
                }
            }
            Write-Host "  リトライ $attempt/$RetryCount..." -ForegroundColor Yellow
            Start-Sleep -Seconds $TestConfig.Retry.WaitSeconds
            $attempt++
        }
    }
}

function Take-Screenshot {
    param(
        [object]$EdgeDriver,
        [string]$FileName
    )

    try {
        $screenshotDir = $TestConfig.Report.ScreenshotDirectory
        if (-not (Test-Path $screenshotDir)) {
            New-Item -ItemType Directory -Path $screenshotDir -Force | Out-Null
        }

        $screenshotPath = Join-Path $screenshotDir "$FileName.png"
        $EdgeDriver.TakeScreenshot($screenshotPath)
        return $screenshotPath
    } catch {
        Write-Host "スクリーンショット取得失敗: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Initialize-TestEnvironment {
    Write-Host "`n=== テスト環境初期化 ===" -ForegroundColor Cyan

    # 出力ディレクトリ作成
    foreach ($dir in @($TestConfig.Report.OutputDirectory, $TestConfig.Report.ScreenshotDirectory)) {
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Host "  ディレクトリ作成: $dir" -ForegroundColor Gray
        }
    }

    # テスト開始時刻を記録
    $global:TestStartTime = Get-Date
    Write-Host "  テスト開始時刻: $($global:TestStartTime.ToString('yyyy/MM/dd HH:mm:ss'))" -ForegroundColor Gray
}

# ========================================
# メインテスト処理
# ========================================

# スクリプトパス設定
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LibDir = Join-Path $ScriptDir '_lib'

# 依存ライブラリ読み込み
Write-Host "依存ライブラリを読み込んでいます..." -ForegroundColor Yellow
. (Join-Path $LibDir 'Common.ps1')
. (Join-Path $LibDir 'WebDriver.ps1')
. (Join-Path $LibDir 'EdgeDriver.ps1')
. (Join-Path $LibDir 'WordDriver.ps1')
. (Join-Path $LibDir 'OracleDriver.ps1')

# アセンブリ読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -Path (Join-Path $LibDir 'Oracle.ManagedDataAccess.dll')

# Commonオブジェクト初期化
$global:Common = [Common]::new()

# テスト環境初期化
Initialize-TestEnvironment

# ドライバー変数初期化
$edge = $null
$word = $null
$oracle = $null
$testResults = @()

try {
    Write-Host "`n=== Webサービステスト開始 ===" -ForegroundColor Cyan

    # ========================================
    # 1. ドライバー初期化
    # ========================================
    Write-Host "`n--- 1. ドライバー初期化 ---" -ForegroundColor Yellow

    # EdgeDriver初期化
    $edgeInit = Invoke-TestStep -Name "EdgeDriver初期化" -ScriptBlock {
        $script:edge = [EdgeDriver]::new()
        Start-Sleep -Seconds 2
    } -RetryCount 2

    # WordDriver初期化
    $wordInit = Invoke-TestStep -Name "WordDriver初期化" -ScriptBlock {
        $script:word = [WordDriver]::new()
        $script:word.StartWord($true)
        $script:wordDoc = $script:word.CreateNewDocument()
    }

    # OracleDriver初期化（実環境では実際の接続を行う）
    $oracleInit = Invoke-TestStep -Name "OracleDriver初期化" -ScriptBlock {
        $script:oracle = [OracleDriver]::new()
        # 実環境では以下のような接続を行う
        # $script:oracle.SetConnectionParameters(
        #     $TestConfig.Database.Server,
        #     $TestConfig.Database.Port,
        #     $TestConfig.Database.ServiceName,
        #     $TestConfig.Database.Username,
        #     $TestConfig.Database.Password,
        #     $TestConfig.Database.Schema
        # )
        # $script:oracle.Connect()
        Write-Host "    (Oracleデータベース接続はシミュレーションモード)" -ForegroundColor Gray
    } -ContinueOnError $true

    # ========================================
    # 2. Webサービステストシナリオ
    # ========================================
    Write-Host "`n--- 2. Webサービステストシナリオ ---" -ForegroundColor Yellow

    # テストケース1: ホームページアクセステスト
    Write-Host "`n[テストケース1: ホームページアクセス]" -ForegroundColor Cyan

    $tc1Result = Invoke-TestStep -Name "ホームページにアクセス" -ScriptBlock {
        $edge.NavigateToUrl($TestConfig.WebService.BaseUrl)
        Start-Sleep -Seconds 3
        $edge.WaitForPageLoad($TestConfig.Timeout.PageLoad)
    } -RetryCount $TestConfig.Retry.MaxAttempts

    $testResults += @{
        TestCase = "ホームページアクセス"
        Result = $tc1Result.Success
        Timestamp = Get-Date
        Details = "ページロード完了"
    }

    # ページタイトル取得
    $pageTitle = Invoke-TestStep -Name "ページタイトル取得" -ScriptBlock {
        $edge.GetTitle()
    }

    if ($pageTitle.Success) {
        Write-Host "  取得タイトル: $($pageTitle.Result)" -ForegroundColor Gray
    }

    # スクリーンショット取得
    $screenshot1 = Take-Screenshot -EdgeDriver $edge -FileName "01_HomePage"

    # テストケース2: ログイン機能テスト
    Write-Host "`n[テストケース2: ログイン機能テスト]" -ForegroundColor Cyan

    # ログインページへ遷移
    $tc2Navigate = Invoke-TestStep -Name "ログインページへ遷移" -ScriptBlock {
        $loginUrl = $TestConfig.WebService.BaseUrl + $TestConfig.WebService.LoginUrl
        $edge.NavigateToUrl($loginUrl)
        Start-Sleep -Seconds 2
        $edge.WaitForPageLoad($TestConfig.Timeout.PageLoad)
    }

    # ログインフォーム要素検索と入力（サンプル）
    $tc2Login = Invoke-TestStep -Name "ログインフォーム入力" -ScriptBlock {
        # 実際の要素セレクタは対象サイトに合わせて変更
        # Username入力
        if ($edge.IsExistsElementById("username") -or $edge.IsExistsElementByName("email", 0)) {
            $usernameField = $edge.FindElement("#username, input[name='email']")
            if ($usernameField) {
                $edge.SetElementText($usernameField.nodeId, $TestConfig.WebService.TestAccount.Username)
            }
        }

        # Password入力
        if ($edge.IsExistsElementById("password") -or $edge.IsExistsElementByName("password", 0)) {
            $passwordField = $edge.FindElement("#password, input[name='password']")
            if ($passwordField) {
                $edge.SetElementText($passwordField.nodeId, $TestConfig.WebService.TestAccount.Password)
            }
        }

        # ログインボタンクリック
        if ($edge.IsExistsElementById("loginButton") -or $edge.IsExistsElementByXPath("//button[contains(text(), 'Login')]")) {
            $loginButton = $edge.FindElement("#loginButton, button[type='submit']")
            if ($loginButton) {
                $edge.ClickElement($loginButton.nodeId)
                Start-Sleep -Seconds 3
            }
        }

        Write-Host "    (ログインフォーム操作はサンプル実装)" -ForegroundColor Gray
    } -ContinueOnError $true

    $testResults += @{
        TestCase = "ログイン機能"
        Result = $tc2Login.Success
        Timestamp = Get-Date
        Details = if ($tc2Login.Success) { "ログイン成功" } else { "ログイン失敗または要素未検出" }
    }

    # スクリーンショット取得
    $screenshot2 = Take-Screenshot -EdgeDriver $edge -FileName "02_LoginPage"

    # テストケース3: フォーム要素操作テスト
    Write-Host "`n[テストケース3: フォーム要素操作テスト]" -ForegroundColor Cyan

    # サンプルHTMLページを使用したフォーム操作デモ
    $tc3Form = Invoke-TestStep -Name "フォーム要素操作テスト" -ScriptBlock {
        # サンプルHTMLがある場合はそれを使用
        $samplePath = Join-Path (Split-Path $ScriptDir) 'sample.html'
        if (Test-Path $samplePath) {
            $sampleUri = [System.Uri]$samplePath
            $edge.NavigateToUrl($sampleUri.AbsoluteUri)
            Start-Sleep -Seconds 2

            # テキストボックス操作
            if ($edge.IsExistsElementById("textbox")) {
                $textbox = $edge.FindElementById("textbox")
                $edge.SetElementText($textbox.nodeId, "テストデータ入力")
            }

            # ドロップダウン操作
            if ($edge.IsExistsElementById("dropdown")) {
                $dropdown = $edge.FindElementById("dropdown")
                $edge.SelectOptionByIndex($dropdown.nodeId, 1)
            }

            # チェックボックス操作
            if ($edge.IsExistsElementById("checkbox1")) {
                $checkbox = $edge.FindElementById("checkbox1")
                $edge.ClickElement($checkbox.nodeId)
            }

            # ラジオボタン操作
            if ($edge.IsExistsElementById("radio2")) {
                $radio = $edge.FindElementById("radio2")
                $edge.ClickElement($radio.nodeId)
            }

            Write-Host "    フォーム要素操作完了" -ForegroundColor Gray
        } else {
            Write-Host "    サンプルHTMLファイルが見つかりません" -ForegroundColor Gray
        }
    } -ContinueOnError $true

    $testResults += @{
        TestCase = "フォーム要素操作"
        Result = $tc3Form.Success
        Timestamp = Get-Date
        Details = "各種フォーム要素の操作テスト"
    }

    # スクリーンショット取得
    $screenshot3 = Take-Screenshot -EdgeDriver $edge -FileName "03_FormTest"

    # ========================================
    # 3. テストデータのデータベース記録（シミュレーション）
    # ========================================
    Write-Host "`n--- 3. テストデータのデータベース記録 ---" -ForegroundColor Yellow

    # テスト実行IDの生成
    $testRunId = [System.Guid]::NewGuid().ToString()
    $testEndTime = Get-Date
    $testDuration = ($testEndTime - $global:TestStartTime).TotalSeconds

    Write-Host "テスト実行ID: $testRunId" -ForegroundColor Gray
    Write-Host "テスト実行時間: $testDuration 秒" -ForegroundColor Gray

    # データベースへの記録（実環境では実際のSQL実行）
    foreach ($result in $testResults) {
        $insertSQL = @"
INSERT INTO TEST_RESULTS (
    TEST_RUN_ID,
    TEST_CASE_NAME,
    RESULT,
    EXECUTION_TIME,
    DETAILS,
    CREATED_DATE
) VALUES (
    '$testRunId',
    '$($result.TestCase)',
    '$(if ($result.Result) { 'PASS' } else { 'FAIL' })',
    $testDuration,
    '$($result.Details)',
    SYSDATE
)
"@
        Write-Host "`nSQL実行（シミュレーション）:" -ForegroundColor Gray
        Write-Host $insertSQL -ForegroundColor DarkGray
    }

    # ========================================
    # 4. Wordレポート生成
    # ========================================
    Write-Host "`n--- 4. テストレポート生成 ---" -ForegroundColor Yellow

    if ($word) {
        # レポートタイトル
        $word.AddText("Webサービステストレポート", $true, $true, 18)
        $word.AddText("`n", $false, $false, 12)

        # 実行情報
        $word.AddText("テスト実行情報", $true, $false, 14)
        $word.AddText("`n", $false, $false, 12)
        $word.AddText("実行ID: $testRunId", $false, $false, 10)
        $word.AddText("`n", $false, $false, 10)
        $word.AddText("開始時刻: $($global:TestStartTime.ToString('yyyy/MM/dd HH:mm:ss'))", $false, $false, 10)
        $word.AddText("`n", $false, $false, 10)
        $word.AddText("終了時刻: $($testEndTime.ToString('yyyy/MM/dd HH:mm:ss'))", $false, $false, 10)
        $word.AddText("`n", $false, $false, 10)
        $word.AddText("実行時間: $testDuration 秒", $false, $false, 10)
        $word.AddText("`n`n", $false, $false, 12)

        # テスト対象情報
        $word.AddText("テスト対象", $true, $false, 14)
        $word.AddText("`n", $false, $false, 12)
        $word.AddText("URL: $($TestConfig.WebService.BaseUrl)", $false, $false, 10)
        $word.AddText("`n", $false, $false, 10)
        $word.AddText("ページタイトル: $(if ($pageTitle.Success) { $pageTitle.Result } else { 'N/A' })", $false, $false, 10)
        $word.AddText("`n`n", $false, $false, 12)

        # テスト結果サマリー
        $word.AddText("テスト結果サマリー", $true, $false, 14)
        $word.AddText("`n", $false, $false, 12)

        $passCount = ($testResults | Where-Object { $_.Result }).Count
        $failCount = ($testResults | Where-Object { -not $_.Result }).Count
        $totalCount = $testResults.Count

        $word.AddText("総テストケース数: $totalCount", $false, $false, 10)
        $word.AddText("`n", $false, $false, 10)
        $word.AddText("成功: $passCount", $false, $false, 10)
        $word.AddText("`n", $false, $false, 10)
        $word.AddText("失敗: $failCount", $false, $false, 10)
        $word.AddText("`n", $false, $false, 10)
        $word.AddText("成功率: $([Math]::Round($passCount / $totalCount * 100, 2))%", $false, $false, 10)
        $word.AddText("`n`n", $false, $false, 12)

        # 詳細結果テーブル
        $word.AddText("詳細テスト結果", $true, $false, 14)
        $word.AddText("`n", $false, $false, 12)

        # テーブルデータ作成
        $tableData = New-Object 'object[,]' ($testResults.Count + 1), 4
        $tableData[0,0] = "テストケース"
        $tableData[0,1] = "結果"
        $tableData[0,2] = "実行時刻"
        $tableData[0,3] = "詳細"

        for ($i = 0; $i -lt $testResults.Count; $i++) {
            $tableData[$i+1,0] = $testResults[$i].TestCase
            $tableData[$i+1,1] = if ($testResults[$i].Result) { "成功" } else { "失敗" }
            $tableData[$i+1,2] = $testResults[$i].Timestamp.ToString("HH:mm:ss")
            $tableData[$i+1,3] = $testResults[$i].Details
        }

        $word.AddTable($tableData, "テスト結果一覧")

        # スクリーンショット追加（存在する場合）
        $word.AddPageBreak()
        $word.AddText("スクリーンショット", $true, $false, 14)
        $word.AddText("`n", $false, $false, 12)

        foreach ($screenshot in @($screenshot1, $screenshot2, $screenshot3)) {
            if ($screenshot -and (Test-Path $screenshot)) {
                $word.AddText("`n", $false, $false, 12)
                $word.AddImage($screenshot, 400, 0)
                $word.AddText("`n", $false, $false, 12)
            }
        }

        # レポート保存
        $reportFileName = "$($TestConfig.Report.ReportPrefix)_$(Get-Date -Format 'yyyyMMdd_HHmmss').docx"
        $reportPath = Join-Path $TestConfig.Report.OutputDirectory $reportFileName
        $word.SaveAs($reportPath)

        Write-Host "`nテストレポート保存完了: $reportPath" -ForegroundColor Green
    }

    # ========================================
    # 5. テスト結果サマリー表示
    # ========================================
    Write-Host "`n=== テスト結果サマリー ===" -ForegroundColor Cyan
    Write-Host "総テストケース数: $totalCount" -ForegroundColor White
    Write-Host "成功: $passCount" -ForegroundColor Green
    Write-Host "失敗: $failCount" -ForegroundColor $(if ($failCount -gt 0) { 'Red' } else { 'Gray' })
    Write-Host "成功率: $([Math]::Round($passCount / $totalCount * 100, 2))%" -ForegroundColor White
    Write-Host "実行時間: $testDuration 秒" -ForegroundColor White

    # 個別結果表示
    Write-Host "`n詳細結果:" -ForegroundColor White
    foreach ($result in $testResults) {
        $statusColor = if ($result.Result) { 'Green' } else { 'Red' }
        $status = if ($result.Result) { '[成功]' } else { '[失敗]' }
        Write-Host "  $status $($result.TestCase): $($result.Details)" -ForegroundColor $statusColor
    }

} catch {
    Write-Host "`n!!! エラーが発生しました !!!" -ForegroundColor Red
    Write-Host "エラー詳細: $($_.Exception.Message)" -ForegroundColor Red

    # エラーログ記録
    if ($global:Common) {
        $global:Common.HandleError(
            "TEST_ERROR",
            $_.Exception.Message,
            "WebServiceTest",
            ".\WebServiceTest_Error.log"
        )
    }

} finally {
    # ========================================
    # 6. クリーンアップ
    # ========================================
    Write-Host "`n--- クリーンアップ ---" -ForegroundColor Yellow

    # EdgeDriver終了
    if ($edge) {
        try {
            $edge.CloseBrowser()
            Write-Host "EdgeDriver終了" -ForegroundColor Gray
        } catch {
            Write-Host "EdgeDriver終了エラー: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # WordDriver終了
    if ($word) {
        try {
            $word.CloseWord()
            Write-Host "WordDriver終了" -ForegroundColor Gray
        } catch {
            Write-Host "WordDriver終了エラー: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # OracleDriver切断
    if ($oracle -and $oracle.IsConnected()) {
        try {
            $oracle.Disconnect()
            Write-Host "OracleDriver切断" -ForegroundColor Gray
        } catch {
            Write-Host "OracleDriver切断エラー: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Host "`n=== Webサービステスト完了 ===" -ForegroundColor Cyan
}

# スクリプト終了
Write-Host "`nテストスクリプトの実行が終了しました。" -ForegroundColor Green
Write-Host "レポートは $($TestConfig.Report.OutputDirectory) フォルダに保存されています。" -ForegroundColor White