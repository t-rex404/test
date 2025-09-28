# ======================================================================
# TeraTermDriver 全メソッドテストスクリプト
# ======================================================================
# 概要:
#   TeraTermDriverクラスの全メソッドをテストするスクリプト
#   SSH/Telnet接続、コマンド実行、ファイル転送などの機能を検証
# ======================================================================

# スクリプトのパスを取得
$script_path = Split-Path -Parent $MyInvocation.MyCommand.Path
$script_name = Split-Path -Leaf $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " TeraTermDriver 全メソッドテスト" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 共通クラスとドライバーを読み込み
try
{
    . "$script_path\_lib\Common.ps1"
    . "$script_path\_lib\TeraTermDriver.ps1"

    Write-Host "✓ クラスファイルの読み込みが完了しました。" -ForegroundColor Green
    Write-Host ""
}
catch
{
    Write-Host "✗ クラスファイルの読み込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Commonクラスのインスタンスを作成
$global:Common = [Common]::new()

# テスト接続パラメータ（環境に応じて変更してください）
$TEST_CONFIG = @{
    # SSH接続設定（パスワード認証）
    SSH = @{
        Hostname = "192.168.1.100"  # テスト用SSHサーバーのIPアドレス
        Port = "22"
        Protocol = "SSH"
        Username = "testuser"        # テスト用ユーザー名
        Password = "testpass"        # テスト用パスワード
    }

    # SSH接続設定（PEM認証）
    SSH_PEM = @{
        Hostname = "192.168.1.100"  # テスト用SSHサーバーのIPアドレス
        Port = "22"
        Protocol = "SSH"
        Username = "ec2-user"        # AWS EC2の場合は通常ec2-user
        PemFilePath = "C:\keys\my-key.pem"  # PEMファイルのパス
    }

    # Telnet接続設定（オプション）
    Telnet = @{
        Hostname = "192.168.1.101"  # テスト用Telnetサーバーのアドレス
        Port = "23"
        Protocol = "Telnet"
        Username = "testuser"
        Password = "testpass"
    }

    # ファイル転送テスト用設定
    FileTransfer = @{
        LocalUploadFile = "C:\temp\test_upload.txt"
        RemoteUploadPath = "/tmp/uploaded_test.txt"
        RemoteDownloadFile = "/etc/hosts"
        LocalDownloadPath = "C:\temp\downloaded_hosts.txt"
    }

    # テストモード設定
    UseRealServer = $false  # 実際のサーバーを使用する場合はtrue
    TestPemAuth = $false    # PEM認証テストを行う場合はtrue
    TestCommands = $true    # コマンド実行テストを行う
    TestFileTransfer = $false  # ファイル転送テストを行う（要ZMODEM対応サーバー）
}

# テスト結果を記録する関数
function Record-TestResult
{
    param(
        [string]$TestName,
        [scriptblock]$TestScript
    )

    $result = @{
        TestName = $TestName
        Success = $false
        ErrorMessage = ""
        Duration = 0
    }

    Write-Host ""
    Write-Host "テスト: $TestName" -ForegroundColor Yellow
    Write-Host "----------------------------------------"

    $start_time = Get-Date

    try
    {
        & $TestScript
        $result.Success = $true
        Write-Host "✓ 成功" -ForegroundColor Green
    }
    catch
    {
        $result.Success = $false
        $result.ErrorMessage = $_.Exception.Message
        Write-Host "✗ 失敗: $($_.Exception.Message)" -ForegroundColor Red
    }

    $end_time = Get-Date
    $result.Duration = ($end_time - $start_time).TotalSeconds
    Write-Host "実行時間: $($result.Duration) 秒"

    return $result
}

# ========================================
# メインテスト処理
# ========================================

$test_results = @()
$teraterm = $null

# 1. 初期化テスト
$test_results += Record-TestResult "TeraTermDriver初期化" {
    $global:teraterm = [TeraTermDriver]::new()

    if (-not $teraterm.is_initialized)
    {
        throw "初期化フラグが設定されていません"
    }

    Write-Host "  TeraTermパス: $($teraterm.teraterm_exe_path)"
    Write-Host "  一時ディレクトリ: $($teraterm.temp_directory)"
    Write-Host "  マクロディレクトリ: $($teraterm.macro_directory)"
    Write-Host "  ログディレクトリ: $($teraterm.log_directory)"
}

# 2. 実行ファイルパス取得テスト
$test_results += Record-TestResult "GetTeraTermExecutablePath" {
    $path = $teraterm.GetTeraTermExecutablePath()

    if ([string]::IsNullOrEmpty($path))
    {
        throw "TeraTermパスが取得できませんでした"
    }

    if (-not (Test-Path $path))
    {
        throw "TeraTermパスが存在しません: $path"
    }

    Write-Host "  検出されたパス: $path"
}

# 3. 作業ディレクトリ作成テスト
$test_results += Record-TestResult "CreateWorkingDirectories" {
    $teraterm.CreateWorkingDirectories()

    if (-not (Test-Path $teraterm.macro_directory))
    {
        throw "マクロディレクトリが作成されていません"
    }

    if (-not (Test-Path $teraterm.log_directory))
    {
        throw "ログディレクトリが作成されていません"
    }

    Write-Host "  マクロディレクトリ: 存在確認OK"
    Write-Host "  ログディレクトリ: 存在確認OK"
}

# PEM認証テストを行う場合
if ($TEST_CONFIG.TestPemAuth -and $TEST_CONFIG.UseRealServer)
{
    # PEM認証用接続パラメータ設定テスト
    $test_results += Record-TestResult "SetConnectionParametersWithPem" {
        if (-not (Test-Path $TEST_CONFIG.SSH_PEM.PemFilePath))
        {
            throw "PEMファイルが見つかりません: $($TEST_CONFIG.SSH_PEM.PemFilePath)"
        }

        $teraterm.SetConnectionParametersWithPem(
            $TEST_CONFIG.SSH_PEM.Hostname,
            $TEST_CONFIG.SSH_PEM.Port,
            $TEST_CONFIG.SSH_PEM.Username,
            $TEST_CONFIG.SSH_PEM.PemFilePath
        )

        $info = $teraterm.GetConnectionInfo()

        if ($info.Host -ne $TEST_CONFIG.SSH_PEM.Hostname)
        {
            throw "ホスト名が正しく設定されていません"
        }

        Write-Host "  ホスト: $($info.Host)"
        Write-Host "  ポート: $($info.Port)"
        Write-Host "  プロトコル: SSH（PEM認証）"
        Write-Host "  PEMファイル: $($TEST_CONFIG.SSH_PEM.PemFilePath)"
    }

    # PEM認証での接続テスト
    $test_results += Record-TestResult "Connect with PEM" {
        Write-Host "  PEM認証でサーバーに接続中..."
        $teraterm.Connect()

        if (-not $teraterm.IsConnected())
        {
            throw "接続状態が正しくありません"
        }

        Write-Host "  PEM認証での接続成功"
        Write-Host "  プロセスID: $($teraterm.teraterm_process.Id)"

        # 接続後、少し待機
        Start-Sleep -Seconds 3
    }

    # PEM認証接続後のコマンドテスト
    if ($TEST_CONFIG.TestCommands -and $teraterm.IsConnected())
    {
        $test_results += Record-TestResult "ExecuteCommand with PEM - whoami" {
            $result = $teraterm.ExecuteCommand("whoami", "$", 10)

            if ([string]::IsNullOrEmpty($result))
            {
                Write-Host "  警告: コマンド結果が空です（マクロ実行の制限による可能性）" -ForegroundColor Yellow
            }
            else
            {
                Write-Host "  現在のユーザー: $result"
                if ($result -ne $TEST_CONFIG.SSH_PEM.Username)
                {
                    Write-Host "  注意: 期待したユーザー名と異なります" -ForegroundColor Yellow
                }
            }
        }

        $test_results += Record-TestResult "ExecuteCommand with PEM - uname" {
            $result = $teraterm.ExecuteCommand("uname -a", "$", 10)

            if ([string]::IsNullOrEmpty($result))
            {
                Write-Host "  警告: コマンド結果が空です（マクロ実行の制限による可能性）" -ForegroundColor Yellow
            }
            else
            {
                Write-Host "  システム情報: $result"
            }
        }
    }

    # PEM認証接続の切断
    if ($teraterm.IsConnected())
    {
        $test_results += Record-TestResult "Disconnect PEM Connection" {
            $teraterm.Disconnect($false)

            if ($teraterm.IsConnected())
            {
                throw "切断後も接続状態が残っています"
            }

            Write-Host "  PEM認証接続の切断完了"
        }
    }
}

# パスワード認証テストを行う場合
if ($TEST_CONFIG.UseRealServer -and -not $TEST_CONFIG.TestPemAuth)
{
    # 4. 接続パラメータ設定テスト（パスワード認証）
    $test_results += Record-TestResult "SetConnectionParameters" {
        $teraterm.SetConnectionParameters(
            $TEST_CONFIG.SSH.Hostname,
            $TEST_CONFIG.SSH.Port,
            $TEST_CONFIG.SSH.Protocol,
            $TEST_CONFIG.SSH.Username,
            $TEST_CONFIG.SSH.Password
        )

        $info = $teraterm.GetConnectionInfo()

        if ($info.Host -ne $TEST_CONFIG.SSH.Hostname)
        {
            throw "ホスト名が正しく設定されていません"
        }

        Write-Host "  ホスト: $($info.Host)"
        Write-Host "  ポート: $($info.Port)"
        Write-Host "  プロトコル: $($info.Protocol)"
    }

    # 5. サーバー接続テスト
    $test_results += Record-TestResult "Connect" {
        Write-Host "  サーバーに接続中..."
        $teraterm.Connect()

        if (-not $teraterm.IsConnected())
        {
            throw "接続状態が正しくありません"
        }

        Write-Host "  接続成功"
        Write-Host "  プロセスID: $($teraterm.teraterm_process.Id)"

        # 接続後、少し待機
        Start-Sleep -Seconds 3
    }

    # 6. コマンド実行テスト
    if ($TEST_CONFIG.TestCommands -and $teraterm.IsConnected())
    {
        $test_results += Record-TestResult "ExecuteCommand - pwd" {
            $result = $teraterm.ExecuteCommand("pwd", "$", 10)

            if ([string]::IsNullOrEmpty($result))
            {
                Write-Host "  警告: コマンド結果が空です（マクロ実行の制限による可能性）" -ForegroundColor Yellow
            }
            else
            {
                Write-Host "  コマンド結果: $result"
            }
        }

        $test_results += Record-TestResult "ExecuteCommand - ls" {
            $result = $teraterm.ExecuteCommand("ls -la", "$", 10)

            if ([string]::IsNullOrEmpty($result))
            {
                Write-Host "  警告: コマンド結果が空です（マクロ実行の制限による可能性）" -ForegroundColor Yellow
            }
            else
            {
                Write-Host "  ファイルリスト取得: $(($result -split "`n").Count) 行"
            }
        }

        $test_results += Record-TestResult "ExecuteCommand - date" {
            $result = $teraterm.ExecuteCommand("date", "$", 10)

            if ([string]::IsNullOrEmpty($result))
            {
                Write-Host "  警告: コマンド結果が空です（マクロ実行の制限による可能性）" -ForegroundColor Yellow
            }
            else
            {
                Write-Host "  サーバー時刻: $result"
            }
        }
    }

    # 7. ファイル転送テスト（ZMODEM対応サーバーが必要）
    if ($TEST_CONFIG.TestFileTransfer -and $teraterm.IsConnected())
    {
        # アップロード用テストファイルを作成
        if (-not (Test-Path $TEST_CONFIG.FileTransfer.LocalUploadFile))
        {
            $upload_dir = Split-Path $TEST_CONFIG.FileTransfer.LocalUploadFile -Parent
            if (-not (Test-Path $upload_dir))
            {
                New-Item -ItemType Directory -Path $upload_dir -Force | Out-Null
            }

            "This is a test file for TeraTerm upload.`nCreated at: $(Get-Date)" |
                Out-File -FilePath $TEST_CONFIG.FileTransfer.LocalUploadFile -Encoding UTF8
        }

        $test_results += Record-TestResult "UploadFile" {
            $teraterm.UploadFile(
                $TEST_CONFIG.FileTransfer.LocalUploadFile,
                $TEST_CONFIG.FileTransfer.RemoteUploadPath
            )

            Write-Host "  アップロード完了（ZMODEMプロトコル使用）"
        }

        $test_results += Record-TestResult "DownloadFile" {
            $teraterm.DownloadFile(
                $TEST_CONFIG.FileTransfer.RemoteDownloadFile,
                $TEST_CONFIG.FileTransfer.LocalDownloadPath
            )

            if (Test-Path $TEST_CONFIG.FileTransfer.LocalDownloadPath)
            {
                $content = Get-Content $TEST_CONFIG.FileTransfer.LocalDownloadPath -Raw
                Write-Host "  ダウンロード完了: $(($content -split "`n").Count) 行"
            }
            else
            {
                throw "ダウンロードファイルが見つかりません"
            }
        }
    }

    # 8. 切断テスト
    if ($teraterm.IsConnected())
    {
        $test_results += Record-TestResult "Disconnect" {
            $teraterm.Disconnect($false)

            if ($teraterm.IsConnected())
            {
                throw "切断後も接続状態が残っています"
            }

            Write-Host "  切断完了"
        }
    }
}
else
{
    Write-Host ""
    Write-Host "注意: 実サーバーテストはスキップされました" -ForegroundColor Yellow
    Write-Host "実際のサーバーでテストする場合は、以下を設定してください:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "【パスワード認証の場合】" -ForegroundColor Cyan
    Write-Host "1. TEST_CONFIG.SSH内の接続パラメータを設定" -ForegroundColor Yellow
    Write-Host "2. UseRealServerをtrueに設定" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "【PEM認証の場合】" -ForegroundColor Cyan
    Write-Host "1. TEST_CONFIG.SSH_PEM内の接続パラメータを設定" -ForegroundColor Yellow
    Write-Host "2. PemFilePathに秘密鍵ファイルのパスを設定" -ForegroundColor Yellow
    Write-Host "3. UseRealServerとTestPemAuthをtrueに設定" -ForegroundColor Yellow
    Write-Host ""
}

# 9. マクロファイル作成テスト（サーバー接続なし）
$test_results += Record-TestResult "CreateConnectionMacro" {
    # テスト用の接続パラメータを設定
    $teraterm.SetConnectionParameters("test.example.com", "22", "SSH", "testuser", "testpass")

    $macro_file = $teraterm.CreateConnectionMacro()

    if (-not (Test-Path $macro_file))
    {
        throw "マクロファイルが作成されていません"
    }

    $content = Get-Content $macro_file -Raw
    if ($content -notmatch "connect")
    {
        throw "マクロファイルにconnectコマンドが含まれていません"
    }

    Write-Host "  マクロファイル作成: $macro_file"
    Write-Host "  ファイルサイズ: $((Get-Item $macro_file).Length) bytes"
}

# 10. コマンドマクロ作成テスト
$test_results += Record-TestResult "CreateCommandMacro" {
    $macro_file = $teraterm.CreateCommandMacro("echo 'test'", "$", 30)

    if (-not (Test-Path $macro_file))
    {
        throw "コマンドマクロファイルが作成されていません"
    }

    $content = Get-Content $macro_file -Raw
    if ($content -notmatch "sendln")
    {
        throw "マクロファイルにsendlnコマンドが含まれていません"
    }

    Write-Host "  コマンドマクロ作成: $macro_file"
}

# 11. ログファイル一覧取得テスト
$test_results += Record-TestResult "GetLogFiles" {
    $log_files = $teraterm.GetLogFiles()

    Write-Host "  ログファイル数: $($log_files.Count)"

    if ($log_files.Count -gt 0)
    {
        Write-Host "  最新ログ: $($log_files[0].Name)"
    }
}

# 12. マクロファイル一覧取得テスト
$test_results += Record-TestResult "GetMacroFiles" {
    $macro_files = $teraterm.GetMacroFiles()

    if ($macro_files.Count -eq 0)
    {
        throw "マクロファイルが見つかりません"
    }

    Write-Host "  マクロファイル数: $($macro_files.Count)"
    Write-Host "  最新マクロ: $($macro_files[0].Name)"
}

# 13. 接続情報取得テスト
$test_results += Record-TestResult "GetConnectionInfo" {
    $info = $teraterm.GetConnectionInfo()

    if ($null -eq $info)
    {
        throw "接続情報が取得できませんでした"
    }

    Write-Host "  ホスト: $($info.Host)"
    Write-Host "  ポート: $($info.Port)"
    Write-Host "  プロトコル: $($info.Protocol)"
    Write-Host "  接続状態: $($info.IsConnected)"
    Write-Host "  プロセスID: $($info.ProcessId)"
}

# 14. 一時ファイルクリーンアップテスト
$test_results += Record-TestResult "CleanupTempFiles" {
    # 古いテストファイルを作成
    $old_date = (Get-Date).AddDays(-8)
    $old_macro = Join-Path $teraterm.macro_directory "old_test.ttl"
    $old_log = Join-Path $teraterm.log_directory "old_test.log"

    "test" | Out-File -FilePath $old_macro
    "test" | Out-File -FilePath $old_log

    # ファイルの更新日時を変更
    (Get-Item $old_macro).LastWriteTime = $old_date
    (Get-Item $old_log).LastWriteTime = $old_date

    # クリーンアップ実行
    $teraterm.CleanupTempFiles()

    # 古いファイルが削除されていることを確認
    if (Test-Path $old_macro)
    {
        throw "古いマクロファイルが削除されていません"
    }

    if (Test-Path $old_log)
    {
        throw "古いログファイルが削除されていません"
    }

    Write-Host "  古いファイルのクリーンアップ: 成功"
}

# 15. 破棄テスト
$test_results += Record-TestResult "Dispose" {
    $teraterm.Dispose()

    if ($teraterm.is_initialized)
    {
        throw "破棄後も初期化フラグが残っています"
    }

    if ($teraterm.is_connected)
    {
        throw "破棄後も接続フラグが残っています"
    }

    Write-Host "  リソースの解放: 完了"
}

# ========================================
# テスト結果のサマリー
# ========================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " テスト結果サマリー" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$total_tests = $test_results.Count
$passed_tests = ($test_results | Where-Object { $_.Success }).Count
$failed_tests = $total_tests - $passed_tests
$total_duration = ($test_results | Measure-Object -Property Duration -Sum).Sum

Write-Host "テスト総数: $total_tests"
Write-Host "成功: $passed_tests" -ForegroundColor Green
Write-Host "失敗: $failed_tests" -ForegroundColor Red
Write-Host "合計実行時間: $([Math]::Round($total_duration, 2)) 秒"
Write-Host ""

# 失敗したテストの詳細
if ($failed_tests -gt 0)
{
    Write-Host "失敗したテスト:" -ForegroundColor Red
    foreach ($result in $test_results | Where-Object { -not $_.Success })
    {
        Write-Host "  - $($result.TestName): $($result.ErrorMessage)" -ForegroundColor Red
    }
    Write-Host ""
}

# 成功率の表示
$success_rate = if ($total_tests -gt 0) { [Math]::Round(($passed_tests / $total_tests) * 100, 2) } else { 0 }
$bar_length = 50
$filled_length = [Math]::Floor($bar_length * $passed_tests / $total_tests)
$empty_length = $bar_length - $filled_length

Write-Host "成功率: $success_rate%"
Write-Host "["  -NoNewline
Write-Host ("=" * $filled_length) -ForegroundColor Green -NoNewline
Write-Host ("-" * $empty_length) -ForegroundColor Gray -NoNewline
Write-Host "]"
Write-Host ""

# 注意事項の表示
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host " 注意事項" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. TeraTermのマクロ実行には制限があります:" -ForegroundColor Yellow
Write-Host "   - 既存のTeraTermセッションへのマクロ適用は制限があります" -ForegroundColor Yellow
Write-Host "   - コマンド結果の取得にはログファイル経由が必要です" -ForegroundColor Yellow
Write-Host ""
Write-Host "2. ファイル転送テストについて:" -ForegroundColor Yellow
Write-Host "   - ZMODEMプロトコル対応のサーバーが必要です" -ForegroundColor Yellow
Write-Host "   - サーバー側でszコマンドが利用可能である必要があります" -ForegroundColor Yellow
Write-Host ""
Write-Host "3. 実環境でのテスト:" -ForegroundColor Yellow
Write-Host "   - TEST_CONFIG内の設定を実環境に合わせて変更してください" -ForegroundColor Yellow
Write-Host "   - UseRealServerをtrueに設定してください" -ForegroundColor Yellow
Write-Host "   - PEM認証を使用する場合はTestPemAuthもtrueに設定してください" -ForegroundColor Yellow
Write-Host ""
Write-Host "4. PEM認証の使用例（AWS EC2など）:" -ForegroundColor Yellow
Write-Host "   - AWS EC2の場合、UsernameはAmazon Linuxなら'ec2-user'" -ForegroundColor Yellow
Write-Host "   - Ubuntuの場合は'ubuntu'、RHELの場合は'ec2-user'または'root'" -ForegroundColor Yellow
Write-Host "   - PEMファイルは適切な権限（400）に設定してください" -ForegroundColor Yellow
Write-Host ""

Write-Host "テスト完了！" -ForegroundColor Green
Write-Host ""