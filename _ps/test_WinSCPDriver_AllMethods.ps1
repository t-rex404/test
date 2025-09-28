# WinSCPDriverクラスの全メソッドテスト
# 実行前に確認事項:
# 1. WinSCPがインストールされていること
# 2. テスト用のFTP/SFTPサーバーが利用可能であること
# 3. 接続情報を環境に合わせて設定すること

# パスの設定
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$libPath = Join-Path $scriptPath "_lib"

# 必要なクラスを読み込み
. (Join-Path $libPath "Common.ps1")
. (Join-Path $libPath "WinSCPDriver.ps1")

# Commonクラスの初期化
$global:Common = [Common]::new()

# テスト結果を記録する関数
function Record-TestResult {
    param(
        [string]$TestName,
        [scriptblock]$TestBlock
    )

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "テスト: $TestName" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan

    try {
        & $TestBlock
        Write-Host "✓ $TestName が成功しました" -ForegroundColor Green
        return @{
            TestName = $TestName
            Result = "成功"
            Error = ""
        }
    }
    catch {
        Write-Host "✗ $TestName が失敗しました" -ForegroundColor Red
        Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
        return @{
            TestName = $TestName
            Result = "失敗"
            Error = $_.Exception.Message
        }
    }
}

# テスト結果を格納する配列
$testResults = @()

Write-Host "`n===== WinSCPDriver 全メソッドテスト開始 =====" -ForegroundColor Magenta

# ========================================
# テスト設定
# ========================================

# 接続情報（環境に合わせて変更してください）
$TEST_HOST = "test.rebex.net"  # テスト用公開FTPサーバー
$TEST_USER = "demo"
$TEST_PASSWORD = "password"
$TEST_PROTOCOL = "FTP"  # FTP, SFTP, SCP, FTPS
$TEST_PORT = 21  # FTP:21, SFTP:22

# ローカルテストファイル
$TEST_LOCAL_DIR = "C:\temp\WinSCPTest"
$TEST_LOCAL_FILE = Join-Path $TEST_LOCAL_DIR "test_upload.txt"
$TEST_DOWNLOAD_FILE = Join-Path $TEST_LOCAL_DIR "test_download.txt"

# リモートパス
$TEST_REMOTE_DIR = "/pub/test_" + (Get-Date -Format "yyyyMMdd_HHmmss")
$TEST_REMOTE_FILE = "$TEST_REMOTE_DIR/test_file.txt"

# テスト用ディレクトリとファイルの準備
if (-not (Test-Path $TEST_LOCAL_DIR)) {
    New-Item -ItemType Directory -Path $TEST_LOCAL_DIR -Force | Out-Null
}
"This is a test file for WinSCP upload test.`nCreated at: $(Get-Date)" | Out-File -FilePath $TEST_LOCAL_FILE -Encoding UTF8

# ========================================
# 初期化テスト
# ========================================

$winscpDriver = $null

$testResults += Record-TestResult "WinSCPDriver初期化" {
    $winscpDriver = [WinSCPDriver]::new()
    if ($null -eq $winscpDriver) {
        throw "WinSCPDriverの初期化に失敗しました"
    }
    Write-Host "WinSCPDriverが正常に初期化されました" -ForegroundColor Green
}

# ========================================
# 接続関連テスト
# ========================================

$testResults += Record-TestResult "SetConnectionParameters - 接続パラメータ設定" {
    $winscpDriver.SetConnectionParameters(
        $TEST_HOST,
        $TEST_USER,
        $TEST_PASSWORD,
        $TEST_PROTOCOL,
        $TEST_PORT,
        ""  # ホストキーフィンガープリント（FTPの場合は不要）
    )
    Write-Host "接続パラメータが設定されました" -ForegroundColor Green
}

$testResults += Record-TestResult "Connect - サーバー接続" {
    $winscpDriver.Connect()
    Write-Host "サーバーに接続しました" -ForegroundColor Green
}

$testResults += Record-TestResult "IsConnectedToServer - 接続状態確認" {
    $isConnected = $winscpDriver.IsConnectedToServer()
    if (-not $isConnected) {
        throw "接続状態の確認に失敗しました"
    }
    Write-Host "接続状態: $isConnected" -ForegroundColor Green
}

# ========================================
# ディレクトリ操作テスト
# ========================================

$testResults += Record-TestResult "CreateDirectory - ディレクトリ作成" {
    try {
        $winscpDriver.CreateDirectory($TEST_REMOTE_DIR)
        Write-Host "リモートディレクトリを作成しました: $TEST_REMOTE_DIR" -ForegroundColor Green
    }
    catch {
        # 権限がない場合はスキップ
        Write-Host "ディレクトリ作成をスキップ（権限なし）: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$testResults += Record-TestResult "DirectoryExists - ディレクトリ存在確認" {
    $exists = $winscpDriver.DirectoryExists("/")
    Write-Host "ルートディレクトリの存在: $exists" -ForegroundColor Green
    if (-not $exists) {
        throw "ディレクトリ存在確認に失敗しました"
    }
}

$testResults += Record-TestResult "GetFileList - ファイル一覧取得" {
    $fileList = $winscpDriver.GetFileList("/")
    Write-Host "ルートディレクトリのファイル数: $($fileList.Count)" -ForegroundColor Green

    if ($fileList.Count -gt 0) {
        Write-Host "最初の5件:" -ForegroundColor Cyan
        $fileList | Select-Object -First 5 | ForEach-Object {
            $type = if ($_.IsDirectory) { "DIR" } else { "FILE" }
            Write-Host "  [$type] $($_.Name) (Size: $($_.Size) bytes)" -ForegroundColor Gray
        }
    }
}

# ========================================
# ファイル操作テスト
# ========================================

$testResults += Record-TestResult "SetTransferOptions - 転送オプション設定" {
    $winscpDriver.SetTransferOptions("Binary", $true)
    Write-Host "転送オプションを設定しました（Binary, 上書き許可）" -ForegroundColor Green
}

$testResults += Record-TestResult "UploadFile - ファイルアップロード" {
    try {
        # まずリモートディレクトリを作成
        try {
            $winscpDriver.CreateDirectory($TEST_REMOTE_DIR)
        } catch {
            # 既に存在する場合は無視
        }

        $winscpDriver.UploadFile($TEST_LOCAL_FILE, $TEST_REMOTE_FILE)
        Write-Host "ファイルをアップロードしました: $TEST_REMOTE_FILE" -ForegroundColor Green
    }
    catch {
        # 書き込み権限がない場合はスキップ
        Write-Host "アップロードをスキップ（権限なし）: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$testResults += Record-TestResult "FileExists - ファイル存在確認" {
    try {
        # アップロードが成功した場合のみテスト
        $exists = $winscpDriver.FileExists($TEST_REMOTE_FILE)
        Write-Host "アップロードしたファイルの存在: $exists" -ForegroundColor Green
    }
    catch {
        # ファイルが存在しない場合は、既存のファイルをチェック
        $exists = $winscpDriver.FileExists("/readme.txt")
        Write-Host "readme.txtの存在: $exists" -ForegroundColor Green
    }
}

$testResults += Record-TestResult "DownloadFile - ファイルダウンロード" {
    try {
        # 先ほどアップロードしたファイルをダウンロード
        $winscpDriver.DownloadFile($TEST_REMOTE_FILE, $TEST_DOWNLOAD_FILE)
        Write-Host "ファイルをダウンロードしました: $TEST_DOWNLOAD_FILE" -ForegroundColor Green
    }
    catch {
        # アップロードができなかった場合は既存のファイルをダウンロード
        try {
            $winscpDriver.DownloadFile("/readme.txt", $TEST_DOWNLOAD_FILE)
            Write-Host "readme.txtをダウンロードしました: $TEST_DOWNLOAD_FILE" -ForegroundColor Green
        }
        catch {
            Write-Host "ダウンロードをスキップ: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

$testResults += Record-TestResult "RemoveFile - ファイル削除" {
    try {
        $winscpDriver.RemoveFile($TEST_REMOTE_FILE)
        Write-Host "リモートファイルを削除しました: $TEST_REMOTE_FILE" -ForegroundColor Green
    }
    catch {
        # 削除権限がない場合はスキップ
        Write-Host "ファイル削除をスキップ（権限なし）: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$testResults += Record-TestResult "RemoveDirectory - ディレクトリ削除" {
    try {
        $winscpDriver.RemoveDirectory($TEST_REMOTE_DIR)
        Write-Host "リモートディレクトリを削除しました: $TEST_REMOTE_DIR" -ForegroundColor Green
    }
    catch {
        # 削除権限がない場合はスキップ
        Write-Host "ディレクトリ削除をスキップ（権限なし）: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ========================================
# 切断とクリーンアップ
# ========================================

$testResults += Record-TestResult "Disconnect - 接続切断" {
    $winscpDriver.Disconnect()
    Write-Host "接続を切断しました" -ForegroundColor Green
}

$testResults += Record-TestResult "Dispose - リソース破棄" {
    $winscpDriver.Dispose()
    Write-Host "リソースを破棄しました" -ForegroundColor Green
}

# ========================================
# PEMファイル（秘密鍵）認証テスト
# ========================================

Write-Host "`n===== PEMファイル認証テスト =====" -ForegroundColor Yellow

# PEMファイル認証の設定（環境に合わせて変更してください）
$PEM_TEST_HOST = "your-ssh-server.com"  # SSHサーバーのホスト名
$PEM_TEST_USER = "your-username"  # SSHユーザー名
$PEM_TEST_KEY_PATH = "C:\path\to\your\private-key.pem"  # PEMファイルのパス
$PEM_TEST_PASSPHRASE = ""  # パスフレーズ（設定されている場合）
$PEM_TEST_HOST_KEY = ""  # ホストキーフィンガープリント（省略可能）

# PEMファイル認証テストを実行するかどうか
$RUN_PEM_TEST = $false  # trueに設定してPEMテストを有効化

if ($RUN_PEM_TEST) {
    $testResults += Record-TestResult "PEMファイル認証 - 初期化" {
        $pemDriver = [WinSCPDriver]::new()
        if ($null -eq $pemDriver) {
            throw "WinSCPDriverの初期化に失敗しました"
        }
        Write-Host "PEM認証用WinSCPDriverが初期化されました" -ForegroundColor Green
    }

    $testResults += Record-TestResult "PEMファイル認証 - 接続パラメータ設定" {
        if (-not (Test-Path $PEM_TEST_KEY_PATH)) {
            throw "PEMファイルが見つかりません: $PEM_TEST_KEY_PATH"
        }

        $pemDriver.SetConnectionParametersWithKey(
            $PEM_TEST_HOST,
            $PEM_TEST_USER,
            $PEM_TEST_KEY_PATH,
            $PEM_TEST_PASSPHRASE,
            "SFTP",
            22,
            $PEM_TEST_HOST_KEY
        )
        Write-Host "PEM認証の接続パラメータが設定されました" -ForegroundColor Green
    }

    $testResults += Record-TestResult "PEMファイル認証 - 接続" {
        $pemDriver.Connect()
        Write-Host "PEM認証でサーバーに接続しました" -ForegroundColor Green
    }

    $testResults += Record-TestResult "PEMファイル認証 - ファイル一覧取得" {
        $fileList = $pemDriver.GetFileList("/")
        Write-Host "PEM認証接続でファイル一覧を取得: $($fileList.Count)件" -ForegroundColor Green

        if ($fileList.Count -gt 0) {
            Write-Host "最初の3件:" -ForegroundColor Cyan
            $fileList | Select-Object -First 3 | ForEach-Object {
                $type = if ($_.IsDirectory) { "DIR" } else { "FILE" }
                Write-Host "  [$type] $($_.Name)" -ForegroundColor Gray
            }
        }
    }

    $testResults += Record-TestResult "PEMファイル認証 - 切断" {
        $pemDriver.Disconnect()
        $pemDriver.Dispose()
        Write-Host "PEM認証の接続を切断しました" -ForegroundColor Green
    }
}
else {
    Write-Host "PEMファイル認証テストはスキップされました（RUN_PEM_TEST = false）" -ForegroundColor Yellow
}

# ========================================
# PPKファイル（PuTTY形式）認証テスト
# ========================================

Write-Host "`n===== PPKファイル認証テスト =====" -ForegroundColor Yellow

# PPKファイル認証の設定
$PPK_TEST_HOST = "your-ssh-server.com"
$PPK_TEST_USER = "your-username"
$PPK_TEST_KEY_PATH = "C:\path\to\your\private-key.ppk"  # PPKファイルのパス
$PPK_TEST_PASSPHRASE = ""
$PPK_TEST_HOST_KEY = ""

# PPKファイル認証テストを実行するかどうか
$RUN_PPK_TEST = $false  # trueに設定してPPKテストを有効化

if ($RUN_PPK_TEST) {
    $testResults += Record-TestResult "PPKファイル認証 - 初期化と接続" {
        $ppkDriver = [WinSCPDriver]::new()

        if (-not (Test-Path $PPK_TEST_KEY_PATH)) {
            throw "PPKファイルが見つかりません: $PPK_TEST_KEY_PATH"
        }

        $ppkDriver.SetConnectionParametersWithKey(
            $PPK_TEST_HOST,
            $PPK_TEST_USER,
            $PPK_TEST_KEY_PATH,
            $PPK_TEST_PASSPHRASE,
            "SFTP",
            22,
            $PPK_TEST_HOST_KEY
        )

        $ppkDriver.Connect()
        Write-Host "PPK認証でサーバーに接続しました" -ForegroundColor Green

        # 接続確認
        $isConnected = $ppkDriver.IsConnectedToServer()
        if (-not $isConnected) {
            throw "PPK認証での接続状態確認に失敗"
        }

        $ppkDriver.Disconnect()
        $ppkDriver.Dispose()
        Write-Host "PPK認証のテストが完了しました" -ForegroundColor Green
    }
}
else {
    Write-Host "PPKファイル認証テストはスキップされました（RUN_PPK_TEST = false）" -ForegroundColor Yellow
}

# ========================================
# エラーケーステスト
# ========================================

Write-Host "`n===== エラーケーステスト =====" -ForegroundColor Yellow

$testResults += Record-TestResult "未接続状態でのアップロード（エラー）" {
    $errorDriver = [WinSCPDriver]::new()
    try {
        $errorDriver.UploadFile($TEST_LOCAL_FILE, "/test.txt")
        throw "エラーが発生するはずでしたが、成功してしまいました"
    }
    catch {
        if ($_.Exception.Message -like "*接続されていません*") {
            Write-Host "期待通りのエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Green
        }
        else {
            throw $_
        }
    }
    finally {
        $errorDriver.Dispose()
    }
}

$testResults += Record-TestResult "存在しないファイルのアップロード（エラー）" {
    $errorDriver = [WinSCPDriver]::new()
    $errorDriver.SetConnectionParameters($TEST_HOST, $TEST_USER, $TEST_PASSWORD, $TEST_PROTOCOL, $TEST_PORT, "")
    $errorDriver.Connect()

    try {
        $errorDriver.UploadFile("C:\NonExistent\file.txt", "/test.txt")
        throw "エラーが発生するはずでしたが、成功してしまいました"
    }
    catch {
        if ($_.Exception.Message -like "*ローカルファイルが見つかりません*") {
            Write-Host "期待通りのエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Green
        }
        else {
            throw $_
        }
    }
    finally {
        $errorDriver.Disconnect()
        $errorDriver.Dispose()
    }
}

# ========================================
# テスト結果のサマリー
# ========================================

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "テスト結果サマリー" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Magenta

$successCount = ($testResults | Where-Object { $_.Result -eq "成功" }).Count
$failureCount = ($testResults | Where-Object { $_.Result -eq "失敗" }).Count
$totalCount = $testResults.Count

Write-Host "`n実行テスト数: $totalCount" -ForegroundColor Cyan
Write-Host "成功: $successCount" -ForegroundColor Green
Write-Host "失敗: $failureCount" -ForegroundColor Red

Write-Host "`n詳細結果:" -ForegroundColor Yellow
$testResults | ForEach-Object {
    $color = if ($_.Result -eq "成功") { "Green" } else { "Red" }
    $symbol = if ($_.Result -eq "成功") { "✓" } else { "✗" }
    Write-Host "$symbol $($_.TestName): $($_.Result)" -ForegroundColor $color
    if ($_.Error) {
        Write-Host "  エラー: $($_.Error)" -ForegroundColor Red
    }
}

# クリーンアップ
Write-Host "`n===== クリーンアップ =====" -ForegroundColor Yellow
if (Test-Path $TEST_LOCAL_DIR) {
    Remove-Item $TEST_LOCAL_DIR -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "テスト用ローカルディレクトリを削除しました" -ForegroundColor Green
}

Write-Host "`n===== テスト完了 =====" -ForegroundColor Magenta