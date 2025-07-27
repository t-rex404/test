# 共通ライブラリ
# 各ドライバークラスで使用する共通機能を提供

# エラー管理モジュールをインポート
. "$PSScriptRoot\WebDriverErrors.ps1"
. "$PSScriptRoot\EdgeDriverErrors.ps1"
. "$PSScriptRoot\ChromeDriverErrors.ps1"
. "$PSScriptRoot\WordDriverErrors.ps1"

class Common
{
    Common()
    {
        # 共通クラスの初期化
    }
    
    # 共通のログ出力関数
    [void] WriteLog([string]$message, [string]$level = "INFO")
    {
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        $logMessage = "[$timestamp] [$level] $message"
        
        # コンソールに出力
        switch ($level.ToUpper())
        {
            "ERROR" { Write-Error $logMessage }
            "WARNING" { Write-Warning $logMessage }
            "INFO" { Write-Host $logMessage }
            "DEBUG" { Write-Host $logMessage -ForegroundColor Gray }
            default { Write-Host $logMessage }
        }
        
        # ログファイルに出力
        $logFile = ".\Common_$($level.ToLower()).log"
        $logMessage | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    
    # 共通のエラーハンドリング関数
    [void] HandleError([string]$errorCode, [string]$message, [string]$module = "Common")
    {
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        $errorMessage = "[$timestamp], ERROR_CODE:$errorCode, MODULE:$module, ERROR_MESSAGE:$message"
        
        # ログファイルに書き込み
        $logFile = ".\Common_Error.log"
        $errorMessage | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
        
        # コンソールにエラーを表示
        Write-Error $errorMessage
        
        # 詳細なエラー情報をログに追加（デバッグ用）
        $debugInfo = @"
詳細情報:
- エラーコード: $errorCode
- エラーメッセージ: $message
- モジュール: $module
- タイムスタンプ: $timestamp
- PowerShellバージョン: $($PSVersionTable.PSVersion)
- OS: $($PSVersionTable.OS)
- 実行ユーザー: $env:USERNAME
- 実行パス: $PWD
"@
        
        $debugInfo | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# 共通インスタンスを作成
$Common = [Common]::new()

Write-Host "Commonライブラリが正常にインポートされました。" -ForegroundColor Green