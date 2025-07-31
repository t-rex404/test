# ChromeDriverエラー管理モジュール
# ChromeDriverクラスで使用するエラーコードとエラーメッセージを管理

# エラーコード定数
$ChromeDriverErrorCodes = @{
    # 初期化関連エラー (3001-3005)
    INIT_ERROR = "3001"
    EXECUTABLE_PATH_ERROR = "3002"
    USER_DATA_DIR_ERROR = "3003"
    DEBUG_MODE_ERROR = "3004"
    DISPOSE_ERROR = "3005"
}

# エラーメッセージ定数
$ChromeDriverErrorMessages = @{
    # 初期化関連エラーメッセージ
    "3001" = "ChromeDriver初期化エラー"
    "3002" = "Chrome実行ファイルパス取得エラー"
    "3003" = "ユーザーデータディレクトリ作成エラー"
    "3004" = "デバッグモード有効化エラー"
    "3005" = "ChromeDriver Disposeエラー"
}

# ChromeDriver専用のLogError関数
function LogChromeDriverError([string]$error_code, [string]$message)
{
    try
    {
        $log_file = '.\ChromeDriver_Error.log'
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        
        # エラーコードに対応するメッセージを取得
        $error_title = if ($ChromeDriverErrorMessages.ContainsKey($error_code)) {
            $ChromeDriverErrorMessages[$error_code]
        } else {
            "不明なエラー"
        }
        
        $error_message = "[$timestamp], ERROR_CODE:$error_code, ERROR_TITLE:$error_title, ERROR_MESSAGE:$message"

        # ログファイルに書き込み
        $error_message | Out-File -Append -FilePath $log_file -Encoding UTF8 -ErrorAction SilentlyContinue

        # コンソールにエラーを表示
        Write-Error $error_message

        # 詳細なエラー情報をログに追加（デバッグ用）
        $debug_info = @"
詳細情報:
- エラーコード: $error_code
- エラータイトル: $error_title
- エラーメッセージ: $message
- タイムスタンプ: $timestamp
- PowerShellバージョン: $($PSVersionTable.PSVersion)
- OS: $($PSVersionTable.OS)
- 実行ユーザー: $env:USERNAME
- 実行パス: $PWD
- モジュール: ChromeDriver
"@

        $debug_info | Out-File -Append -FilePath $log_file -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch
    {
        # ログ書き込みに失敗した場合でも、コンソールには表示
        Write-Error "ログファイルの書き込みに失敗しました: $($_.Exception.Message)"
        Write-Error "元のエラー: [$error_code] $message"
    }
}

# エクスポートする関数と変数
Export-ModuleMember -Function LogChromeDriverError -Variable ChromeDriverErrorCodes, ChromeDriverErrorMessages 