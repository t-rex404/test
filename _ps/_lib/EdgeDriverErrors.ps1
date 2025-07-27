# EdgeDriverエラー管理モジュール
# EdgeDriverクラスで使用するエラーコードとエラーメッセージを管理

# エラーコード定数
$EdgeDriverErrorCodes = @{
    # 初期化関連エラー (2001-2005)
    INIT_ERROR = "2001"
    EXECUTABLE_PATH_ERROR = "2002"
    USER_DATA_DIR_ERROR = "2003"
    DEBUG_MODE_ERROR = "2004"
    DISPOSE_ERROR = "2005"
}

# エラーメッセージ定数
$EdgeDriverErrorMessages = @{
    # 初期化関連エラーメッセージ
    "2001" = "EdgeDriver初期化エラー"
    "2002" = "Edge実行ファイルパス取得エラー"
    "2003" = "ユーザーデータディレクトリ作成エラー"
    "2004" = "デバッグモード有効化エラー"
    "2005" = "EdgeDriver Disposeエラー"
}

# EdgeDriver専用のLogError関数
function LogEdgeDriverError([string]$error_code, [string]$message)
{
    try
    {
        $log_file = '.\EdgeDriver_Error.log'
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        
        # エラーコードに対応するメッセージを取得
        $error_title = if ($EdgeDriverErrorMessages.ContainsKey($error_code)) {
            $EdgeDriverErrorMessages[$error_code]
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
- モジュール: EdgeDriver
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
Export-ModuleMember -Function LogEdgeDriverError -Variable EdgeDriverErrorCodes, EdgeDriverErrorMessages 