# WordDriverエラー管理モジュール
# WordDriverクラスで使用するエラーコードとエラーメッセージを管理

# エラーコード定数
$WordDriverErrorCodes = @{
    # 初期化関連エラー (4001-4004)
    INIT_ERROR = "4001"
    TEMP_DIR_ERROR = "4002"
    WORD_APP_ERROR = "4003"
    NEW_DOCUMENT_ERROR = "4004"
    
    # テキスト操作関連エラー (4005-4007)
    ADD_TEXT_ERROR = "4005"
    ADD_HEADING_ERROR = "4006"
    ADD_PARAGRAPH_ERROR = "4007"
    
    # コンテンツ操作関連エラー (4008-4010)
    ADD_TABLE_ERROR = "4008"
    ADD_IMAGE_ERROR = "4009"
    ADD_PAGE_BREAK_ERROR = "4010"
    
    # 目次関連エラー (4011-4013)
    ADD_TOC_ERROR = "4011"
    UPDATE_TOC_ERROR = "4013"
    
    # ファイル操作関連エラー (4012, 4014)
    SAVE_ERROR = "4012"
    OPEN_DOCUMENT_ERROR = "4014"
    
    # フォーマット関連エラー (4015)
    SET_FONT_ERROR = "4015"
    
    # リソース管理関連エラー (4016)
    DISPOSE_ERROR = "4016"
}

# エラーメッセージ定数
$WordDriverErrorMessages = @{
    # 初期化関連エラーメッセージ
    "4001" = "WordDriver初期化エラー"
    "4002" = "一時ディレクトリ作成エラー"
    "4003" = "Wordアプリケーション初期化エラー"
    "4004" = "新規ドキュメント作成エラー"
    
    # テキスト操作関連エラーメッセージ
    "4005" = "テキスト追加エラー"
    "4006" = "見出し追加エラー"
    "4007" = "段落追加エラー"
    
    # コンテンツ操作関連エラーメッセージ
    "4008" = "表追加エラー"
    "4009" = "画像追加エラー"
    "4010" = "ページ区切り追加エラー"
    
    # 目次関連エラーメッセージ
    "4011" = "目次追加エラー"
    "4013" = "目次更新エラー"
    
    # ファイル操作関連エラーメッセージ
    "4012" = "ドキュメント保存エラー"
    "4014" = "ドキュメントオープンエラー"
    
    # フォーマット関連エラーメッセージ
    "4015" = "フォント設定エラー"
    
    # リソース管理関連エラーメッセージ
    "4016" = "WordDriver Disposeエラー"
}

# WordDriver専用のLogError関数
function LogWordDriverError([string]$error_code, [string]$message)
{
    try
    {
        $log_file = '.\WordDriver_Error.log'
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        
        # エラーコードに対応するメッセージを取得
        $error_title = if ($WordDriverErrorMessages.ContainsKey($error_code)) {
            $WordDriverErrorMessages[$error_code]
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
- モジュール: WordDriver
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
Export-ModuleMember -Function LogWordDriverError -Variable WordDriverErrorCodes, WordDriverErrorMessages 