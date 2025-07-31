# 共通ライブラリ
# 各ドライバークラスで使用する共通機能を提供

# エラー管理モジュールをインポート
#. "$PSScriptRoot\WebDriverErrors.ps1"
#. "$PSScriptRoot\EdgeDriverErrors.ps1"
#. "$PSScriptRoot\ChromeDriverErrors.ps1"
#. "$PSScriptRoot\WordDriverErrors.ps1"
#import-module "$PSScriptRoot\WebDriverErrors.psm1"
#import-module "$PSScriptRoot\EdgeDriverErrors.psm1"
#import-module "$PSScriptRoot\ChromeDriverErrors.psm1"
#import-module "$PSScriptRoot\WordDriverErrors.psm1"

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
    
    # 共通のエラーハンドリング関数（拡張版）
    [void] HandleError([string]$errorCode, [string]$message, [string]$module = "Common", [string]$logFile = "")
    {
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        
        # エラータイトルを取得
        $error_title = $this.GetErrorTitle($errorCode, $module)
        
        # ログファイル名を決定
        if ([string]::IsNullOrEmpty($logFile))
        {
            $logFile = ".\Common_Error.log"
        }
        
        $errorMessage = "[$timestamp], ERROR_CODE:$errorCode, ERROR_TITLE:$error_title, MODULE:$module, ERROR_MESSAGE:$message"
        
        # ログファイルに書き込み
        $errorMessage | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
        
        # コンソールにエラーを表示
        Write-Error $errorMessage
        
        # 詳細なエラー情報をログに追加（デバッグ用）
        $debugInfo = @"
詳細情報:
- エラーコード: $errorCode
- エラータイトル: $error_title
- エラーメッセージ: $message
- モジュール: $module
- タイムスタンプ: $timestamp
- 実行ユーザー: $env:USERNAME
- 実行パス: $PWD
"@
        
        $debugInfo | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    
    # エラータイトルを取得する関数
    [string] GetErrorTitle([string]$errorCode, [string]$module)
    {
        try
        {
            switch ($module.ToUpper())
            {
                "WEBDRIVER" {
                    switch ($errorCode)
                    {
                        "1001" { return "WebDriver初期化エラー" }
                        "1002" { return "ブラウザ起動エラー" }
                        "1003" { return "タブ情報取得エラー" }
                        "1004" { return "WebSocket接続エラー" }
                        "1005" { return "WebSocket メッセージ送信エラー" }
                        "1006" { return "WebSocket メッセージ受信エラー" }
                        "1007" { return "Disposeエラー" }
                        "1011" { return "ページ遷移エラー" }
                        "1012" { return "ブラウザ履歴移動エラー (戻る)" }
                        "1013" { return "ブラウザ履歴移動エラー (進む)" }
                        "1014" { return "ブラウザ更新エラー" }
                        "1015" { return "ページロード待機エラー" }
                        "1016" { return "広告読み込み待機エラー" }
                        "1021" { return "要素検索エラー (CSS)" }
                        "1022" { return "要素検索エラー" }
                        "1028" { return "複数要素検索エラー (CSS)" }
                        "1029" { return "複数要素検索エラー" }
                        "1030" { return "複数要素検索エラー (ClassName)" }
                        "1031" { return "複数要素検索エラー (Name)" }
                        "1032" { return "複数要素検索エラー (TagName)" }
                        "1033" { return "要素存在確認エラー" }
                        "1034" { return "XPath要素存在確認エラー" }
                        "1035" { return "ClassName要素存在確認エラー" }
                        "1036" { return "Id要素存在確認エラー" }
                        "1037" { return "Name要素存在確認エラー" }
                        "1038" { return "TagName要素存在確認エラー" }
                        "1039" { return "要素表示待機エラー" }
                        "1040" { return "要素クリック可能性待機エラー" }
                        "1041" { return "カスタム条件待機エラー" }
                        "1051" { return "要素テキスト取得エラー" }
                        "1052" { return "要素テキスト設定エラー" }
                        "1053" { return "要素クリアエラー" }
                        "1054" { return "要素クリックエラー" }
                        "1055" { return "ダブルクリックエラー" }
                        "1056" { return "右クリックエラー" }
                        "1057" { return "マウスホバーエラー" }
                        "1058" { return "キーボード入力エラー" }
                        "1059" { return "特殊キー送信エラー" }
                        "1060" { return "要素属性取得エラー" }
                        "1061" { return "要素属性設定エラー" }
                        "1062" { return "CSSプロパティ取得エラー" }
                        "1063" { return "CSSプロパティ設定エラー" }
                        "1064" { return "href取得エラー" }
                        "1065" { return "オプション選択エラー (インデックス)" }
                        "1066" { return "オプション選択エラー (テキスト)" }
                        "1067" { return "オプション未選択エラー" }
                        "1068" { return "チェックボックス設定エラー" }
                        "1069" { return "ラジオボタン選択エラー" }
                        "1070" { return "ファイルアップロードエラー" }
                        "1071" { return "ウィンドウリサイズエラー" }
                        "1072" { return "ウィンドウ状態変更エラー (通常)" }
                        "1073" { return "ウィンドウ状態変更エラー (最大化)" }
                        "1074" { return "ウィンドウ状態変更エラー (最小化)" }
                        "1075" { return "ウィンドウ状態変更エラー (フルスクリーン)" }
                        "1076" { return "ウィンドウ移動エラー" }
                        "1077" { return "ウィンドウを閉じるエラー" }
                        "1078" { return "ウィンドウハンドル取得エラー" }
                        "1079" { return "複数ウィンドウハンドル取得エラー" }
                        "1080" { return "ウィンドウサイズ取得エラー" }
                        "1081" { return "ターゲット発見エラー" }
                        "1082" { return "タブ情報取得エラー" }
                        "1083" { return "タブアクティブ化エラー" }
                        "1084" { return "タブ切断エラー" }
                        "1085" { return "ページイベント有効化エラー" }
                        "1091" { return "URL取得エラー" }
                        "1092" { return "タイトル取得エラー" }
                        "1093" { return "ソースコード取得エラー" }
                        "1101" { return "スクリーンショット取得エラー" }
                        "1102" { return "要素スクリーンショット取得エラー" }
                        "1103" { return "複数要素スクリーンショット取得エラー" }
                        "1111" { return "JavaScript実行エラー" }
                        "1112" { return "JavaScript非同期実行エラー" }
                        "1121" { return "クッキー取得エラー" }
                        "1122" { return "クッキー設定エラー" }
                        "1123" { return "クッキー削除エラー" }
                        "1124" { return "全クッキー削除エラー" }
                        "1125" { return "ローカルストレージ取得エラー" }
                        "1126" { return "ローカルストレージ設定エラー" }
                        "1127" { return "ローカルストレージ削除エラー" }
                        "1128" { return "ローカルストレージクリアエラー" }
                        default { return "不明なエラー" }
                    }
                }
                "EDGEDRIVER" {
                    switch ($errorCode)
                    {
                        "2001" { return "EdgeDriver初期化エラー" }
                        "2002" { return "Edge実行ファイルパス取得エラー" }
                        "2003" { return "ユーザーデータディレクトリ作成エラー" }
                        "2004" { return "デバッグモード有効化エラー" }
                        "2005" { return "EdgeDriver Disposeエラー" }
                        default { return "不明なエラー" }
                    }
                }
                "CHROMEDRIVER" {
                    switch ($errorCode)
                    {
                        "3001" { return "ChromeDriver初期化エラー" }
                        "3002" { return "Chrome実行ファイルパス取得エラー" }
                        "3003" { return "ユーザーデータディレクトリ作成エラー" }
                        "3004" { return "デバッグモード有効化エラー" }
                        "3005" { return "ChromeDriver Disposeエラー" }
                        default { return "不明なエラー" }
                    }
                }
                "WORDDRIVER" {
                    switch ($errorCode)
                    {
                        "4001" { return "WordDriver初期化エラー" }
                        "4002" { return "一時ディレクトリ作成エラー" }
                        "4003" { return "Wordアプリケーション初期化エラー" }
                        "4004" { return "新規ドキュメント作成エラー" }
                        "4005" { return "テキスト追加エラー" }
                        "4006" { return "見出し追加エラー" }
                        "4007" { return "段落追加エラー" }
                        "4008" { return "表追加エラー" }
                        "4009" { return "画像追加エラー" }
                        "4010" { return "ページ区切り追加エラー" }
                        "4011" { return "目次追加エラー" }
                        "4012" { return "ドキュメント保存エラー" }
                        "4013" { return "目次更新エラー" }
                        "4014" { return "ドキュメントオープンエラー" }
                        "4015" { return "フォント設定エラー" }
                        "4016" { return "WordDriver Disposeエラー" }
                        default { return "不明なエラー" }
                    }
                }
                default {
                    return "不明なエラー"
                }
            }
        }
        catch
        {
            return "不明なエラー"
        }
        
        return "不明なエラー"
    }
}

# 共通インスタンスを作成（グローバルスコープ）
$global:Common = [Common]::new()

Write-Host "Commonライブラリが正常にインポートされました。" -ForegroundColor Green