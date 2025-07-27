# WebDriverエラー管理モジュール
# WebDriverクラスで使用するエラーコードとエラーメッセージを管理

# エラーコード定数
$WebDriverErrorCodes = @{
    # 初期化関連エラー (1001-1010)
    INIT_ERROR = "1001"
    BROWSER_START_ERROR = "1002"
    TAB_INFO_ERROR = "1003"
    WEBSOCKET_CONNECTION_ERROR = "1004"
    WEBSOCKET_SEND_ERROR = "1005"
    WEBSOCKET_RECEIVE_ERROR = "1006"
    NAVIGATION_ERROR = "1007"
    AD_LOAD_WAIT_ERROR = "1008"
    CLOSE_WINDOW_ERROR = "1009"
    DISPOSE_ERROR = "1010"
    
    # 要素操作関連エラー (1011-1026)
    FIND_ELEMENT_ERROR = "1011"
    FIND_ELEMENT_GENERIC_ERROR = "1012"
    FIND_ELEMENT_XPATH_ERROR = "1013"
    FIND_ELEMENT_CLASSNAME_ERROR = "1014"
    FIND_ELEMENT_ID_ERROR = "1015"
    FIND_ELEMENT_NAME_ERROR = "1016"
    FIND_ELEMENT_TAGNAME_ERROR = "1017"
    GET_ELEMENT_TEXT_ERROR = "1018"
    SET_ELEMENT_TEXT_ERROR = "1019"
    CLICK_ELEMENT_ERROR = "1020"
    GET_ELEMENT_ATTRIBUTE_ERROR = "1021"
    SET_ELEMENT_ATTRIBUTE_ERROR = "1022"
    GET_HREF_ERROR = "1023"
    SCREENSHOT_ERROR = "1024"
    SCREENSHOT_OBJECT_ERROR = "1025"
    SCREENSHOT_OBJECTS_ERROR = "1026"
    
    # タブ・ターゲット関連エラー (1027-1031)
    DISCOVER_TARGETS_ERROR = "1027"
    GET_AVAILABLE_TABS_ERROR = "1028"
    SET_ACTIVE_TAB_ERROR = "1029"
    CLOSE_TAB_ERROR = "1030"
    ENABLE_PAGE_EVENTS_ERROR = "1031"
    
    # 複数要素操作関連エラー (1032-1033)
    FIND_ELEMENTS_ERROR = "1032"
    FIND_ELEMENTS_GENERIC_ERROR = "1033"
    
    # 要素存在確認関連エラー (1034-1039)
    IS_EXISTS_ELEMENT_GENERIC_ERROR = "1034"
    IS_EXISTS_ELEMENT_XPATH_ERROR = "1035"
    IS_EXISTS_ELEMENT_CLASSNAME_ERROR = "1036"
    IS_EXISTS_ELEMENT_ID_ERROR = "1037"
    IS_EXISTS_ELEMENT_NAME_ERROR = "1038"
    IS_EXISTS_ELEMENT_TAGNAME_ERROR = "1039"
    
    # CSS・スタイル関連エラー (1040-1041)
    GET_ELEMENT_CSS_PROPERTY_ERROR = "1040"
    SET_ELEMENT_CSS_PROPERTY_ERROR = "1041"
    
    # フォーム操作関連エラー (1042-1045)
    SELECT_OPTION_BY_INDEX_ERROR = "1042"
    SELECT_OPTION_BY_TEXT_ERROR = "1043"
    DESELECT_ALL_OPTIONS_ERROR = "1044"
    CLEAR_ELEMENT_ERROR = "1045"
    
    # ウィンドウ操作関連エラー (1046-1051)
    RESIZE_WINDOW_ERROR = "1046"
    NORMAL_WINDOW_ERROR = "1047"
    MAXIMIZE_WINDOW_ERROR = "1048"
    MINIMIZE_WINDOW_ERROR = "1049"
    FULLSCREEN_WINDOW_ERROR = "1050"
    MOVE_WINDOW_ERROR = "1051"
    
    # ナビゲーション関連エラー (1052-1054)
    GO_BACK_ERROR = "1052"
    GO_FORWARD_ERROR = "1053"
    REFRESH_ERROR = "1054"
    
    # 情報取得関連エラー (1055-1060)
    GET_URL_ERROR = "1055"
    GET_TITLE_ERROR = "1056"
    GET_SOURCE_CODE_ERROR = "1057"
    GET_WINDOW_HANDLE_ERROR = "1058"
    GET_WINDOW_HANDLES_ERROR = "1059"
    GET_WINDOW_SIZE_ERROR = "1060"
    
    # 複数要素検索関連エラー (1061-1063)
    FIND_ELEMENTS_BY_CLASSNAME_ERROR = "1061"
    FIND_ELEMENTS_BY_NAME_ERROR = "1062"
    FIND_ELEMENTS_BY_TAGNAME_ERROR = "1063"
    
    # 待機機能関連エラー (1064-1067)
    WAIT_FOR_ELEMENT_VISIBLE_ERROR = "1064"
    WAIT_FOR_ELEMENT_CLICKABLE_ERROR = "1065"
    WAIT_FOR_PAGE_LOAD_ERROR = "1066"
    WAIT_FOR_CONDITION_ERROR = "1067"
    
    # キーボード操作関連エラー (1068-1069)
    SEND_KEYS_ERROR = "1068"
    SEND_SPECIAL_KEY_ERROR = "1069"
    
    # マウス操作関連エラー (1070-1072)
    MOUSE_HOVER_ERROR = "1070"
    DOUBLE_CLICK_ERROR = "1071"
    RIGHT_CLICK_ERROR = "1072"
    
    # フォーム操作関連エラー (1073-1074)
    SET_CHECKBOX_ERROR = "1073"
    SELECT_RADIO_BUTTON_ERROR = "1074"
    
    # ファイルアップロード関連エラー (1075)
    UPLOAD_FILE_ERROR = "1075"
    
    # JavaScript実行関連エラー (1076-1077)
    EXECUTE_SCRIPT_ERROR = "1076"
    EXECUTE_SCRIPT_ASYNC_ERROR = "1077"
    
    # クッキー操作関連エラー (1078-1081)
    GET_COOKIE_ERROR = "1078"
    SET_COOKIE_ERROR = "1079"
    DELETE_COOKIE_ERROR = "1080"
    CLEAR_ALL_COOKIES_ERROR = "1081"
    
    # ローカルストレージ操作関連エラー (1082-1085)
    GET_LOCAL_STORAGE_ERROR = "1082"
    SET_LOCAL_STORAGE_ERROR = "1083"
    REMOVE_LOCAL_STORAGE_ERROR = "1084"
    CLEAR_LOCAL_STORAGE_ERROR = "1085"
}

# エラーメッセージ定数
$WebDriverErrorMessages = @{
    # 初期化関連エラーメッセージ
    "1001" = "WebDriver初期化エラー"
    "1002" = "ブラウザ起動エラー"
    "1003" = "タブ情報取得エラー"
    "1004" = "WebSocket接続エラー"
    "1005" = "WebSocketメッセージ送信エラー"
    "1006" = "WebSocketメッセージ受信エラー"
    "1007" = "ページ遷移エラー"
    "1008" = "広告読み込み待機エラー"
    "1009" = "ウィンドウを閉じるエラー"
    "1010" = "Disposeエラー"
    
    # 要素操作関連エラーメッセージ
    "1011" = "要素検索エラー (CSS)"
    "1012" = "要素検索エラー (JavaScript)"
    "1013" = "要素検索エラー (XPath)"
    "1014" = "要素検索エラー (ClassName)"
    "1015" = "要素検索エラー (Id)"
    "1016" = "要素検索エラー (Name)"
    "1017" = "要素検索エラー (TagName)"
    "1018" = "要素テキスト取得エラー"
    "1019" = "要素テキスト設定エラー"
    "1020" = "要素クリックエラー"
    "1021" = "要素属性取得エラー"
    "1022" = "要素属性設定エラー"
    "1023" = "href取得エラー"
    "1024" = "スクリーンショット取得エラー"
    "1025" = "要素スクリーンショット取得エラー"
    "1026" = "複数要素スクリーンショット取得エラー"
    
    # タブ・ターゲット関連エラーメッセージ
    "1027" = "ターゲット発見エラー"
    "1028" = "タブ情報取得エラー"
    "1029" = "タブアクティブ化エラー"
    "1030" = "タブ切断エラー"
    "1031" = "ページイベント有効化エラー"
    
    # 複数要素操作関連エラーメッセージ
    "1032" = "複数要素検索エラー (CSS)"
    "1033" = "複数要素検索エラー (JavaScript)"
    
    # 要素存在確認関連エラーメッセージ
    "1034" = "要素存在確認エラー (JavaScript)"
    "1035" = "XPath要素存在確認エラー"
    "1036" = "ClassName要素存在確認エラー"
    "1037" = "Id要素存在確認エラー"
    "1038" = "Name要素存在確認エラー"
    "1039" = "TagName要素存在確認エラー"
    
    # CSS・スタイル関連エラーメッセージ
    "1040" = "CSSプロパティ取得エラー"
    "1041" = "CSSプロパティ設定エラー"
    
    # フォーム操作関連エラーメッセージ
    "1042" = "オプション選択エラー (インデックス)"
    "1043" = "オプション選択エラー (テキスト)"
    "1044" = "オプション未選択エラー"
    "1045" = "要素クリアエラー"
    
    # ウィンドウ操作関連エラーメッセージ
    "1046" = "ウィンドウリサイズエラー"
    "1047" = "ウィンドウ状態変更エラー (通常)"
    "1048" = "ウィンドウ状態変更エラー (最大化)"
    "1049" = "ウィンドウ状態変更エラー (最小化)"
    "1050" = "ウィンドウ状態変更エラー (フルスクリーン)"
    "1051" = "ウィンドウ移動エラー"
    
    # ナビゲーション関連エラーメッセージ
    "1052" = "ブラウザ履歴移動エラー (戻る)"
    "1053" = "ブラウザ履歴移動エラー (進む)"
    "1054" = "ブラウザ更新エラー"
    
    # 情報取得関連エラーメッセージ
    "1055" = "URL取得エラー"
    "1056" = "タイトル取得エラー"
    "1057" = "ソースコード取得エラー"
    "1058" = "ウィンドウハンドル取得エラー"
    "1059" = "複数ウィンドウハンドル取得エラー"
    "1060" = "ウィンドウサイズ取得エラー"
    
    # 複数要素検索関連エラーメッセージ
    "1061" = "複数要素検索エラー (ClassName)"
    "1062" = "複数要素検索エラー (Name)"
    "1063" = "複数要素検索エラー (TagName)"
    
    # 待機機能関連エラーメッセージ
    "1064" = "要素表示待機エラー"
    "1065" = "要素クリック可能性待機エラー"
    "1066" = "ページロード待機エラー"
    "1067" = "カスタム条件待機エラー"
    
    # キーボード操作関連エラーメッセージ
    "1068" = "キーボード入力エラー"
    "1069" = "特殊キー送信エラー"
    
    # マウス操作関連エラーメッセージ
    "1070" = "マウスホバーエラー"
    "1071" = "ダブルクリックエラー"
    "1072" = "右クリックエラー"
    
    # フォーム操作関連エラーメッセージ
    "1073" = "チェックボックス設定エラー"
    "1074" = "ラジオボタン選択エラー"
    
    # ファイルアップロード関連エラーメッセージ
    "1075" = "ファイルアップロードエラー"
    
    # JavaScript実行関連エラーメッセージ
    "1076" = "JavaScript実行エラー"
    "1077" = "JavaScript非同期実行エラー"
    
    # クッキー操作関連エラーメッセージ
    "1078" = "クッキー取得エラー"
    "1079" = "クッキー設定エラー"
    "1080" = "クッキー削除エラー"
    "1081" = "全クッキー削除エラー"
    
    # ローカルストレージ操作関連エラーメッセージ
    "1082" = "ローカルストレージ取得エラー"
    "1083" = "ローカルストレージ設定エラー"
    "1084" = "ローカルストレージ削除エラー"
    "1085" = "ローカルストレージクリアエラー"
}

# WebDriver専用のLogError関数
function LogWebDriverError([string]$error_code, [string]$message)
{
    try
    {
        $log_file = '.\WebDriver_Error.log'
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        
        # エラーコードに対応するメッセージを取得
        $error_title = if ($WebDriverErrorMessages.ContainsKey($error_code)) {
            $WebDriverErrorMessages[$error_code]
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
- モジュール: WebDriver
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
Export-ModuleMember -Function LogWebDriverError -Variable WebDriverErrorCodes, WebDriverErrorMessages 