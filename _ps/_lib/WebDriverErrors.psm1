# WebDriverエラー管理モジュール
# WebDriverクラスで使用するエラーコードとエラーメッセージを管理

# エラーコード定数
$WebDriverErrorCodes = @{
    # ========================================
    # 初期化・接続関連エラー (1001-1010)
    # ========================================
    INIT_ERROR = "1001"                           # WebDriver初期化エラー
    BROWSER_START_ERROR = "1002"                  # ブラウザ起動エラー
    TAB_INFO_ERROR = "1003"                       # タブ情報取得エラー
    WEBSOCKET_CONNECTION_ERROR = "1004"           # WebSocket接続エラー
    WEBSOCKET_SEND_ERROR = "1005"                 # WebSocketメッセージ送信エラー
    WEBSOCKET_RECEIVE_ERROR = "1006"              # WebSocketメッセージ受信エラー
    DISPOSE_ERROR = "1007"                        # Disposeエラー
    
    # ========================================
    # ナビゲーション関連エラー (1011-1020)
    # ========================================
    NAVIGATION_ERROR = "1011"                     # ページ遷移エラー
    GO_BACK_ERROR = "1012"                        # ブラウザ履歴移動エラー (戻る)
    GO_FORWARD_ERROR = "1013"                     # ブラウザ履歴移動エラー (進む)
    REFRESH_ERROR = "1014"                        # ブラウザ更新エラー
    WAIT_FOR_PAGE_LOAD_ERROR = "1015"             # ページロード待機エラー
    AD_LOAD_WAIT_ERROR = "1016"                   # 広告読み込み待機エラー
    
    # ========================================
    # 要素検索関連エラー (1021-1040)
    # ========================================
    # 単一要素検索
    FIND_ELEMENT_ERROR = "1021"                   # 要素検索エラー (CSS)
    FIND_ELEMENT_GENERIC_ERROR = "1022"           # 要素検索エラー (JavaScript)
    FIND_ELEMENT_XPATH_ERROR = "1023"             # 要素検索エラー (XPath)
    FIND_ELEMENT_CLASSNAME_ERROR = "1024"         # 要素検索エラー (ClassName)
    FIND_ELEMENT_ID_ERROR = "1025"                # 要素検索エラー (Id)
    FIND_ELEMENT_NAME_ERROR = "1026"              # 要素検索エラー (Name)
    FIND_ELEMENT_TAGNAME_ERROR = "1027"           # 要素検索エラー (TagName)
    
    # 複数要素検索
    FIND_ELEMENTS_ERROR = "1028"                  # 複数要素検索エラー (CSS)
    FIND_ELEMENTS_GENERIC_ERROR = "1029"          # 複数要素検索エラー (JavaScript)
    FIND_ELEMENTS_BY_CLASSNAME_ERROR = "1030"     # 複数要素検索エラー (ClassName)
    FIND_ELEMENTS_BY_NAME_ERROR = "1031"          # 複数要素検索エラー (Name)
    FIND_ELEMENTS_BY_TAGNAME_ERROR = "1032"       # 複数要素検索エラー (TagName)
    
    # 要素存在確認
    IS_EXISTS_ELEMENT_GENERIC_ERROR = "1033"      # 要素存在確認エラー (JavaScript)
    IS_EXISTS_ELEMENT_XPATH_ERROR = "1034"        # XPath要素存在確認エラー
    IS_EXISTS_ELEMENT_CLASSNAME_ERROR = "1035"    # ClassName要素存在確認エラー
    IS_EXISTS_ELEMENT_ID_ERROR = "1036"           # Id要素存在確認エラー
    IS_EXISTS_ELEMENT_NAME_ERROR = "1037"         # Name要素存在確認エラー
    IS_EXISTS_ELEMENT_TAGNAME_ERROR = "1038"      # TagName要素存在確認エラー
    
    # 要素待機
    WAIT_FOR_ELEMENT_VISIBLE_ERROR = "1039"       # 要素表示待機エラー
    WAIT_FOR_ELEMENT_CLICKABLE_ERROR = "1040"     # 要素クリック可能性待機エラー
    WAIT_FOR_CONDITION_ERROR = "1041"             # カスタム条件待機エラー
    
    # ========================================
    # 要素操作関連エラー (1051-1070)
    # ========================================
    # テキスト操作
    GET_ELEMENT_TEXT_ERROR = "1051"               # 要素テキスト取得エラー
    SET_ELEMENT_TEXT_ERROR = "1052"               # 要素テキスト設定エラー
    CLEAR_ELEMENT_ERROR = "1053"                  # 要素クリアエラー
    
    # マウス操作
    CLICK_ELEMENT_ERROR = "1054"                  # 要素クリックエラー
    DOUBLE_CLICK_ERROR = "1055"                   # ダブルクリックエラー
    RIGHT_CLICK_ERROR = "1056"                    # 右クリックエラー
    MOUSE_HOVER_ERROR = "1057"                    # マウスホバーエラー
    
    # キーボード操作
    SEND_KEYS_ERROR = "1058"                      # キーボード入力エラー
    SEND_SPECIAL_KEY_ERROR = "1059"               # 特殊キー送信エラー
    
    # 属性・CSS操作
    GET_ELEMENT_ATTRIBUTE_ERROR = "1060"          # 要素属性取得エラー
    SET_ELEMENT_ATTRIBUTE_ERROR = "1061"          # 要素属性設定エラー
    GET_ELEMENT_CSS_PROPERTY_ERROR = "1062"       # CSSプロパティ取得エラー
    SET_ELEMENT_CSS_PROPERTY_ERROR = "1063"       # CSSプロパティ設定エラー
    GET_HREF_ERROR = "1064"                       # href取得エラー
    
    # フォーム操作
    SELECT_OPTION_BY_INDEX_ERROR = "1065"         # オプション選択エラー (インデックス)
    SELECT_OPTION_BY_TEXT_ERROR = "1066"          # オプション選択エラー (テキスト)
    DESELECT_ALL_OPTIONS_ERROR = "1067"           # オプション未選択エラー
    SET_CHECKBOX_ERROR = "1068"                   # チェックボックス設定エラー
    SELECT_RADIO_BUTTON_ERROR = "1069"            # ラジオボタン選択エラー
    UPLOAD_FILE_ERROR = "1070"                    # ファイルアップロードエラー
    
    # ========================================
    # ウィンドウ・タブ操作関連エラー (1071-1080)
    # ========================================
    # ウィンドウ操作
    RESIZE_WINDOW_ERROR = "1071"                  # ウィンドウリサイズエラー
    NORMAL_WINDOW_ERROR = "1072"                  # ウィンドウ状態変更エラー (通常)
    MAXIMIZE_WINDOW_ERROR = "1073"                # ウィンドウ状態変更エラー (最大化)
    MINIMIZE_WINDOW_ERROR = "1074"                # ウィンドウ状態変更エラー (最小化)
    FULLSCREEN_WINDOW_ERROR = "1075"              # ウィンドウ状態変更エラー (フルスクリーン)
    MOVE_WINDOW_ERROR = "1076"                    # ウィンドウ移動エラー
    CLOSE_WINDOW_ERROR = "1077"                   # ウィンドウを閉じるエラー
    
    # ウィンドウ情報取得
    GET_WINDOW_HANDLE_ERROR = "1078"              # ウィンドウハンドル取得エラー
    GET_WINDOW_HANDLES_ERROR = "1079"             # 複数ウィンドウハンドル取得エラー
    GET_WINDOW_SIZE_ERROR = "1080"                # ウィンドウサイズ取得エラー
    
    # タブ操作
    DISCOVER_TARGETS_ERROR = "1081"               # ターゲット発見エラー
    GET_AVAILABLE_TABS_ERROR = "1082"             # タブ情報取得エラー
    SET_ACTIVE_TAB_ERROR = "1083"                 # タブアクティブ化エラー
    CLOSE_TAB_ERROR = "1084"                      # タブ切断エラー
    ENABLE_PAGE_EVENTS_ERROR = "1085"             # ページイベント有効化エラー
    
    # ========================================
    # 情報取得関連エラー (1091-1100)
    # ========================================
    GET_URL_ERROR = "1091"                        # URL取得エラー
    GET_TITLE_ERROR = "1092"                      # タイトル取得エラー
    GET_SOURCE_CODE_ERROR = "1093"                # ソースコード取得エラー
    
    # ========================================
    # スクリーンショット関連エラー (1101-1110)
    # ========================================
    SCREENSHOT_ERROR = "1101"                     # スクリーンショット取得エラー
    SCREENSHOT_OBJECT_ERROR = "1102"              # 要素スクリーンショット取得エラー
    SCREENSHOT_OBJECTS_ERROR = "1103"             # 複数要素スクリーンショット取得エラー
    
    # ========================================
    # JavaScript実行関連エラー (1111-1120)
    # ========================================
    EXECUTE_SCRIPT_ERROR = "1111"                 # JavaScript実行エラー
    EXECUTE_SCRIPT_ASYNC_ERROR = "1112"           # JavaScript非同期実行エラー
    
    # ========================================
    # ストレージ操作関連エラー (1121-1130)
    # ========================================
    # クッキー操作
    GET_COOKIE_ERROR = "1121"                     # クッキー取得エラー
    SET_COOKIE_ERROR = "1122"                     # クッキー設定エラー
    DELETE_COOKIE_ERROR = "1123"                  # クッキー削除エラー
    CLEAR_ALL_COOKIES_ERROR = "1124"              # 全クッキー削除エラー
    
    # ローカルストレージ操作
    GET_LOCAL_STORAGE_ERROR = "1125"              # ローカルストレージ取得エラー
    SET_LOCAL_STORAGE_ERROR = "1126"              # ローカルストレージ設定エラー
    REMOVE_LOCAL_STORAGE_ERROR = "1127"           # ローカルストレージ削除エラー
    CLEAR_LOCAL_STORAGE_ERROR = "1128"            # ローカルストレージクリアエラー
}

# エラーメッセージ定数
$WebDriverErrorMessages = @{
    # ========================================
    # 初期化・接続関連エラーメッセージ
    # ========================================
    "1001" = "WebDriver初期化エラー"
    "1002" = "ブラウザ起動エラー"
    "1003" = "タブ情報取得エラー"
    "1004" = "WebSocket接続エラー"
    "1005" = "WebSocketメッセージ送信エラー"
    "1006" = "WebSocketメッセージ受信エラー"
    "1007" = "Disposeエラー"
    
    # ========================================
    # ナビゲーション関連エラーメッセージ
    # ========================================
    "1011" = "ページ遷移エラー"
    "1012" = "ブラウザ履歴移動エラー (戻る)"
    "1013" = "ブラウザ履歴移動エラー (進む)"
    "1014" = "ブラウザ更新エラー"
    "1015" = "ページロード待機エラー"
    "1016" = "広告読み込み待機エラー"
    
    # ========================================
    # 要素検索関連エラーメッセージ
    # ========================================
    # 単一要素検索
    "1021" = "要素検索エラー (CSS)"
    "1022" = "要素検索エラー (JavaScript)"
    "1023" = "要素検索エラー (XPath)"
    "1024" = "要素検索エラー (ClassName)"
    "1025" = "要素検索エラー (Id)"
    "1026" = "要素検索エラー (Name)"
    "1027" = "要素検索エラー (TagName)"
    
    # 複数要素検索
    "1028" = "複数要素検索エラー (CSS)"
    "1029" = "複数要素検索エラー (JavaScript)"
    "1030" = "複数要素検索エラー (ClassName)"
    "1031" = "複数要素検索エラー (Name)"
    "1032" = "複数要素検索エラー (TagName)"
    
    # 要素存在確認
    "1033" = "要素存在確認エラー (JavaScript)"
    "1034" = "XPath要素存在確認エラー"
    "1035" = "ClassName要素存在確認エラー"
    "1036" = "Id要素存在確認エラー"
    "1037" = "Name要素存在確認エラー"
    "1038" = "TagName要素存在確認エラー"
    
    # 要素待機
    "1039" = "要素表示待機エラー"
    "1040" = "要素クリック可能性待機エラー"
    "1041" = "カスタム条件待機エラー"
    
    # ========================================
    # 要素操作関連エラーメッセージ
    # ========================================
    # テキスト操作
    "1051" = "要素テキスト取得エラー"
    "1052" = "要素テキスト設定エラー"
    "1053" = "要素クリアエラー"
    
    # マウス操作
    "1054" = "要素クリックエラー"
    "1055" = "ダブルクリックエラー"
    "1056" = "右クリックエラー"
    "1057" = "マウスホバーエラー"
    
    # キーボード操作
    "1058" = "キーボード入力エラー"
    "1059" = "特殊キー送信エラー"
    
    # 属性・CSS操作
    "1060" = "要素属性取得エラー"
    "1061" = "要素属性設定エラー"
    "1062" = "CSSプロパティ取得エラー"
    "1063" = "CSSプロパティ設定エラー"
    "1064" = "href取得エラー"
    
    # フォーム操作
    "1065" = "オプション選択エラー (インデックス)"
    "1066" = "オプション選択エラー (テキスト)"
    "1067" = "オプション未選択エラー"
    "1068" = "チェックボックス設定エラー"
    "1069" = "ラジオボタン選択エラー"
    "1070" = "ファイルアップロードエラー"
    
    # ========================================
    # ウィンドウ・タブ操作関連エラーメッセージ
    # ========================================
    # ウィンドウ操作
    "1071" = "ウィンドウリサイズエラー"
    "1072" = "ウィンドウ状態変更エラー (通常)"
    "1073" = "ウィンドウ状態変更エラー (最大化)"
    "1074" = "ウィンドウ状態変更エラー (最小化)"
    "1075" = "ウィンドウ状態変更エラー (フルスクリーン)"
    "1076" = "ウィンドウ移動エラー"
    "1077" = "ウィンドウを閉じるエラー"
    
    # ウィンドウ情報取得
    "1078" = "ウィンドウハンドル取得エラー"
    "1079" = "複数ウィンドウハンドル取得エラー"
    "1080" = "ウィンドウサイズ取得エラー"
    
    # タブ操作
    "1081" = "ターゲット発見エラー"
    "1082" = "タブ情報取得エラー"
    "1083" = "タブアクティブ化エラー"
    "1084" = "タブ切断エラー"
    "1085" = "ページイベント有効化エラー"
    
    # ========================================
    # 情報取得関連エラーメッセージ
    # ========================================
    "1091" = "URL取得エラー"
    "1092" = "タイトル取得エラー"
    "1093" = "ソースコード取得エラー"
    
    # ========================================
    # スクリーンショット関連エラーメッセージ
    # ========================================
    "1101" = "スクリーンショット取得エラー"
    "1102" = "要素スクリーンショット取得エラー"
    "1103" = "複数要素スクリーンショット取得エラー"
    
    # ========================================
    # JavaScript実行関連エラーメッセージ
    # ========================================
    "1111" = "JavaScript実行エラー"
    "1112" = "JavaScript非同期実行エラー"
    
    # ========================================
    # ストレージ操作関連エラーメッセージ
    # ========================================
    # クッキー操作
    "1121" = "クッキー取得エラー"
    "1122" = "クッキー設定エラー"
    "1123" = "クッキー削除エラー"
    "1124" = "全クッキー削除エラー"
    
    # ローカルストレージ操作
    "1125" = "ローカルストレージ取得エラー"
    "1126" = "ローカルストレージ設定エラー"
    "1127" = "ローカルストレージ削除エラー"
    "1128" = "ローカルストレージクリアエラー"
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