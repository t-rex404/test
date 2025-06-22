class EdgeDriver : WebDriver
{
    EdgeDriver()
    {
        # ブラウザの実行ファイルのパスを取得
        $browser_exe_reg_key   = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\'
        $browser_user_data_dir = 'C:\temp\UserDataDirectoryForEdge\'
        try
        {
            # ブラウザの実行ファイルのパスを取得
            $browser_exe_path = Get-ItemPropertyValue -Path $browser_exe_reg_key -Name '(default)'
            if (-not $browser_exe_path)
            {
                throw 'ブラウザ実行ファイルが見つかりませんでした。使用するブラウザのインストール状況を確認してください。'
            }
        }
        catch  
        {
            throw '使用するブラウザのパスが見つかりませんでした。エラーメッセージ：' + $_
        }

        # ブラウザの実行
        $this.StartBrowser($browser_exe_path,$browser_user_data_dir)

        # タブ情報を取得
        $tab_infomation = $this.GetTabInfomation()

        # WebSocket接続
        $this.GetWebSocketInfomation($tab_infomation.webSocketDebuggerUrl)

        # タブをアクティブにする
        $this.SetActiveTab($tab_infomation.id)

        # デバッグモードを有効化
        $this.SendWebSocketMessage('Emulation.setEmulatedMedia', @{ media = 'screen' })
        $this.ReceiveWebSocketMessage() | Out-Null

    }
}
