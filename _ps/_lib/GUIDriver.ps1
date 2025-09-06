# GUIDriverクラス
# GUIアプリケーションの自動操作を行うクラス

# 必要なアセンブリを読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# GUIDriverクラス
class GUIDriver
{
    # プロパティ
    [System.Diagnostics.Process]$process
    [IntPtr]$window_handle
    [string]$application_path
    [string]$window_title
    [bool]$is_initialized
    [string]$temp_directory

    # コンストラクタ
    GUIDriver()
    {
        try
        {
            $this.is_initialized = $false
            $this.process = $null
            $this.window_handle = [IntPtr]::Zero
            $this.application_path = ""
            $this.window_title = ""
            $this.temp_directory = ""
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0001", "GUIDriver初期化エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "GUIDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "GUIDriverの初期化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # 初期化・接続関連
    # ========================================

    # アプリケーション起動
    [void] StartApplication([string]$app_path, [string]$arguments = "")
    {
        try
        {
            if ([string]::IsNullOrEmpty($app_path))
            {
                throw "アプリケーションパスが指定されていません。"
            }

            if (-not (Test-Path $app_path))
            {
                throw "指定されたアプリケーションファイルが見つかりません: $app_path"
            }

            $this.application_path = $app_path

            # プロセス起動
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $app_path
            $processInfo.Arguments = $arguments
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $false

            $this.process = [System.Diagnostics.Process]::Start($processInfo)
            
            if (-not $this.process)
            {
                throw "アプリケーションの起動に失敗しました。"
            }

            # プロセスが完全に起動するまで待機
            Start-Sleep -Milliseconds 1000

            Write-Host "アプリケーションを起動しました: $app_path" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0002", "アプリケーション起動エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "アプリケーション起動エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "アプリケーションの起動に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウ検索
    [IntPtr] FindWindow([string]$window_title)
    {
        try
        {
            if ([string]::IsNullOrEmpty($window_title))
            {
                throw "ウィンドウタイトルが指定されていません。"
            }

            $this.window_title = $window_title

            # ウィンドウを検索
            $window_handle = [System.Windows.Forms.Control]::FromHandle([IntPtr]::Zero)
            $found_window = $null

            # プロセスに関連するウィンドウを検索
            if ($this.process -and -not $this.process.HasExited)
            {
                $this.process.Refresh()
                $main_window_handle = $this.process.MainWindowHandle
                
                if ($main_window_handle -ne [IntPtr]::Zero)
                {
                    $window_text = [System.Windows.Forms.Control]::FromHandle($main_window_handle).Text
                    if ($window_text -like "*$window_title*")
                    {
                        $this.window_handle = $main_window_handle
                        return $main_window_handle
                    }
                }
            }

            # タイムアウト付きでウィンドウ検索
            $timeout = 10000 # 10秒
            $elapsed = 0
            $interval = 500 # 500ms間隔

            while ($elapsed -lt $timeout)
            {
                if ($this.process -and -not $this.process.HasExited)
                {
                    $this.process.Refresh()
                    $main_window_handle = $this.process.MainWindowHandle
                    
                    if ($main_window_handle -ne [IntPtr]::Zero)
                    {
                        $window_text = [System.Windows.Forms.Control]::FromHandle($main_window_handle).Text
                        if ($window_text -like "*$window_title*")
                        {
                            $this.window_handle = $main_window_handle
                            Write-Host "ウィンドウを発見しました: $window_text" -ForegroundColor Green
                            return $main_window_handle
                        }
                    }
                }

                Start-Sleep -Milliseconds $interval
                $elapsed += $interval
            }

            throw "指定されたウィンドウが見つかりませんでした: $window_title"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0003", "ウィンドウ検索エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウ検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "指定されたウィンドウが見つかりませんでした: $($_.Exception.Message)"
        }
    }

    # ウィンドウアクティブ化
    [void] ActivateWindow()
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウを前面に表示
            [System.Windows.Forms.Control]::FromHandle($this.window_handle).BringToFront()
            [System.Windows.Forms.Control]::FromHandle($this.window_handle).Focus()

            Write-Host "ウィンドウをアクティブ化しました。" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0004", "ウィンドウアクティブ化エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウアクティブ化エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ウィンドウのアクティブ化に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # マウス操作関連
    # ========================================

    # マウスクリック
    [void] ClickMouse([int]$x, [int]$y, [string]$button = "Left")
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # マウス位置を設定
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
            Start-Sleep -Milliseconds 100

            # クリック実行
            switch ($button.ToLower())
            {
                "left" { [System.Windows.Forms.Cursor]::Click() }
                "right" { [System.Windows.Forms.Cursor]::Click([System.Windows.Forms.MouseButtons]::Right) }
                "middle" { [System.Windows.Forms.Cursor]::Click([System.Windows.Forms.MouseButtons]::Middle) }
                default { [System.Windows.Forms.Cursor]::Click() }
            }

            Write-Host "マウスクリックを実行しました: ($x, $y) $button" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0005", "マウスクリックエラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "マウスクリックエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "マウスクリックの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # マウスダブルクリック
    [void] DoubleClickMouse([int]$x, [int]$y)
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # マウス位置を設定
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
            Start-Sleep -Milliseconds 100

            # ダブルクリック実行
            [System.Windows.Forms.Cursor]::Click()
            Start-Sleep -Milliseconds 50
            [System.Windows.Forms.Cursor]::Click()

            Write-Host "マウスダブルクリックを実行しました: ($x, $y)" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0007", "マウスダブルクリックエラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "マウスダブルクリックエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "マウスダブルクリックの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # マウス右クリック
    [void] RightClickMouse([int]$x, [int]$y)
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # マウス位置を設定
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
            Start-Sleep -Milliseconds 100

            # 右クリック実行
            [System.Windows.Forms.Cursor]::Click([System.Windows.Forms.MouseButtons]::Right)

            Write-Host "マウス右クリックを実行しました: ($x, $y)" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0006", "マウス右クリックエラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "マウス右クリックエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "マウス右クリックの実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # マウス移動
    [void] MoveMouse([int]$x, [int]$y)
    {
        try
        {
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
            Start-Sleep -Milliseconds 100

            Write-Host "マウスを移動しました: ($x, $y)" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0008", "マウス移動エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "マウス移動エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "マウスの移動に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # キーボード操作関連
    # ========================================

    # キーボード入力
    [void] SendKeys([string]$keys)
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # キー送信
            [System.Windows.Forms.SendKeys]::SendWait($keys)

            Write-Host "キーボード入力を実行しました: $keys" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0009", "キーボード入力エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "キーボード入力エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "キーボード入力の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # 特殊キー送信
    [void] SendSpecialKey([string]$key)
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # 特殊キーのマッピング
            $specialKeys = @{
                "Enter" = "{ENTER}"
                "Tab" = "{TAB}"
                "Escape" = "{ESC}"
                "Space" = " "
                "Backspace" = "{BACKSPACE}"
                "Delete" = "{DELETE}"
                "Home" = "{HOME}"
                "End" = "{END}"
                "PageUp" = "{PGUP}"
                "PageDown" = "{PGDN}"
                "Up" = "{UP}"
                "Down" = "{DOWN}"
                "Left" = "{LEFT}"
                "Right" = "{RIGHT}"
                "F1" = "{F1}"
                "F2" = "{F2}"
                "F3" = "{F3}"
                "F4" = "{F4}"
                "F5" = "{F5}"
                "F6" = "{F6}"
                "F7" = "{F7}"
                "F8" = "{F8}"
                "F9" = "{F9}"
                "F10" = "{F10}"
                "F11" = "{F11}"
                "F12" = "{F12}"
            }

            if ($specialKeys.ContainsKey($key))
            {
                [System.Windows.Forms.SendKeys]::SendWait($specialKeys[$key])
            }
            else
            {
                throw "サポートされていない特殊キーです: $key"
            }

            Write-Host "特殊キーを送信しました: $key" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0010", "特殊キー送信エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "特殊キー送信エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "特殊キーの送信に失敗しました: $($_.Exception.Message)"
        }
    }

    # キー組み合わせ送信
    [void] SendKeyCombination([string[]]$keys)
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # キー組み合わせを構築
            $keyCombination = ($keys -join "+")
            [System.Windows.Forms.SendKeys]::SendWait($keyCombination)

            Write-Host "キー組み合わせを送信しました: $keyCombination" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0011", "キー組み合わせ送信エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "キー組み合わせ送信エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "キー組み合わせの送信に失敗しました: $($_.Exception.Message)"
        }
    }

    # テキスト入力
    [void] TypeText([string]$text)
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # テキスト入力
            [System.Windows.Forms.SendKeys]::SendWait($text)

            Write-Host "テキストを入力しました: $text" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0012", "テキスト入力エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テキスト入力エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "テキスト入力の実行に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # スクリーンショット関連
    # ========================================

    # スクリーンショット取得
    [void] TakeScreenshot([string]$file_path)
    {
        try
        {
            if ([string]::IsNullOrEmpty($file_path))
            {
                throw "ファイルパスが指定されていません。"
            }

            # スクリーンショット取得
            $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
            $bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.CopyFromScreen($screen.Left, $screen.Top, 0, 0, $screen.Size)
            $graphics.Dispose()

            # ファイル保存
            $bitmap.Save($file_path, [System.Drawing.Imaging.ImageFormat]::Png)
            $bitmap.Dispose()

            Write-Host "スクリーンショットを保存しました: $file_path" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0013", "スクリーンショット取得エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "スクリーンショット取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "スクリーンショットの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ウィンドウスクリーンショット取得
    [void] TakeWindowScreenshot([string]$file_path)
    {
        try
        {
            if ($this.window_handle -eq [IntPtr]::Zero)
            {
                throw "ウィンドウハンドルが設定されていません。"
            }

            if ([string]::IsNullOrEmpty($file_path))
            {
                throw "ファイルパスが指定されていません。"
            }

            # ウィンドウをアクティブ化
            $this.ActivateWindow()

            # ウィンドウの位置とサイズを取得
            $windowRect = New-Object System.Drawing.Rectangle
            [System.Windows.Forms.Control]::FromHandle($this.window_handle).Invoke([System.Action]{
                $windowRect = [System.Windows.Forms.Control]::FromHandle($this.window_handle).Bounds
            })

            # スクリーンショット取得
            $bitmap = New-Object System.Drawing.Bitmap($windowRect.Width, $windowRect.Height)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.CopyFromScreen($windowRect.Left, $windowRect.Top, 0, 0, $windowRect.Size)
            $graphics.Dispose()

            # ファイル保存
            $bitmap.Save($file_path, [System.Drawing.Imaging.ImageFormat]::Png)
            $bitmap.Dispose()

            Write-Host "ウィンドウスクリーンショットを保存しました: $file_path" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0014", "ウィンドウスクリーンショット取得エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "ウィンドウスクリーンショット取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "ウィンドウスクリーンショットの取得に失敗しました: $($_.Exception.Message)"
        }
    }

    # ========================================
    # プロセス管理関連
    # ========================================

    # プロセス終了
    [void] CloseApplication()
    {
        try
        {
            if ($this.process -and -not $this.process.HasExited)
            {
                # 通常の終了を試行
                $this.process.CloseMainWindow()
                $this.process.WaitForExit(5000)
                
                # 強制終了が必要な場合
                if (-not $this.process.HasExited)
                {
                    $this.process.Kill()
                    $this.process.WaitForExit(5000)
                }

                Write-Host "アプリケーションを終了しました。" -ForegroundColor Green
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0015", "プロセス終了エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "プロセス終了エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "プロセスの終了に失敗しました: $($_.Exception.Message)"
        }
    }

    # プロセス強制終了
    [void] KillApplication()
    {
        try
        {
            if ($this.process -and -not $this.process.HasExited)
            {
                $this.process.Kill()
                $this.process.WaitForExit(5000)

                Write-Host "アプリケーションを強制終了しました。" -ForegroundColor Green
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0016", "プロセス強制終了エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "プロセス強制終了エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "プロセスの強制終了に失敗しました: $($_.Exception.Message)"
        }
    }

    # プロセス状態確認
    [bool] IsProcessRunning()
    {
        try
        {
            if ($this.process -and -not $this.process.HasExited)
            {
                return $true
            }
            return $false
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0017", "プロセス状態確認エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "プロセス状態確認エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            return $false
        }
    }

    # 待機
    [void] Wait([int]$milliseconds)
    {
        try
        {
            Start-Sleep -Milliseconds $milliseconds
            Write-Host "待機しました: $milliseconds ms" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0019", "タイムアウトエラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "タイムアウトエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "待機処理に失敗しました: $($_.Exception.Message)"
        }
    }

    # デストラクタ
    [void] Dispose()
    {
        try
        {
            if ($this.process -and -not $this.process.HasExited)
            {
                $this.process.CloseMainWindow()
                $this.process.WaitForExit(5000)
                
                if (-not $this.process.HasExited)
                {
                    $this.process.Kill()
                }
            }

            $this.is_initialized = $false
            Write-Host "GUIDriverを破棄しました。" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("GUIError_0018", "GUIDriver破棄エラー: $($_.Exception.Message)", "GUIDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "GUIDriver破棄エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
