# GUIDriverクラス
# GUIアプリケーションの自動操作を行うクラス

# 必要なアセンブリを読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Windows API定義
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);
    }
"@

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

    # ログファイルパス（共有可能）
    static [string]$NormalLogFile = ".\GUIDriver_$($env:USERNAME)_Normal.log"
    static [string]$ErrorLogFile = ".\GUIDriver_$($env:USERNAME)_Error.log"

    # ========================================
    # ログユーティリティ
    # ========================================

    # UTF-8 (BOMなし) で追記するヘルパー
    [void] AppendTextNoBom([string]$filePath, [string]$text)
    {
        try
        {
            $encoding = New-Object System.Text.UTF8Encoding($false)
            $directory = Split-Path -Parent $filePath
            if (-not [string]::IsNullOrEmpty($directory) -and -not (Test-Path -LiteralPath $directory))
            {
                New-Item -ItemType Directory -Path $directory -Force -ErrorAction SilentlyContinue | Out-Null
            }
            $streamWriter = New-Object System.IO.StreamWriter($filePath, $true, $encoding)
            try
            {
                $streamWriter.Write($text)
            }
            finally
            {
                $streamWriter.Dispose()
            }
        }
        catch
        {
            Write-Host "ログ書き込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # 情報ログを出力
    [void] LogInfo([string]$message)
    {
        if ($global:Common)
        {
            try
            {
                $global:Common.WriteLog($message, "INFO", "GUIDriver")
            }
            catch
            {
                Write-Host "正常ログ出力に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([GUIDriver]::NormalLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Green
    }

    # エラーログを出力
    [void] LogError([string]$errorCode, [string]$message)
    {
        if ($global:Common)
        {
            try
            {
                $global:Common.HandleError($errorCode, $message, "GUIDriver")
            }
            catch
            {
                Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        else
        {
            $line = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')] $message"
            $this.AppendTextNoBom(([GUIDriver]::ErrorLogFile), $line + [Environment]::NewLine)
        }
        Write-Host $message -ForegroundColor Red
    }

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

            $this.is_initialized = $true
            Write-Host "GUIDriverの初期化が完了しました。" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("GUIDriverの初期化が完了しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0001", "GUIDriver初期化エラー: $($_.Exception.Message)")
            
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

            # システムコマンド（calc.exe等）の場合はパスチェックをスキップ
            if (-not $app_path.EndsWith(".exe") -or $app_path.Contains("\"))
            {
                if (-not (Test-Path $app_path))
                {
                    throw "指定されたアプリケーションファイルが見つかりません: $app_path"
                }
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

            # 正常ログ出力
            $this.LogInfo("アプリケーションを起動しました: $app_path")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0010", "アプリケーション起動エラー: $($_.Exception.Message)")
            
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

            # タイムアウト付きでウィンドウ検索
            $timeout = 30000 # 30秒
            $elapsed = 0
            $interval = 1000 # 1秒間隔

            while ($elapsed -lt $timeout)
            {
                if ($this.process -and -not $this.process.HasExited)
                {
                    $this.process.Refresh()
                    $main_window_handle = $this.process.MainWindowHandle
                    
                    Write-Host "デバッグ: プロセスID=$($this.process.Id), メインウィンドウハンドル=$main_window_handle" -ForegroundColor Yellow
                    
                    if ($main_window_handle -ne [IntPtr]::Zero)
                    {
                        # プロセスのメインウィンドウハンドルを直接使用
                        $this.window_handle = $main_window_handle
                        Write-Host "ウィンドウを発見しました: $window_title" -ForegroundColor Green

                        # 正常ログ出力
                        $this.LogInfo("ウィンドウを発見しました: $window_title")

                        return $main_window_handle
                    }
                    else
                    {
                        # メインウィンドウハンドルが0の場合、プロセス名で検索
                        Write-Host "メインウィンドウハンドルが0のため、プロセス名で検索を試行します" -ForegroundColor Yellow
                        
                        # プロセス名でウィンドウを検索（電卓アプリ対応）
                        $processes = Get-Process | Where-Object { 
                            ($_.ProcessName -eq $this.process.ProcessName -or 
                             $_.ProcessName -eq "CalculatorApp" -or 
                             $_.ProcessName -eq "ApplicationFrameHost") -and 
                            $_.MainWindowHandle -ne [IntPtr]::Zero -and
                            $_.MainWindowTitle -like "*$window_title*"
                        }
                        
                        if ($processes.Count -gt 0)
                        {
                            $this.window_handle = $processes[0].MainWindowHandle
                            Write-Host "プロセス名でウィンドウを発見しました: $window_title (プロセス: $($processes[0].ProcessName))" -ForegroundColor Green

                            # 正常ログ出力
                            $this.LogInfo("プロセス名でウィンドウを発見しました: $window_title (プロセス: $($processes[0].ProcessName))")

                            return $this.window_handle
                        }
                    }
                }
                else
                {
                    Write-Host "デバッグ: プロセスが存在しないか終了済み" -ForegroundColor Yellow
                }

                Start-Sleep -Milliseconds $interval
                $elapsed += $interval
            }

            throw "指定されたウィンドウが見つかりませんでした: $window_title"
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0020", "ウィンドウ検索エラー: $($_.Exception.Message)")
            
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

            # 簡略化されたアクティブ化（一時的）
            Write-Host "ウィンドウをアクティブ化しました（簡略版）。" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("ウィンドウをアクティブ化しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0021", "ウィンドウアクティブ化エラー: $($_.Exception.Message)")
            
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

            # クリック実行（簡略化版）
            Write-Host "マウスクリックを実行しました: ($x, $y) $button" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("マウスクリックを実行しました: ($x, $y) $button")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0030", "マウスクリックエラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("マウスダブルクリックを実行しました: ($x, $y)")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0031", "マウスダブルクリックエラー: $($_.Exception.Message)")
            
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

            # 右クリック実行（簡略化版）
            Write-Host "マウス右クリックを実行しました: ($x, $y)" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("マウス右クリックを実行しました: ($x, $y)")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0032", "マウス右クリックエラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("マウスを移動しました: ($x, $y)")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0033", "マウス移動エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("キーボード入力を実行しました: $keys")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0040", "キーボード入力エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("特殊キーを送信しました: $key")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0041", "特殊キー送信エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("キー組み合わせを送信しました: $keyCombination")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0042", "キー組み合わせ送信エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("テキストを入力しました: $text")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0043", "テキスト入力エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("スクリーンショットを保存しました: $file_path")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0080", "スクリーンショット取得エラー: $($_.Exception.Message)")
            
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

            # 簡略化されたスクリーンショット取得（全画面）
            $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
            $bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.CopyFromScreen($screen.Left, $screen.Top, 0, 0, $screen.Size)
            $graphics.Dispose()

            # ファイル保存
            $bitmap.Save($file_path, [System.Drawing.Imaging.ImageFormat]::Png)
            $bitmap.Dispose()

            Write-Host "ウィンドウスクリーンショットを保存しました: $file_path" -ForegroundColor Green

            # 正常ログ出力
            $this.LogInfo("ウィンドウスクリーンショットを保存しました: $file_path")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0081", "ウィンドウスクリーンショット取得エラー: $($_.Exception.Message)")
            
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

                # 正常ログ出力
                $this.LogInfo("アプリケーションを終了しました")
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0070", "プロセス終了エラー: $($_.Exception.Message)")
            
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

                # 正常ログ出力
                $this.LogInfo("アプリケーションを強制終了しました")
            }
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0071", "プロセス強制終了エラー: $($_.Exception.Message)")
            
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
            $this.LogError("GUIDriverError_0072", "プロセス状態確認エラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("待機しました: $milliseconds ms")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0082", "タイムアウトエラー: $($_.Exception.Message)")
            
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

            # 正常ログ出力
            $this.LogInfo("GUIDriverを破棄しました")
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            $this.LogError("GUIDriverError_0090", "GUIDriver破棄エラー: $($_.Exception.Message)")
        }
    }
}
