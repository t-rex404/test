# UIAutomationDriverクラス
# UI Automationを使用してGUIアプリケーションの自動操作を行うクラス

# 必要なアセンブリを読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# UIAutomationDriverクラス
class UIAutomationDriver
{
    # プロパティ
    [System.Diagnostics.Process]$process
    [System.Windows.Automation.AutomationElement]$root_element
    [string]$application_path
    [string]$window_title
    [bool]$is_initialized
    [string]$temp_directory

    # コンストラクタ
    UIAutomationDriver()
    {
        try
        {
            $this.is_initialized = $false
            $this.process = $null
            $this.root_element = $null
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
                    $global:Common.HandleError("UIAError_0001", "UIAutomationDriver初期化エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "UIAutomationDriverの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "UIAutomationDriverの初期化に失敗しました: $($_.Exception.Message)"
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

            # システムアプリケーションの場合はGet-Commandでパスを取得
            $actual_path = $app_path
            if (-not (Test-Path $app_path))
            {
                try
                {
                    $command = Get-Command $app_path -ErrorAction Stop
                    $actual_path = $command.Source
                    Write-Host "システムアプリケーションのパスを取得しました: $actual_path" -ForegroundColor Yellow
                }
                catch
                {
                    throw "指定されたアプリケーションファイルが見つかりません: $app_path"
                }
            }

            $this.application_path = $actual_path

            # プロセス起動
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $actual_path
            $processInfo.Arguments = $arguments
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $false

            $this.process = [System.Diagnostics.Process]::Start($processInfo)
            
            if (-not $this.process)
            {
                throw "アプリケーションの起動に失敗しました。"
            }

            # プロセスが完全に起動するまで待機
            Start-Sleep -Milliseconds 2000

            Write-Host "アプリケーションを起動しました: $app_path" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0002", "アプリケーション起動エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
    [System.Windows.Automation.AutomationElement] FindWindow([string]$window_title)
    {
        try
        {
            if ([string]::IsNullOrEmpty($window_title))
            {
                throw "ウィンドウタイトルが指定されていません。"
            }

            $this.window_title = $window_title

            # デスクトップのルート要素を取得
            $desktop = [System.Windows.Automation.AutomationElement]::RootElement

            # ウィンドウを検索する条件
            $condition = New-Object System.Windows.Automation.AndCondition(
                (New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                    [System.Windows.Automation.ControlType]::Window
                )),
                (New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::NameProperty,
                    $window_title
                ))
            )

            # タイムアウト付きでウィンドウ検索
            $timeout = 10000 # 10秒
            $elapsed = 0
            $interval = 500 # 500ms間隔

            while ($elapsed -lt $timeout)
            {
                # プロセスに関連するウィンドウを検索
                if ($this.process -and -not $this.process.HasExited)
                {
                    $processCondition = New-Object System.Windows.Automation.AndCondition(
                        $condition,
                        (New-Object System.Windows.Automation.PropertyCondition(
                            [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
                            $this.process.Id
                        ))
                    )

                    $window = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $processCondition)
                    if ($window -ne $null)
                    {
                        $this.root_element = $window
                        Write-Host "ウィンドウを発見しました: $window_title" -ForegroundColor Green
                        return $window
                    }
                }

                # プロセスIDに関係なくウィンドウを検索
                $window = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
                if ($window -ne $null)
                {
                    $this.root_element = $window
                    Write-Host "ウィンドウを発見しました: $window_title" -ForegroundColor Green
                    return $window
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
                    $global:Common.HandleError("UIAError_0003", "ウィンドウ検索エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
            if ($this.root_element -eq $null)
            {
                throw "ウィンドウ要素が設定されていません。"
            }

            # ウィンドウパターンを取得してアクティブ化
            $windowPattern = $this.root_element.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
            if ($windowPattern -ne $null)
            {
                $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
            }

            # ウィンドウを前面に表示
            $transformPattern = $this.root_element.GetCurrentPattern([System.Windows.Automation.TransformPattern]::Pattern)
            if ($transformPattern -ne $null)
            {
                $transformPattern.Move(0, 0)
            }

            Write-Host "ウィンドウをアクティブ化しました。" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0004", "ウィンドウアクティブ化エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
    # 要素検索・操作関連
    # ========================================

    # 要素検索（名前で）
    [System.Windows.Automation.AutomationElement] FindElementByName([string]$name)
    {
        try
        {
            if ($this.root_element -eq $null)
            {
                throw "ルート要素が設定されていません。"
            }

            $condition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::NameProperty,
                $name
            )

            $element = $this.root_element.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
            if ($element -eq $null)
            {
                throw "指定された要素が見つかりませんでした: $name"
            }

            Write-Host "要素を発見しました: $name" -ForegroundColor Green
            return $element
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0005", "要素検索エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "指定された要素が見つかりませんでした: $($_.Exception.Message)"
        }
    }

    # 要素検索（コントロールタイプで）
    [System.Windows.Automation.AutomationElement] FindElementByControlType([System.Windows.Automation.ControlType]$controlType)
    {
        try
        {
            if ($this.root_element -eq $null)
            {
                throw "ルート要素が設定されていません。"
            }

            $condition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                $controlType
            )

            $element = $this.root_element.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
            if ($element -eq $null)
            {
                throw "指定されたコントロールタイプの要素が見つかりませんでした: $controlType"
            }

            Write-Host "要素を発見しました: $controlType" -ForegroundColor Green
            return $element
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0006", "要素検索エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "指定されたコントロールタイプの要素が見つかりませんでした: $controlType"
        }
    }

    # 要素検索（複合条件）
    [System.Windows.Automation.AutomationElement] FindElement([hashtable]$conditions)
    {
        try
        {
            if ($this.root_element -eq $null)
            {
                throw "ルート要素が設定されていません。"
            }

            $conditionList = @()
            foreach ($key in $conditions.Keys)
            {
                $conditionList += New-Object System.Windows.Automation.PropertyCondition($key, $conditions[$key])
            }

            $andCondition = New-Object System.Windows.Automation.AndCondition($conditionList)
            $element = $this.root_element.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $andCondition)
            
            if ($element -eq $null)
            {
                throw "指定された条件の要素が見つかりませんでした"
            }

            Write-Host "要素を発見しました" -ForegroundColor Green
            return $element
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0007", "要素検索エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素検索エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "指定された条件の要素が見つかりませんでした: $($_.Exception.Message)"
        }
    }

    # 要素クリック
    [void] ClickElement([System.Windows.Automation.AutomationElement]$element)
    {
        try
        {
            if ($element -eq $null)
            {
                throw "要素が指定されていません。"
            }

            # 要素を表示状態にする
            $element.SetFocus()

            # InvokePatternを取得してクリック
            $invokePattern = $element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
            if ($invokePattern -ne $null)
            {
                $invokePattern.Invoke()
            }
            else
            {
                # InvokePatternが利用できない場合は、要素の中心をクリック
                $clickablePoint = $element.GetClickablePoint()
                $this.ClickMouse([int]$clickablePoint.X, [int]$clickablePoint.Y)
            }

            Write-Host "要素をクリックしました" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0008", "要素クリックエラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "要素クリックエラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "要素のクリックに失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素にテキスト入力
    [void] SetElementText([System.Windows.Automation.AutomationElement]$element, [string]$text)
    {
        try
        {
            if ($element -eq $null)
            {
                throw "要素が指定されていません。"
            }

            # 要素をフォーカス
            $element.SetFocus()

            # ValuePatternを取得してテキスト設定
            $valuePattern = $element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
            if ($valuePattern -ne $null)
            {
                $valuePattern.SetValue($text)
            }
            else
            {
                # ValuePatternが利用できない場合は、キーボード入力を使用
                $this.TypeText($text)
            }

            Write-Host "要素にテキストを設定しました: $text" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0009", "テキスト設定エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テキスト設定エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            throw "要素へのテキスト設定に失敗しました: $($_.Exception.Message)"
        }
    }

    # 要素のテキスト取得
    [string] GetElementText([System.Windows.Automation.AutomationElement]$element)
    {
        try
        {
            if ($element -eq $null)
            {
                throw "要素が指定されていません。"
            }

            $text = $element.GetCurrentPropertyValue([System.Windows.Automation.AutomationElement]::NameProperty)
            if ([string]::IsNullOrEmpty($text))
            {
                $text = $element.GetCurrentPropertyValue([System.Windows.Automation.AutomationElement]::ValueProperty)
            }

            Write-Host "要素のテキストを取得しました: $text" -ForegroundColor Green
            return $text
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0010", "テキスト取得エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "テキスト取得エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
            
            return ""
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
                    $global:Common.HandleError("UIAError_0011", "マウスクリックエラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0012", "マウスダブルクリックエラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0013", "マウス右クリックエラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0014", "マウス移動エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0015", "キーボード入力エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0016", "特殊キー送信エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0017", "キー組み合わせ送信エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0018", "テキスト入力エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0019", "スクリーンショット取得エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
            if ($this.root_element -eq $null)
            {
                throw "ウィンドウ要素が設定されていません。"
            }

            if ([string]::IsNullOrEmpty($file_path))
            {
                throw "ファイルパスが指定されていません。"
            }

            # ウィンドウの位置とサイズを取得
            $windowRect = $this.root_element.Current.BoundingRectangle

            # スクリーンショット取得
            $bitmap = New-Object System.Drawing.Bitmap([int]$windowRect.Width, [int]$windowRect.Height)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.CopyFromScreen([int]$windowRect.Left, [int]$windowRect.Top, 0, 0, [System.Drawing.Size]::new([int]$windowRect.Width, [int]$windowRect.Height))
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
                    $global:Common.HandleError("UIAError_0020", "ウィンドウスクリーンショット取得エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0021", "プロセス終了エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0022", "プロセス強制終了エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0023", "プロセス状態確認エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
                    $global:Common.HandleError("UIAError_0025", "タイムアウトエラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
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
            $this.root_element = $null
            Write-Host "UIAutomationDriverを破棄しました。" -ForegroundColor Green
        }
        catch
        {
            # Commonオブジェクトが利用可能な場合はエラーログに記録
            if ($global:Common)
            {
                try
                {
                    $global:Common.HandleError("UIAError_0024", "UIAutomationDriver破棄エラー: $($_.Exception.Message)", "UIAutomationDriver", ".\AllDrivers_Error.log")
                }
                catch
                {
                    Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "UIAutomationDriver破棄エラー: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
