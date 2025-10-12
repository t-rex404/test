# 共通ライブラリ
# 各ドライバークラスで使用する共通機能を提供

class Common : IDisposable
{
    [hashtable]$ErrorCodes = @{}
    [string]$TempDir
    [string]$LogFilePath
    [string]$ErrorLogFilePath
    [bool]$Disposed = $false

    # コンストラクタ
    Common()
    {
        # 一時ディレクトリとログの初期化
        try
        {
            $this.SetupTempDirectory()
            $this.InitializeLogFile()
            $this.InitializeErrorLogFile()
        }
        catch
        {
            Write-Host "一時ディレクトリ/ログの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # エラーコードの読み込み
        try
        {
            $this.LoadErrorCodes()
        }
        catch
        {
            Write-Host "エラーコードの読み込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
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
        
        # ログファイルに出力（インスタンス生成時に決定した一時ディレクトリ配下のファイル）
        $logFile = $this.LogFilePath
        if ([string]::IsNullOrEmpty($logFile))
        {
            $logFile = ".\Common_Info.log"
        }
        $logMessage | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    # リソース解放とログ退避
    [void] Dispose()
    {
        if ($this.Disposed)
        {
            return
        }

        try
        {
            # ログファイルが存在する場合は実行ディレクトリの _log 配下へコピー
            $executionDir = $PWD.Path
            $destDir = Join-Path $executionDir "_log"
            if (-not (Test-Path -LiteralPath $destDir))
            {
                New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null
            }

            if ($this.LogFilePath -and (Test-Path -LiteralPath $this.LogFilePath))
            {
                try
                {
                    $logFile = Get-Item -LiteralPath $this.LogFilePath -ErrorAction Stop
                    if ($logFile.Length -eq 0)
                    {
                        Remove-Item -LiteralPath $this.LogFilePath -Force -ErrorAction Stop
                        Write-Host "空のログファイルを削除しました: $($logFile.FullName)" -ForegroundColor Yellow
                    }
                    else
                    {
                        $logFileName = $logFile.Name
                        $destFile = Join-Path $destDir $logFileName
                        Copy-Item -LiteralPath $this.LogFilePath -Destination $destFile -Force -ErrorAction Stop
                        Write-Host "ログファイルを退避しました: $destFile" -ForegroundColor Green
                    }
                }
                catch
                {
                    Write-Host "ログファイルの退避に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # エラーログファイルが存在する場合も退避
            if ($this.ErrorLogFilePath -and (Test-Path -LiteralPath $this.ErrorLogFilePath))
            {
                try
                {
                    $errorLog = Get-Item -LiteralPath $this.ErrorLogFilePath -ErrorAction Stop
                    if ($errorLog.Length -eq 0)
                    {
                        Remove-Item -LiteralPath $this.ErrorLogFilePath -Force -ErrorAction Stop
                        Write-Host "空のエラーログファイルを削除しました: $($errorLog.FullName)" -ForegroundColor Yellow
                    }
                    else
                    {
                        $errFileName = $errorLog.Name
                        $destErrFile = Join-Path $destDir $errFileName
                        Copy-Item -LiteralPath $this.ErrorLogFilePath -Destination $destErrFile -Force -ErrorAction Stop
                        Write-Host "エラーログファイルを退避しました: $destErrFile" -ForegroundColor Green
                    }
                }
                catch
                {
                    Write-Host "エラーログファイルの退避に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            # 一時ディレクトリを削除
            if ($this.TempDir -and (Test-Path -LiteralPath $this.TempDir))
            {
                try
                {
                    Remove-Item -Path $this.TempDir -Recurse -Force -ErrorAction Stop
                    Write-Host "一時ディレクトリを削除しました: $($this.TempDir)" -ForegroundColor Yellow
                }
                catch
                {
                    Write-Host "一時ディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        finally
        {
            $this.Disposed = $true
        }
    }
    
    # 共通のエラーハンドリング関数（拡張版）
    [void] HandleError([string]$errorCode, [string]$message, [string]$module = "Common", [string]$logFile = "")
    {
        $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        
        # エラータイトルを取得
        $error_title = $this.GetErrorTitle($errorCode, $module)
        
        # ログファイル名を決定（一時ディレクトリ配下、インスタンス生成時の命名に従う）
        if ([string]::IsNullOrEmpty($logFile))
        {
            $logFile = $this.ErrorLogFilePath
            if ([string]::IsNullOrEmpty($logFile))
            {
                $logFile = $this.LogFilePath
                if ([string]::IsNullOrEmpty($logFile))
                {
                    $logFile = ".\Common_Error.log"
                }
            }
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

    # 一時ディレクトリを初期化
    [void] SetupTempDirectory()
    {
        # ベースディレクトリを決定（優先: C:\temp → 環境変数TEMP）
        $base_dirs = @(
            "C:\temp",
            "$($env:TEMP)"
        )
        $base_dir = $base_dirs[0]
        if (-not (Test-Path $base_dir))
        {
            $base_dir = $base_dirs[1]
        }

        $temp_dir = Join-Path $base_dir "Common_Temp"

        # 既存のディレクトリがある場合は削除
        if (Test-Path $temp_dir)
        {
            try
            {
                Remove-Item -Path $temp_dir -Recurse -Force -ErrorAction Stop
                Write-Host "既存の一時ディレクトリを削除しました: $temp_dir" -ForegroundColor Yellow
            }
            catch
            {
                Write-Host "既存の一時ディレクトリの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        # 新しいディレクトリを作成
        New-Item -ItemType Directory -Path $temp_dir -Force -ErrorAction Stop | Out-Null
        $this.TempDir = $temp_dir
        Write-Host "Common一時ディレクトリを作成しました: $temp_dir" -ForegroundColor Green
    }

    # ログファイル（ファイル名）を初期化
    [void] InitializeLogFile()
    {
        $machineName = $env:COMPUTERNAME
        $username = $env:USERNAME
        $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
        $fileName = "${machineName}_${username}_${timestamp}.log"

        if ([string]::IsNullOrEmpty($this.TempDir))
        {
            $this.SetupTempDirectory()
        }

        $this.LogFilePath = Join-Path $this.TempDir $fileName

        # 空ファイルを作成（以後はAppendで書き込み）
        try
        {
            New-Item -ItemType File -Path $this.LogFilePath -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Host "ログファイルを初期化しました: $($this.LogFilePath)" -ForegroundColor Green
        }
        catch
        {
            Write-Host "ログファイルの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # エラーログファイル（ファイル名）を初期化
    [void] InitializeErrorLogFile()
    {
        if ([string]::IsNullOrEmpty($this.TempDir))
        {
            $this.SetupTempDirectory()
        }

        # 既存の通常ログ名があれば、それに基づきエラーログ名を決定
        if ($this.LogFilePath -and ($this.LogFilePath.ToLower().EndsWith('.log')))
        {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($this.LogFilePath)
            $errorFileName = "$baseName`_error.log"
            $this.ErrorLogFilePath = Join-Path $this.TempDir $errorFileName
        }
        else
        {
            $machineName = $env:COMPUTERNAME
            $username = $env:USERNAME
            $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
            $this.ErrorLogFilePath = Join-Path $this.TempDir "${machineName}_${username}_${timestamp}_error.log"
        }

        try
        {
            New-Item -ItemType File -Path $this.ErrorLogFilePath -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Host "エラーログファイルを初期化しました: $($this.ErrorLogFilePath)" -ForegroundColor Green
        }
        catch
        {
            Write-Host "エラーログファイルの初期化に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # JSONファイルからエラーコードを読み込む
    [void] LoadErrorCodes()
    {
        try
        {
            # JSONファイルのパスを取得
            $jsonPath = Join-Path $PSScriptRoot "..\..\_json\ErrorCode.json"
            
            if (-not (Test-Path $jsonPath))
            {
                Write-Host "エラーコードJSONファイルが見つかりません: $jsonPath" -ForegroundColor Yellow
                return
            }

            # JSONファイルを読み込み
            $jsonContent = Get-Content $jsonPath -Raw -Encoding UTF8
            $jsonObject = $jsonContent | ConvertFrom-Json
            
            # PowerShell 5.1互換のハッシュテーブル変換
            $this.ErrorCodes = @{}
            foreach ($module in $jsonObject.PSObject.Properties.Name)
            {
                $moduleKey = $module.ToUpper()
                $this.ErrorCodes[$moduleKey] = @{}
                foreach ($errorCode in $jsonObject.$module.PSObject.Properties.Name)
                {
                    $this.ErrorCodes[$moduleKey][$errorCode] = $jsonObject.$module.$errorCode
                }
            }

            Write-Host "エラーコードを正常に読み込みました。" -ForegroundColor Green
        }
        catch
        {
            Write-Host "エラーコードの読み込み中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }

    # エラータイトルを取得する関数
    [string] GetErrorTitle([string]$errorCode, [string]$module)
    {
        try
        {
            # モジュール名を正規化
            $moduleKey = $module.ToUpper()
            
            # エラーコードデータが読み込まれているかチェック
            if ($this.ErrorCodes.Count -eq 0)
            {
                Write-Host "エラーコードデータが読み込まれていません。" -ForegroundColor Yellow
                return "不明なエラー"
            }

            # 指定されたモジュールのエラーコードを取得
            if ($this.ErrorCodes.ContainsKey($moduleKey))
            {
                $moduleErrors = $this.ErrorCodes[$moduleKey]
                
                # 指定されたエラーコードのタイトルを取得
                if ($moduleErrors.ContainsKey($errorCode))
                {
                    return $moduleErrors[$errorCode]
                }
            }

            return "不明なエラー"
        }
        catch
        {
            Write-Host "エラータイトル取得中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
            return "不明なエラー"
        }
    }
}

# 共通インスタンスを作成（グローバルスコープ）
$global:Common = [Common]::new()

Write-Host "Commonライブラリが正常にインポートされました。" -ForegroundColor Green
