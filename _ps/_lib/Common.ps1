# 共通ライブラリ
# 各ドライバークラスで使用する共通機能を提供

class Common
{
    [hashtable]$ErrorCodes = @{}

    # コンストラクタでJSONファイルを読み込む
    Common()
    {
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
                $this.ErrorCodes[$module] = @{}
                foreach ($errorCode in $jsonObject.$module.PSObject.Properties.Name)
                {
                    $this.ErrorCodes[$module][$errorCode] = $jsonObject.$module.$errorCode
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