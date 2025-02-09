$Host.UI.RawUI.WindowTitle = 'TEST'


Function Main()
{
    param
    (
        
    )

    begin
    {
        # 定数宣言
        new-variable -name SUCCESS_CODE -value 0 -Description '正常終了コード' -Option Constant -visibility Private
        new-variable -name ERROR_CODE -value 255 -Description '異常終了コード' -Option Constant -visibility Private

        # 変数宣言
        new-variable -name result -value $ERROR_CODE -Description '結果' -Option Private -visibility Private
    }

    process
    {
        try
        {
            # ブラウザの実行ファイルのパスを取得
            $browser_exe_reg_key = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\'

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

            $url = 'https://github.com/t-rex404/test/archive/refs/heads/main.zip'
            $argument_list = '--new-window ' + $url
            # ブラウザを開く
            $browser_exe_process = Start-Process -FilePath $browser_exe_path -ArgumentList $argument_list -WindowStyle Minimized -PassThru
            $browser_exe_process_id = $browser_exe_process.Id
            Write-Host "`$browser_exe_process_id:$browser_exe_process_id"

            $shellapp = New-Object -ComObject Shell.Application
            $download_dir_path = $shellapp.Namespace("shell:Downloads").Self.Path
            Write-Host "`$download_dir_path:$download_dir_path"
            $download_file = $download_dir_path + "\test-main.zip"
            Write-Host "`$download_file:$download_file"

            #ダウンロード待機
            $timeout = [datetime]::Now.AddSeconds(5)
            try
            {
                while ($true)
                {
                    if ([datetime]::Now -gt $timeout)
                    {
                        throw 'ページロード完了待機タイムアウト'
                    }
                    if (Test-Path $download_file)
                    {
                        break
                    }
                    Start-Sleep -Milliseconds 500
                }
            }
            catch
            {
                throw 'ページロード完了待機に失敗しました。エラーメッセージ：' + $_
            }

            #Get-Process -Name msedge
            #Start-Sleep -Milliseconds 2000
            #Stop-Process -Id $browser_exe_process_id -Force
            Stop-Process -Name msedge -Force



            $result = $SUCCESS_CODE
        }
        catch
        {
            #Do this if a terminating exception happens
            write-host 'catch'
            write-host $_.Exception.Message
        }
        finally
        {
            #Do this no matter what
            write-host 'finally'
        }
        return $result
    }
    end
    {
        # 終了処理
        #remove-variable -name 'EDGE_EXE_REG_KEY' -ErrorAction SilentlyContinue
        #remove-variable -name 'result'           -ErrorAction SilentlyContinue
        #remove-variable -name 'edge_exe_path'    -ErrorAction SilentlyContinue
        #remove-variable -name 'process'          -ErrorAction SilentlyContinue    
        #remove-variable -name 'process_id'       -ErrorAction SilentlyContinue
    }
}

$result = Main
if ($result -eq 0)
{
    exit 0
}
else
{
    exit 1
}
