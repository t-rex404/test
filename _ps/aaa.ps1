$script:WinSCPAssemblyLoaded = $false
if (-not $script:WinSCPAssemblyLoaded)
{
    # WinSCPのインストールパスを検索
    $possiblePaths = @(
        "${env:ProgramFiles}\WinSCP\WinSCPnet.dll",
        "${env:ProgramFiles(x86)}\WinSCP\WinSCPnet.dll",
        "C:\Program Files\WinSCP\WinSCPnet.dll",
        "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"
    )

    $assemblyPath = $null
    foreach ($path in $possiblePaths)
    {
        if (Test-Path $path)
        {
            $assemblyPath = $path
            break
        }
    }

    if ($assemblyPath)
    {
        try
        {
            Add-Type -Path $assemblyPath
            $script:WinSCPAssemblyLoaded = $true
            Write-Host "WinSCPアセンブリを読み込みました: $assemblyPath" -ForegroundColor Green
        }
        catch
        {
            Write-Host "WinSCPアセンブリの読み込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "WinSCPがインストールされていることを確認してください。" -ForegroundColor Yellow
            throw
        }
    }
    else
    {
        Write-Host "WinSCPが見つかりません。WinSCPをインストールしてください。" -ForegroundColor Red
        Write-Host "ダウンロード: https://winscp.net/" -ForegroundColor Yellow
        throw "WinSCPアセンブリが見つかりません"
    }
}


# WinSCPDriverクラスの全メソッドテスト
# 実行前に確認事項:
# 1. WinSCPがインストールされていること
# 2. テスト用のFTP/SFTPサーバーが利用可能であること
# 3. 接続情報を環境に合わせて設定すること

# パスの設定
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$libPath = Join-Path $scriptPath "_lib"

# 必要なクラスを読み込み
. (Join-Path $libPath "Common.ps1")
. (Join-Path $libPath "TeraTermDriver.ps1")
. (Join-Path $libPath "WinSCPDriver.ps1")
