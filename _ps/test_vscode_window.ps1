# VSCodeウィンドウ検索テスト

# 必要なクラスを読み込み
. "$PSScriptRoot\_lib\Common.ps1"
. "$PSScriptRoot\_lib\UIAutomationDriver.ps1"

# グローバル変数を初期化
$global:Common = [Common]::new()

Write-Host "`n=====================================" -ForegroundColor Cyan
Write-Host "VSCode ウィンドウ検索テスト" -ForegroundColor Cyan
Write-Host "=====================================`n" -ForegroundColor Cyan

try
{
    # UIAutomationDriverの初期化
    $uiDriver = [UIAutomationDriver]::new()
    Write-Host "UIAutomationDriver初期化完了" -ForegroundColor Green

    # VSCodeのウィンドウを検索（様々なパターンでテスト）
    $testPatterns = @(
        "Visual Studio Code",
        "VSCode",
        "vscode",
        "Code",
        "Visual Studio",
        "*Code*"
    )

    $foundWindow = $null

    foreach ($pattern in $testPatterns)
    {
        Write-Host "`nパターン '$pattern' でウィンドウを検索中..." -ForegroundColor Yellow

        try
        {
            $foundWindow = $uiDriver.FindWindow($pattern)

            if ($null -ne $foundWindow)
            {
                Write-Host "✓ ウィンドウが見つかりました！" -ForegroundColor Green
                Write-Host "  タイトル: $($foundWindow.Current.Name)" -ForegroundColor Green
                Write-Host "  ClassName: $($foundWindow.Current.ClassName)" -ForegroundColor Green
                Write-Host "  ProcessId: $($foundWindow.Current.ProcessId)" -ForegroundColor Green

                # プロセス名も取得
                try
                {
                    $process = [System.Diagnostics.Process]::GetProcessById($foundWindow.Current.ProcessId)
                    Write-Host "  ProcessName: $($process.ProcessName)" -ForegroundColor Green
                }
                catch
                {
                    # プロセス情報が取得できない場合は無視
                }

                break
            }
        }
        catch
        {
            Write-Host "✗ パターン '$pattern' では見つかりませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # 見つからなかった場合は、部分一致検索も試す
    if ($null -eq $foundWindow)
    {
        Write-Host "`n部分一致でウィンドウを検索中..." -ForegroundColor Yellow

        try
        {
            $foundWindow = $uiDriver.FindWindowByPartialName("Code")

            if ($null -ne $foundWindow)
            {
                Write-Host "✓ 部分一致でウィンドウが見つかりました！" -ForegroundColor Green
                Write-Host "  タイトル: $($foundWindow.Current.Name)" -ForegroundColor Green
            }
        }
        catch
        {
            Write-Host "✗ 部分一致でも見つかりませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # プロセス名での検索も試す
    if ($null -eq $foundWindow)
    {
        Write-Host "`nプロセス名 'Code' でウィンドウを検索中..." -ForegroundColor Yellow

        try
        {
            $foundWindow = $uiDriver.FindWindowByProcessName("Code")

            if ($null -ne $foundWindow)
            {
                Write-Host "✓ プロセス名でウィンドウが見つかりました！" -ForegroundColor Green
                Write-Host "  タイトル: $($foundWindow.Current.Name)" -ForegroundColor Green
            }
        }
        catch
        {
            Write-Host "✗ プロセス名でも見つかりませんでした: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($null -eq $foundWindow)
    {
        Write-Host "`n`n=== VSCodeウィンドウが見つかりませんでした ===" -ForegroundColor Red
        Write-Host "VSCodeが起動していることを確認してください。" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "`n`n=== テスト成功！ ===" -ForegroundColor Green
    }
}
catch
{
    Write-Host "`nエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally
{
    if ($null -ne $uiDriver)
    {
        $uiDriver.Dispose()
    }

    Write-Host "`n=====================================`n" -ForegroundColor Cyan
}