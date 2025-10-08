# スクリプトの基準パス
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$LibDir    = Join-Path $ScriptDir '_lib'

# 必要なモジュールを読み込み
. (Join-Path $LibDir 'Common.ps1')
$global:Common.WriteLog("テスト", "INFO")

$global:Common.Dispose()
exit 0