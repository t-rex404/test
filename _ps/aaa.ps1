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
. (Join-Path $libPath "WinSCPDriver.ps1")
