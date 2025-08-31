# PowerShell Driver Classes - Ollama 起動スクリプト
# 実行ポリシーを変更する必要がある場合は以下を実行してください：
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

param(
    [switch]$Silent
)

# エンコーディングをUTF-8に設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# タイトルを設定
$Host.UI.RawUI.WindowTitle = "PowerShell Driver Classes - Ollama 起動スクリプト"

# カラー出力関数
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Write-Success { Write-ColorOutput $args[0] "Green" }
function Write-Error { Write-ColorOutput $args[0] "Red" }
function Write-Warning { Write-ColorOutput $args[0] "Yellow" }
function Write-Info { Write-ColorOutput $args[0] "Cyan" }

# ヘッダー表示
if (-not $Silent) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "PowerShell Driver Classes - Ollama 起動" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
}

# 現在のディレクトリを取得
$CurrentDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OllamaDir = Join-Path $CurrentDir "ollama"
$ModelsDir = Join-Path $OllamaDir "models"

# Ollamaディレクトリが存在するかチェック
if (-not (Test-Path $OllamaDir)) {
    Write-Error "❌ Ollamaディレクトリが見つかりません"
    Write-Info "ディレクトリ: $OllamaDir"
    Write-Host ""
    Write-Host "このドキュメントフォルダに ollama フォルダが含まれているか確認してください。" -ForegroundColor Yellow
    Write-Host ""
    if (-not $Silent) { Read-Host "Enterキーを押して終了" }
    exit 1
}

# Ollama実行ファイルが存在するかチェック
$OllamaExe = Join-Path $OllamaDir "ollama.exe"
if (-not (Test-Path $OllamaExe)) {
    Write-Error "❌ Ollama実行ファイルが見つかりません"
    Write-Info "ファイル: $OllamaExe"
    Write-Host ""
    Write-Host "ollama フォルダに ollama.exe が含まれているか確認してください。" -ForegroundColor Yellow
    Write-Host ""
    if (-not $Silent) { Read-Host "Enterキーを押して終了" }
    exit 1
}

# 環境変数を設定
$env:PATH = "$OllamaDir;$env:PATH"
$env:OLLAMA_HOME = $OllamaDir
$env:OLLAMA_MODELS = $ModelsDir

Write-Success "✅ Ollama環境を設定しました"
Write-Info "実行ファイル: $OllamaExe"
Write-Info "モデルディレクトリ: $ModelsDir"
Write-Host ""

# 既存のOllamaプロセスを終了
Write-Info "🔄 既存のOllamaプロセスを確認中..."
try {
    Get-Process -Name "ollama" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
} catch {
    # プロセスが存在しない場合は無視
}

# Ollamaを起動
Write-Info "🚀 Ollamaを起動中..."
try {
    Start-Process -FilePath $OllamaExe -ArgumentList "serve" -WindowStyle Minimized
} catch {
    Write-Error "❌ Ollamaの起動に失敗しました"
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
    if (-not $Silent) { Read-Host "Enterキーを押して終了" }
    exit 1
}

# 起動待機
Write-Info "⏳ Ollamaの起動を待機中..."
Start-Sleep -Seconds 5

# 接続テスト
Write-Info "🔍 接続テスト中..."
try {
    $response = Invoke-RestMethod -Uri "http://localhost:11434/api/tags" -Method Get -TimeoutSec 10
    Write-Success "✅ Ollamaが正常に起動しました！"
    Write-Host ""
    Write-Info "🌐 ローカルエンドポイント: http://localhost:11434"
    Write-Info "📚 利用可能なモデルを確認中..."
    Write-Host ""
    
    # 利用可能なモデルを表示
    try {
        & $OllamaExe list
    } catch {
        Write-Warning "⚠️ モデル一覧の取得に失敗しました"
    }
    
    Write-Host ""
    Write-Success "🎉 チャットボットが使用可能です！"
    Write-Host "index.html を開いてチャットボットをお試しください。" -ForegroundColor Green
    Write-Host ""
    Write-Host "💡 チャットボットを閉じる際は、このウィンドウを閉じてください。" -ForegroundColor Yellow
    Write-Host ""
    
} catch {
    Write-Error "❌ Ollamaの起動に失敗しました"
    Write-Host ""
    Write-Host "トラブルシューティング:" -ForegroundColor Yellow
    Write-Host "1. ファイアウォールの設定を確認" -ForegroundColor White
    Write-Host "2. ポート11434が他のアプリで使用されていないか確認" -ForegroundColor White
    Write-Host "3. 管理者権限で実行してみてください" -ForegroundColor White
    Write-Host ""
}

Write-Host ""
Write-Host "このウィンドウは開いたままにしてください。" -ForegroundColor Cyan
Write-Host "Ollamaサービスが動作中です。" -ForegroundColor Cyan
Write-Host ""

if (-not $Silent) {
    Read-Host "Enterキーを押して終了"
}
