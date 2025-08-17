# WordDriver 全メソッド機能テスト
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Word

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-StepResult {
	param(
		[string]$Name,
		[bool]$Succeeded,
		[string]$Message
	)
	$status = if ($Succeeded) { 'OK' } else { 'NG' }
	Write-Host ("[STEP] {0,-45} : {1} - {2}" -f $Name, $status, $Message)
}

function Invoke-TestStep {
	param(
		[string]$Name,
		[scriptblock]$ScriptBlock
	)
	try {
		$result = & $ScriptBlock
		Write-StepResult -Name $Name -Succeeded $true -Message "Success"
		return $result
	} catch {
		Write-StepResult -Name $Name -Succeeded $false -Message $_.Exception.Message
	}
}

# スクリプトの基準パス
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$LibDir    = Join-Path $ScriptDir '_lib'

# 依存スクリプト読込（Common -> WordDriver）
. (Join-Path $LibDir 'Common.ps1')
. (Join-Path $LibDir 'WordDriver.ps1')

# 出力ディレクトリ
$OutDir = Join-Path $RepoRoot '_out'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

# 画像生成（Word用）
$ImgPath = Join-Path $OutDir 'sample_word_image.png'
try {
	Add-Type -AssemblyName System.Drawing
	$bmp = New-Object System.Drawing.Bitmap 200, 100
	$gfx = [System.Drawing.Graphics]::FromImage($bmp)
	$brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 30, 144, 255))
	$gfx.FillRectangle($brush, 0, 0, 200, 100)
	$bmp.Save($ImgPath, [System.Drawing.Imaging.ImageFormat]::Png)
	$gfx.Dispose(); $brush.Dispose(); $bmp.Dispose()
} catch {
	Write-Host "画像生成に失敗しましたがテストを継続します: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ドライバー生成
$driver = $null
try {
	$driver = [WordDriver]::new()
} catch {
	Write-Error "WordDriver 初期化に失敗しました: $($_.Exception.Message)"
	throw
}

try {
	Write-Host "=== WordDriver 全メソッド機能テスト開始 ===" -ForegroundColor Cyan

	# 1. コンテンツ追加関連
	Write-Host "`n--- 1. コンテンツ追加関連 ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'AddHeading("Document Title",1)' -ScriptBlock { $driver.AddHeading('Document Title', 1) }
	Invoke-TestStep -Name 'AddParagraph("Intro paragraph")' -ScriptBlock { $driver.AddParagraph('Intro paragraph') }
	Invoke-TestStep -Name 'AddText("Additional text")' -ScriptBlock { $driver.AddText('Additional text') }
	Invoke-TestStep -Name 'AddHeading("Section 1",2)' -ScriptBlock { $driver.AddHeading('Section 1', 2) }

	# 表データ作成 (2x2)
	$data = New-Object 'object[,]' 2,2
	$data[0,0] = 'R1C1'; $data[0,1] = 'R1C2'
	$data[1,0] = 'R2C1'; $data[1,1] = 'R2C2'
	Invoke-TestStep -Name 'AddTable(2x2,"Sample Table")' -ScriptBlock { $driver.AddTable($data, 'Sample Table') }

	# 画像
	if (Test-Path $ImgPath) {
		Invoke-TestStep -Name 'AddImage(sample_word_image.png)' -ScriptBlock { $driver.AddImage($ImgPath) }
	}

	Invoke-TestStep -Name 'AddPageBreak()' -ScriptBlock { $driver.AddPageBreak() }
	Invoke-TestStep -Name 'AddHeading("Section 2",2)' -ScriptBlock { $driver.AddHeading('Section 2', 2) }

	# 2. 目次・スタイル関連
	Write-Host "`n--- 2. 目次・スタイル関連 ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'AddTableOfContents()' -ScriptBlock { $driver.AddTableOfContents() }
	Invoke-TestStep -Name 'UpdateTableOfContents()' -ScriptBlock { $driver.UpdateTableOfContents() }
	Invoke-TestStep -Name 'SetFont(Calibri,12)' -ScriptBlock { $driver.SetFont('Calibri', 12) }
	Invoke-TestStep -Name 'SetPageOrientation(Landscape)' -ScriptBlock { $driver.SetPageOrientation('Landscape') }

	# 3. ファイル操作関連
	Write-Host "`n--- 3. ファイル操作関連 ---" -ForegroundColor Yellow
	$docPath = Join-Path $OutDir 'word_allmethods_test.docx'
	Invoke-TestStep -Name "SaveDocument($docPath)" -ScriptBlock { $driver.SaveDocument($docPath) }
	Invoke-TestStep -Name "OpenDocument($docPath)" -ScriptBlock { $driver.OpenDocument($docPath) }

	Write-Host "`n=== WordDriver 全メソッド機能テスト完了 ===" -ForegroundColor Cyan
}
finally {
	if ($driver) {
		Invoke-TestStep -Name 'Dispose()' -ScriptBlock { $driver.Dispose() }
	}
}

Write-Host "テスト完了: WordDriver の主要メソッドを実行しました。" -ForegroundColor Green


