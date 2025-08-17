# PowerPointDriver 全メソッド機能テスト
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

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

# 依存スクリプト読込（Common -> PowerPointDriver）
. (Join-Path $LibDir 'Common.ps1')
. (Join-Path $LibDir 'PowerPointDriver.ps1')

# 出力ディレクトリ
$OutDir = Join-Path $RepoRoot '_out'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

# 画像生成（PowerPoint用）
$ImgPath = Join-Path $OutDir 'sample_ppt_image.png'
try {
	Add-Type -AssemblyName System.Drawing
	$bmp = New-Object System.Drawing.Bitmap 300, 150
	$gfx = [System.Drawing.Graphics]::FromImage($bmp)
	$brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 46, 204, 113))
	$gfx.FillRectangle($brush, 0, 0, 300, 150)
	$bmp.Save($ImgPath, [System.Drawing.Imaging.ImageFormat]::Png)
	$gfx.Dispose(); $brush.Dispose(); $bmp.Dispose()
} catch {
	Write-Host "画像生成に失敗しましたがテストを継続します: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ドライバー生成
$driver = $null
try {
	$driver = [PowerPointDriver]::new()
} catch {
	Write-Error "PowerPointDriver 初期化に失敗しました: $($_.Exception.Message)"
	throw
}

try {
	Write-Host "=== PowerPointDriver 全メソッド機能テスト開始 ===" -ForegroundColor Cyan

	# 1. スライド操作
	Write-Host "`n--- 1. スライド操作 ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'AddSlide(1)' -ScriptBlock { $driver.AddSlide(1) }
	Invoke-TestStep -Name 'SelectSlide(1)' -ScriptBlock { $driver.SelectSlide(1) }

	# 2. テキスト操作
	Write-Host "`n--- 2. テキスト操作 ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'SetTitle("Presentation Title")' -ScriptBlock { $driver.SetTitle('Presentation Title') }
	Invoke-TestStep -Name 'AddTextBox("Hello PPT")' -ScriptBlock { $driver.AddTextBox('Hello PPT', 120, 160, 300, 80) }
	Invoke-TestStep -Name 'AddText("Additional Text")' -ScriptBlock { $driver.AddText('Additional Text', 150, 260) }

	# 3. 図形・画像
	Write-Host "`n--- 3. 図形・画像 ---" -ForegroundColor Yellow
	$msShapeRectangle = 1
	Invoke-TestStep -Name 'AddShape(Rectangle)' -ScriptBlock { $driver.AddShape($msShapeRectangle, 400, 120, 120, 80) }
	if (Test-Path $ImgPath) {
		Invoke-TestStep -Name 'AddPicture(sample_ppt_image.png)' -ScriptBlock { $driver.AddPicture($ImgPath, 50, 50, 200, 120) }
	}

	# 4. フォーマット
	Write-Host "`n--- 4. フォーマット ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'SetFont(Meiryo,18)' -ScriptBlock { $driver.SetFont('Meiryo', 18) }
	Invoke-TestStep -Name 'SetBackgroundColor(0xFFAA33)' -ScriptBlock { $driver.SetBackgroundColor(0x00FFAA33) }

	# 5. ファイル操作
	Write-Host "`n--- 5. ファイル操作 ---" -ForegroundColor Yellow
	$pptxPath = Join-Path $OutDir 'ppt_allmethods_test.pptx'
	Invoke-TestStep -Name "SavePresentation($pptxPath)" -ScriptBlock { $driver.SavePresentation($pptxPath) }
	Invoke-TestStep -Name "OpenPresentation($pptxPath)" -ScriptBlock { $driver.OpenPresentation($pptxPath) }

	Write-Host "`n=== PowerPointDriver 全メソッド機能テスト完了 ===" -ForegroundColor Cyan
}
finally {
	if ($driver) {
		Invoke-TestStep -Name 'Dispose()' -ScriptBlock { $driver.Dispose() }
	}
}

Write-Host "テスト完了: PowerPointDriver の主要メソッドを実行しました。" -ForegroundColor Green


