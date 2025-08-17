# ExcelDriver 全メソッド機能テスト
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

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

# 依存スクリプト読込（Common -> ExcelDriver）
. (Join-Path $LibDir 'Common.ps1')
. (Join-Path $LibDir 'ExcelDriver.ps1')

# 出力ディレクトリ
$OutDir = Join-Path $RepoRoot '_out'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

# ドライバー生成
$driver = $null
try {
	$driver = [ExcelDriver]::new()
} catch {
	Write-Error "ExcelDriver 初期化に失敗しました: $($_.Exception.Message)"
	throw
}

try {
	Write-Host "=== ExcelDriver 全メソッド機能テスト開始 ===" -ForegroundColor Cyan

	# 1. セル操作
	Write-Host "`n--- 1. セル操作 ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'SetCellValue(A1,Hello)' -ScriptBlock { $driver.SetCellValue('A1','Hello') }
	Invoke-TestStep -Name 'GetCellValue(A1)' -ScriptBlock { $null = $driver.GetCellValue('A1') }

	# 範囲操作 (2x2)
	$values = New-Object 'object[,]' 2,2
	$values[0,0] = 1; $values[0,1] = 2
	$values[1,0] = 3; $values[1,1] = 4
	Invoke-TestStep -Name 'SetRangeValue(A2:B3,[2x2])' -ScriptBlock { $driver.SetRangeValue('A2:B3', $values) }
	Invoke-TestStep -Name 'GetRangeValue(A2:B3)' -ScriptBlock { $null = $driver.GetRangeValue('A2:B3') }

	# 2. フォーマット
	Write-Host "`n--- 2. フォーマット ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'SetCellFont(A1,Meiryo,12)' -ScriptBlock { $driver.SetCellFont('A1','Meiryo',12) }
	Invoke-TestStep -Name 'SetCellBold(A1,true)' -ScriptBlock { $driver.SetCellBold('A1',$true) }
	Invoke-TestStep -Name 'SetCellBackgroundColor(A1,6)' -ScriptBlock { $driver.SetCellBackgroundColor('A1', 6) }

	# 3. ワークシート操作
	Write-Host "`n--- 3. ワークシート操作 ---" -ForegroundColor Yellow
	Invoke-TestStep -Name 'AddWorksheet(Sheet2)' -ScriptBlock { $driver.AddWorksheet('Sheet2') }
	Invoke-TestStep -Name 'SelectWorksheet(Sheet2)' -ScriptBlock { $driver.SelectWorksheet('Sheet2') }
	Invoke-TestStep -Name 'SetCellValue(B2,Sheet2Val)' -ScriptBlock { $driver.SetCellValue('B2','Sheet2Val') }
	Invoke-TestStep -Name 'SelectWorksheet(Sheet1)' -ScriptBlock { $driver.SelectWorksheet('Sheet1') }

	# 4. ファイル操作
	Write-Host "`n--- 4. ファイル操作 ---" -ForegroundColor Yellow
	$xlsxPath = Join-Path $OutDir 'excel_allmethods_test.xlsx'
	Invoke-TestStep -Name "SaveWorkbook($xlsxPath)" -ScriptBlock { $driver.SaveWorkbook($xlsxPath) }
	Invoke-TestStep -Name "OpenWorkbook($xlsxPath)" -ScriptBlock { $driver.OpenWorkbook($xlsxPath) }

	Write-Host "`n=== ExcelDriver 全メソッド機能テスト完了 ===" -ForegroundColor Cyan
}
finally {
	if ($driver) {
		Invoke-TestStep -Name 'Dispose()' -ScriptBlock { $driver.Dispose() }
	}
}

Write-Host "テスト完了: ExcelDriver の主要メソッドを実行しました。" -ForegroundColor Green


