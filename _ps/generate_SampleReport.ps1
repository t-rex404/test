# 疑似的な報告書を作成するスクリプト
# 必要なアセンブリを読み込み
Add-Type -AssemblyName Microsoft.Office.Interop.Word

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# スクリプトの基準パス
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$LibDir    = Join-Path $ScriptDir '_lib'

# 依存スクリプト読込（Common -> WordDriver）
. (Join-Path $LibDir 'Common.ps1')
. (Join-Path $LibDir 'WordDriver.ps1')

# Commonクラスのインスタンスをグローバル変数として作成
$global:Common = [Common]::new()

# 出力ディレクトリ
$OutDir = Join-Path $RepoRoot '_out'
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

# 画像（存在しない場合は簡易生成）
$ImgPath = Join-Path $OutDir 'sample_word_image.png'
if (-not (Test-Path $ImgPath)) {
	try {
		Add-Type -AssemblyName System.Drawing
		$bmp = New-Object System.Drawing.Bitmap 300, 160
		$gfx = [System.Drawing.Graphics]::FromImage($bmp)
		$brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 30, 144, 255))
		$gfx.FillRectangle($brush, 0, 0, 300, 160)
		$bmp.Save($ImgPath, [System.Drawing.Imaging.ImageFormat]::Png)
		$gfx.Dispose(); $brush.Dispose(); $bmp.Dispose()
	} catch {
		Write-Host "画像生成に失敗しましたが続行します: $($_.Exception.Message)" -ForegroundColor Yellow
	}
}

$documentName = 'sample_report_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.docx'
$reportPath = Join-Path $OutDir $documentName 

$driver = $null
try {
	Write-Host "WordDriverの初期化を開始します..." -ForegroundColor Cyan
	$driver = [WordDriver]::new()
	Write-Host "WordDriverの初期化が完了しました" -ForegroundColor Green

	# 表紙
	Write-Host "表紙を作成中..." -ForegroundColor Cyan
	$driver.SetFont('Calibri', 12)
	$driver.AddHeading('AIレポート（サンプル）', 1)
	$driver.AddParagraph("作成日: $(Get-Date -Format 'yyyy/MM/dd')")
	$driver.AddParagraph('著者: サンプルユーザー')
	$driver.AddPageBreak()

	# 目次
	Write-Host "目次を作成中..." -ForegroundColor Cyan
	$driver.AddHeading('目次', 2)
	$driver.AddTableOfContents()
	$driver.UpdateTableOfContents()
	$driver.AddPageBreak()

	# セクション1: 概要
	Write-Host "セクション1（概要）を作成中..." -ForegroundColor Cyan
	$driver.AddHeading('1. 概要', 2)
	$introText = '本ドキュメントはWordDriverクラスを用いた疑似的な報告書のサンプルです。' +
		' 機能デモとして、見出し・段落・表・画像・ページ区切り・目次等を自動生成します。'
	$driver.AddParagraph($introText)
	for ($i = 1; $i -le 8; $i++) {
		$driver.AddText("ダミーテキスト段落 ${i}: この段落はレイアウト確認用のサンプルテキストです。ページ数を確保するために複数回追加しています。")
	}
	$driver.AddPageBreak()
	
	# セクション2: 基本指標（表）
	Write-Host "セクション2（基本指標）を作成中..." -ForegroundColor Cyan
	$driver.AddHeading('2. 基本指標', 2)
	$rows = 5; $cols = 3
	$data = New-Object 'object[,]' $rows, $cols
	$data[0,0] = '項目'; $data[0,1] = '値'; $data[0,2] = '備考'
	$data[1,0] = '検証ケース数'; $data[1,1] = 128; $data[1,2] = '自動生成'
	$data[2,0] = '成功率(%)'; $data[2,1] = 97.6; $data[2,2] = '目標95%以上'
	$data[3,0] = '平均処理時間(ms)'; $data[3,1] = 245; $data[3,2] = 'ローカル環境'
	$data[4,0] = '最終更新'; $data[4,1] = (Get-Date -Format 'yyyy/MM/dd'); $data[4,2] = '日次更新'
	$driver.AddTable($data, '基本指標（サマリ）')

	# セクション3: 画像
	Write-Host "セクション3（画像）を作成中..." -ForegroundColor Cyan
	$driver.AddHeading('3. 画像', 2)
	if (Test-Path $ImgPath) {
		$driver.AddImage($ImgPath)
	}

	# セクション4: 追加テキストでページ数を稼ぐ
	Write-Host "セクション4（詳細）を作成中..." -ForegroundColor Cyan
	$driver.AddHeading('4. 詳細', 2)
	for ($i = 1; $i -le 12; $i++) {
		$driver.AddParagraph("詳細テキスト ${i}: これは詳細節のサンプル段落です。複数ページにわたることを想定しています。")
	}

	# 目次の最終更新（ページ数確定後）
	Write-Host "目次を更新中..." -ForegroundColor Cyan
	$driver.UpdateTableOfContents()
	#$driver.SetupFooterWithPageNumbers()


	# 保存
	Write-Host "ドキュメントを保存中..." -ForegroundColor Cyan
	$driver.SaveDocument($reportPath)

	Write-Host "疑似報告書を作成しました: $reportPath" -ForegroundColor Green
}
catch {
	Write-Host "エラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
	Write-Host "スタックトレース: $($_.ScriptStackTrace)" -ForegroundColor Red
	
	# Commonオブジェクトが利用可能な場合はエラーログに記録
	if ($global:Common) {
		try {
			$global:Common.HandleError("CommonError_0001", "サンプルレポート生成エラー: $($_.Exception.Message)", "Common", ".\SampleReport_Error.log")
		} catch {
			Write-Host "エラーログの記録に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
		}
	}
	
	throw
}
finally {
	if ($driver) {
		Write-Host "WordDriverのリソースを解放中..." -ForegroundColor Cyan
		$driver.Dispose()
		Write-Host "リソース解放が完了しました" -ForegroundColor Green
	}
}


