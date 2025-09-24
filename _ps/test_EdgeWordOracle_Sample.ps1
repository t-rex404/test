# EdgeDriver、WordDriver、OracleDriverを使用したサンプル
# 3つのドライバーを組み合わせた実用的な例

# アセンブリと依存関係の読み込み
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -Path ".\_lib\Oracle.ManagedDataAccess.dll"

# 共通ライブラリとドライバーの読み込み
. ".\_lib\Common.ps1"
. ".\_lib\EdgeDriver.ps1"
. ".\_lib\WordDriver.ps1"
. ".\_lib\OracleDriver.ps1"

# グローバル変数の初期化
$global:Common = [Common]::new()

function Test-EdgeWordOracleIntegration {
    Write-Host "`n=== EdgeDriver、WordDriver、OracleDriverの統合サンプル ===" -ForegroundColor Cyan

    try {
        # 1. EdgeDriverでWebサイトから情報を取得
        Write-Host "`n[1] EdgeDriverでWebサイトから情報を取得" -ForegroundColor Yellow
        $edge = [EdgeDriver]::new()
        $edge.StartBrowser()
        Start-Sleep -Seconds 2

        # サンプルサイトにアクセス
        $edge.NavigateToUrl("https://www.example.com")
        Start-Sleep -Seconds 2

        # ページタイトルを取得
        $pageTitle = $edge.GetTitle()
        Write-Host "  取得したタイトル: $pageTitle" -ForegroundColor Green

        # ページのテキストコンテンツを取得
        $pageContent = $edge.ExecuteScript("return document.body.innerText;")
        Write-Host "  ページコンテンツ取得完了" -ForegroundColor Green

        # 2. WordDriverで取得した情報をWord文書に記録
        Write-Host "`n[2] WordDriverで取得情報をWord文書に記録" -ForegroundColor Yellow
        $word = [WordDriver]::new()
        $word.StartWord($true)

        # 新規文書作成
        $doc = $word.CreateNewDocument()

        # タイトルを追加
        $word.AddText("Webサイト情報レポート", $true, $true, 16)
        $word.AddText("`n", $false, $false, 12)

        # 取得日時を追加
        $currentDateTime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
        $word.AddText("取得日時: $currentDateTime", $false, $false, 10)
        $word.AddText("`n`n", $false, $false, 12)

        # ページ情報を追加
        $word.AddText("ページタイトル: ", $true, $false, 12)
        $word.AddText($pageTitle, $false, $false, 12)
        $word.AddText("`n`n", $false, $false, 12)

        $word.AddText("ページコンテンツ:", $true, $false, 12)
        $word.AddText("`n", $false, $false, 12)
        $word.AddText($pageContent.Substring(0, [Math]::Min($pageContent.Length, 500)), $false, $false, 10)

        # Word文書を保存
        $reportPath = "$PWD\WebReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').docx"
        $word.SaveAs($reportPath)
        Write-Host "  Word文書保存完了: $reportPath" -ForegroundColor Green

        # 3. OracleDriverでデータベースに記録（接続文字列は仮想）
        Write-Host "`n[3] OracleDriverでデータベースに記録（シミュレーション）" -ForegroundColor Yellow

        # 実際の環境では以下のような接続文字列を使用
        # $connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ORCL)));User Id=testuser;Password=testpass;"

        # サンプルのため、データベース操作をシミュレート
        Write-Host "  データベースへの接続をシミュレート..." -ForegroundColor Gray

        # 仮想的なSQL文
        $insertSQL = @"
INSERT INTO WEB_REPORTS (
    REPORT_ID,
    PAGE_TITLE,
    CONTENT_PREVIEW,
    CAPTURE_DATE,
    WORD_DOCUMENT_PATH
) VALUES (
    SEQ_REPORT_ID.NEXTVAL,
    '$($pageTitle -replace "'", "''")',
    '$($pageContent.Substring(0, [Math]::Min($pageContent.Length, 100)) -replace "'", "''")',
    SYSDATE,
    '$($reportPath -replace "'", "''")'
)
"@

        Write-Host "  以下のSQLを実行（シミュレーション）:" -ForegroundColor Gray
        Write-Host $insertSQL -ForegroundColor DarkGray
        Write-Host "  データベース記録完了（シミュレーション）" -ForegroundColor Green

        # 実際のOracle接続例（コメントアウト）
        <#
        $oracle = [OracleDriver]::new()
        $oracle.Connect($connectionString)

        # データ挿入
        $oracle.ExecuteNonQuery($insertSQL)

        # 挿入結果を確認
        $selectSQL = "SELECT * FROM WEB_REPORTS WHERE ROWNUM <= 5 ORDER BY CAPTURE_DATE DESC"
        $result = $oracle.ExecuteQuery($selectSQL)

        Write-Host "`n最近のレポート記録:" -ForegroundColor Yellow
        foreach ($row in $result) {
            Write-Host "  ID: $($row["REPORT_ID"]), Title: $($row["PAGE_TITLE"]), Date: $($row["CAPTURE_DATE"])"
        }

        $oracle.Disconnect()
        #>

        # 4. 統合結果のサマリー
        Write-Host "`n[4] 統合処理のサマリー" -ForegroundColor Cyan
        Write-Host "  ✓ Webサイトから情報取得完了 (EdgeDriver)" -ForegroundColor Green
        Write-Host "  ✓ Word文書作成・保存完了 (WordDriver)" -ForegroundColor Green
        Write-Host "  ✓ データベース記録完了 (OracleDriver - シミュレーション)" -ForegroundColor Green
        Write-Host "`n統合サンプル実行成功！" -ForegroundColor Green

        # クリーンアップ
        $edge.CloseBrowser()
        $word.CloseWord()

        return $true

    } catch {
        Write-Host "エラー発生: $_" -ForegroundColor Red
        $global:Common.HandleError("SAMPLE_ERROR", $_.ToString(), "EdgeWordOracleSample", "sample_error.log")
        return $false
    }
}

function Test-DataCollectionWorkflow {
    Write-Host "`n=== データ収集ワークフローのサンプル ===" -ForegroundColor Cyan

    try {
        # 複数のWebサイトから情報を収集してWord文書にまとめる例
        Write-Host "`n複数サイトから情報収集..." -ForegroundColor Yellow

        $edge = [EdgeDriver]::new()
        $word = [WordDriver]::new()

        $edge.StartBrowser()
        $word.StartWord($true)
        $doc = $word.CreateNewDocument()

        # レポートタイトル
        $word.AddText("マルチサイト情報収集レポート", $true, $true, 18)
        $word.AddText("`n`n", $false, $false, 12)

        # 収集するサイトリスト
        $sites = @(
            @{Url="https://www.google.com"; Name="Google"},
            @{Url="https://www.microsoft.com"; Name="Microsoft"},
            @{Url="https://www.github.com"; Name="GitHub"}
        )

        foreach ($site in $sites) {
            Write-Host "  $($site.Name) から情報取得中..." -ForegroundColor Gray

            # サイトにアクセス
            $edge.NavigateToUrl($site.Url)
            Start-Sleep -Seconds 2

            # 情報取得
            $title = $edge.GetTitle()
            $url = $edge.GetCurrentUrl()

            # Word文書に追加
            $word.AddText("■ $($site.Name)", $true, $false, 14)
            $word.AddText("`n", $false, $false, 12)
            $word.AddText("URL: $url", $false, $false, 10)
            $word.AddText("`n", $false, $false, 10)
            $word.AddText("タイトル: $title", $false, $false, 10)
            $word.AddText("`n`n", $false, $false, 12)

            Write-Host "    ✓ 完了" -ForegroundColor Green
        }

        # レポート保存
        $reportPath = "$PWD\MultiSiteReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').docx"
        $word.SaveAs($reportPath)
        Write-Host "`nレポート保存完了: $reportPath" -ForegroundColor Green

        # クリーンアップ
        $edge.CloseBrowser()
        $word.CloseWord()

        return $true

    } catch {
        Write-Host "エラー発生: $_" -ForegroundColor Red
        return $false
    }
}

# メイン実行
Write-Host "EdgeDriver、WordDriver、OracleDriver統合サンプル" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan

# サンプル1: 基本的な統合
$result1 = Test-EdgeWordOracleIntegration

# サンプル2: データ収集ワークフロー
$result2 = Test-DataCollectionWorkflow

# 結果サマリー
Write-Host "`n" + "=" * 60 -ForegroundColor Cyan
Write-Host "実行結果サマリー" -ForegroundColor Cyan
Write-Host "  統合サンプル: $(if($result1){'成功'}else{'失敗'})" -ForegroundColor $(if($result1){'Green'}else{'Red'})
Write-Host "  データ収集ワークフロー: $(if($result2){'成功'}else{'失敗'})" -ForegroundColor $(if($result2){'Green'}else{'Red'})
Write-Host "=" * 60 -ForegroundColor Cyan