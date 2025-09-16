# TNS接続を使用したOracleDriverテストスクリプト
# ========================================
# 前提条件:
# 1. tnsnames.oraファイルが適切に設定されていること
# 2. TNS_ADMIN環境変数が設定されているか、このスクリプトで設定すること
# 3. Oracle Clientがインストールされていること
# ========================================

# ライブラリの読み込み
. ".\\_ps\\_lib\\Common.ps1"
. ".\\_ps\\_lib\\OracleDriver.ps1"

# グローバルCommonオブジェクトの作成
$global:Common = [Common]::new()

# OracleDriverインスタンスの作成
$oracle = [OracleDriver]::new()

# テスト用の設定（環境に合わせて変更してください）
$TNS_ADMIN_PATH = "C:\oracle\network\admin"  # tnsnames.oraがある場所
$TNS_ALIAS = "ORCL"                          # tnsnames.oraで定義されたエイリアス
$USERNAME = "hr"                              # Oracleユーザー名
$PASSWORD = "password"                        # パスワード

Write-Host "========================================"
Write-Host "OracleDriver TNS接続テスト開始"
Write-Host "========================================"

try {
    # TNS_ADMINパスの設定
    Write-Host "`n[1] TNS_ADMINパスの設定" -ForegroundColor Cyan
    $oracle.SetTnsAdminPath($TNS_ADMIN_PATH)

    # ========================================
    # 接続テスト
    # ========================================
    Write-Host "`n[2] TNS接続テスト" -ForegroundColor Cyan

    # 方法1: TNS接続パラメータを設定してから接続
    Write-Host "方法1: SetTnsConnectionParameters → Connect"
    $oracle.SetTnsConnectionParameters($TNS_ALIAS, $USERNAME, $PASSWORD)
    $oracle.Connect()

    # 接続確認
    if ($oracle.IsConnected()) {
        Write-Host "✓ データベース接続成功" -ForegroundColor Green
    } else {
        throw "データベース接続に失敗しました"
    }

    # ========================================
    # DML操作テスト
    # ========================================
    Write-Host "`n[3] DML操作テスト" -ForegroundColor Cyan

    # テストテーブル作成
    Write-Host "テストテーブル作成中..."
    $createTableSql = @"
CREATE TABLE test_tns_table (
    id NUMBER PRIMARY KEY,
    name VARCHAR2(50),
    created_date DATE DEFAULT SYSDATE
)
"@
    $oracle.ExecuteCreateTable($createTableSql)

    # INSERT
    Write-Host "INSERTテスト..."
    $insertSql = "INSERT INTO test_tns_table (id, name) VALUES (:id, :name)"
    $insertParams = @{":id" = 1; ":name" = "Test User 1"}
    $affectedRows = $oracle.ExecuteInsert($insertSql, $insertParams)
    Write-Host "✓ INSERT完了 (影響行数: $affectedRows)" -ForegroundColor Green

    # SELECT
    Write-Host "SELECTテスト..."
    $selectSql = "SELECT * FROM test_tns_table"
    $dataTable = $oracle.ExecuteSelect($selectSql)
    Write-Host "✓ SELECT完了 (取得行数: $($dataTable.Rows.Count))" -ForegroundColor Green

    # UPDATE
    Write-Host "UPDATEテスト..."
    $updateSql = "UPDATE test_tns_table SET name = :name WHERE id = :id"
    $updateParams = @{":name" = "Updated User"; ":id" = 1}
    $affectedRows = $oracle.ExecuteUpdate($updateSql, $updateParams)
    Write-Host "✓ UPDATE完了 (影響行数: $affectedRows)" -ForegroundColor Green

    # MERGE
    Write-Host "MERGEテスト..."
    $mergeSql = @"
MERGE INTO test_tns_table t
USING (SELECT 2 AS id, 'Test User 2' AS name FROM DUAL) s
ON (t.id = s.id)
WHEN MATCHED THEN UPDATE SET t.name = s.name
WHEN NOT MATCHED THEN INSERT (id, name) VALUES (s.id, s.name)
"@
    $affectedRows = $oracle.ExecuteMerge($mergeSql)
    Write-Host "✓ MERGE完了 (影響行数: $affectedRows)" -ForegroundColor Green

    # EXPLAIN PLAN
    Write-Host "EXPLAIN PLANテスト..."
    $planTable = $oracle.ExecuteExplainPlan($selectSql)
    Write-Host "✓ EXPLAIN PLAN完了 (計画行数: $($planTable.Rows.Count))" -ForegroundColor Green

    # ========================================
    # TCL操作テスト
    # ========================================
    Write-Host "`n[4] TCL操作テスト" -ForegroundColor Cyan

    # トランザクション開始
    Write-Host "トランザクション開始..."
    $oracle.BeginTransaction()

    # SAVEPOINTの作成
    Write-Host "SAVEPOINT作成..."
    $oracle.CreateSavepoint("SP1")

    # トランザクション内でのINSERT
    $insertSql2 = "INSERT INTO test_tns_table (id, name) VALUES (3, 'Transaction Test')"
    $command = New-Object System.Data.OracleClient.OracleCommand($insertSql2, $oracle.Connection, $oracle.Transaction)
    $command.ExecuteNonQuery()

    # SAVEPOINTへのロールバック
    Write-Host "SAVEPOINTへロールバック..."
    $oracle.RollbackToSavepoint("SP1")

    # トランザクションのコミット
    Write-Host "トランザクションコミット..."
    $oracle.CommitTransaction()
    Write-Host "✓ TCL操作完了" -ForegroundColor Green

    # ========================================
    # DDL操作テスト
    # ========================================
    Write-Host "`n[5] DDL操作テスト" -ForegroundColor Cyan

    # CREATE INDEX
    Write-Host "CREATE INDEXテスト..."
    $createIndexSql = "CREATE INDEX idx_test_tns_name ON test_tns_table(name)"
    $oracle.ExecuteCreateIndex($createIndexSql)
    Write-Host "✓ CREATE INDEX完了" -ForegroundColor Green

    # CREATE VIEW
    Write-Host "CREATE VIEWテスト..."
    $createViewSql = "CREATE VIEW v_test_tns AS SELECT id, name FROM test_tns_table"
    $oracle.ExecuteCreateView($createViewSql)
    Write-Host "✓ CREATE VIEW完了" -ForegroundColor Green

    # CREATE SEQUENCE
    Write-Host "CREATE SEQUENCEテスト..."
    $createSequenceSql = "CREATE SEQUENCE seq_test_tns START WITH 100"
    $oracle.ExecuteCreateSequence($createSequenceSql)
    Write-Host "✓ CREATE SEQUENCE完了" -ForegroundColor Green

    # TRUNCATE
    Write-Host "TRUNCATEテスト..."
    $oracle.ExecuteTruncate("test_tns_table")
    Write-Host "✓ TRUNCATE完了" -ForegroundColor Green

    # RENAME
    Write-Host "RENAMEテスト..."
    $oracle.ExecuteRename("test_tns_table", "test_tns_table_renamed")
    Write-Host "✓ RENAME完了" -ForegroundColor Green

    # ========================================
    # DCL操作テスト（権限が必要）
    # ========================================
    Write-Host "`n[6] DCL操作テスト" -ForegroundColor Cyan

    try {
        # GRANT（権限が必要）
        Write-Host "GRANTテスト（権限が必要）..."
        $grantSql = "GRANT SELECT ON test_tns_table_renamed TO PUBLIC"
        $oracle.ExecuteGrant($grantSql)
        Write-Host "✓ GRANT完了" -ForegroundColor Green

        # REVOKE
        Write-Host "REVOKEテスト..."
        $revokeSql = "REVOKE SELECT ON test_tns_table_renamed FROM PUBLIC"
        $oracle.ExecuteRevoke($revokeSql)
        Write-Host "✓ REVOKE完了" -ForegroundColor Green
    }
    catch {
        Write-Host "DCL操作には適切な権限が必要です: $_" -ForegroundColor Yellow
    }

    # ========================================
    # SQLPLUS経由でのSQL実行テスト
    # ========================================
    Write-Host "`n[7] SQLPLUS TNS実行テスト" -ForegroundColor Cyan

    # ExecuteSqlPlusTnsを使用
    $sqlplusSql = "SELECT COUNT(*) FROM all_tables WHERE ROWNUM <= 10"
    $result = $oracle.ExecuteSqlPlusTns($sqlplusSql, $USERNAME, $PASSWORD, $TNS_ALIAS)
    Write-Host "✓ SQLPLUS実行完了" -ForegroundColor Green
    Write-Host "結果: $result"

    # ========================================
    # クリーンアップ
    # ========================================
    Write-Host "`n[8] クリーンアップ" -ForegroundColor Cyan

    # テストオブジェクトの削除
    try {
        $oracle.ExecuteDropTable("test_tns_table_renamed")
        Write-Host "テストテーブル削除完了"
    } catch { }

    try {
        $dropViewSql = "DROP VIEW v_test_tns"
        $command = New-Object System.Data.OracleClient.OracleCommand($dropViewSql, $oracle.Connection)
        $command.ExecuteNonQuery()
        Write-Host "テストビュー削除完了"
    } catch { }

    try {
        $dropSequenceSql = "DROP SEQUENCE seq_test_tns"
        $command = New-Object System.Data.OracleClient.OracleCommand($dropSequenceSql, $oracle.Connection)
        $command.ExecuteNonQuery()
        Write-Host "テストシーケンス削除完了"
    } catch { }

    Write-Host "✓ クリーンアップ完了" -ForegroundColor Green

}
catch {
    Write-Host "`n✗ エラーが発生しました: $_" -ForegroundColor Red
    Write-Host "最後のエラーメッセージ: $($oracle.GetLastErrorMessage())" -ForegroundColor Red
}
finally {
    # 接続を切断
    if ($oracle.IsConnected()) {
        $oracle.Disconnect()
        Write-Host "`n接続を切断しました" -ForegroundColor Green
    }

    # リソースの破棄
    $oracle.Dispose()
}

Write-Host "`n========================================"
Write-Host "OracleDriver TNS接続テスト完了"
Write-Host "========================================"

# 方法2: ConnectTnsメソッドを直接使用する例
Write-Host "`n[追加テスト] ConnectTnsメソッドの直接使用" -ForegroundColor Cyan
$oracle2 = [OracleDriver]::new()
try {
    $oracle2.SetTnsAdminPath($TNS_ADMIN_PATH)
    $oracle2.ConnectTns($TNS_ALIAS, $USERNAME, $PASSWORD)

    if ($oracle2.IsConnected()) {
        Write-Host "✓ ConnectTnsメソッドでの接続成功" -ForegroundColor Green

        # 簡単なテスト
        $result = $oracle2.ExecuteSelect("SELECT 'TNS Connection Test' AS message FROM DUAL")
        Write-Host "テスト結果: $($result.Rows[0].message)"
    }
}
catch {
    Write-Host "ConnectTnsメソッドテストエラー: $_" -ForegroundColor Red
}
finally {
    if ($oracle2.IsConnected()) {
        $oracle2.Disconnect()
    }
    $oracle2.Dispose()
}

Write-Host "`n全テスト終了" -ForegroundColor Green