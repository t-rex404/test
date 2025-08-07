# OracleDriverクラスのテストスクリプト
# DML、DDL、DCL、TCLの各操作をテスト

# 共通ライブラリをインポート
. "$PSScriptRoot\_lib\Common.ps1"

# OracleDriverライブラリをインポート
. "$PSScriptRoot\_lib\OracleDriver.ps1"

# テスト用のOracleDriverインスタンスを作成
$oracleDriver = [OracleDriver]::new()

# テスト用の接続パラメータ（実際の環境に合わせて変更してください）
$server = "localhost"
$port = "1521"
$serviceName = "XE"
$username = "test_user"
$password = "test_password"
$schema = "test_user"

Write-Host "=== OracleDriver テスト開始 ===" -ForegroundColor Cyan

try
{
    # 接続パラメータを設定
    Write-Host "`n1. 接続パラメータを設定中..." -ForegroundColor Yellow
    $oracleDriver.SetConnectionParameters($server, $port, $serviceName, $username, $password, $schema)
    
    # データベースに接続
    Write-Host "`n2. データベースに接続中..." -ForegroundColor Yellow
    $oracleDriver.Connect()
    
    # ========================================
    # DDL (Data Definition Language) テスト
    # ========================================
    Write-Host "`n=== DDL テスト ===" -ForegroundColor Green
    
    # テスト用テーブルの作成
    Write-Host "`n3. テスト用テーブルを作成中..." -ForegroundColor Yellow
    $createTableSQL = @"
CREATE TABLE test_employees (
    employee_id NUMBER(6) PRIMARY KEY,
    first_name VARCHAR2(20) NOT NULL,
    last_name VARCHAR2(25) NOT NULL,
    email VARCHAR2(25) UNIQUE NOT NULL,
    hire_date DATE DEFAULT SYSDATE,
    salary NUMBER(8,2),
    department_id NUMBER(4)
)
"@
    $oracleDriver.ExecuteCreateTable($createTableSQL)
    
    # インデックスの作成
    Write-Host "`n4. インデックスを作成中..." -ForegroundColor Yellow
    $createIndexSQL = "CREATE INDEX idx_employees_email ON test_employees(email)"
    $oracleDriver.ExecuteCreateIndex($createIndexSQL)
    
    # シーケンスの作成
    Write-Host "`n5. シーケンスを作成中..." -ForegroundColor Yellow
    $createSequenceSQL = "CREATE SEQUENCE seq_employee_id START WITH 1 INCREMENT BY 1"
    $oracleDriver.ExecuteCreateSequence($createSequenceSQL)
    
    # ビューの作成
    Write-Host "`n6. ビューを作成中..." -ForegroundColor Yellow
    $createViewSQL = "CREATE VIEW v_employees AS SELECT employee_id, first_name, last_name, email FROM test_employees"
    $oracleDriver.ExecuteCreateView($createViewSQL)
    
    # ========================================
    # DML (Data Manipulation Language) テスト
    # ========================================
    Write-Host "`n=== DML テスト ===" -ForegroundColor Green
    
    # トランザクションを開始
    Write-Host "`n7. トランザクションを開始中..." -ForegroundColor Yellow
    $oracleDriver.BeginTransaction()
    
    # INSERT文の実行
    Write-Host "`n8. INSERT文を実行中..." -ForegroundColor Yellow
    $insertSQL = @"
INSERT INTO test_employees (employee_id, first_name, last_name, email, salary, department_id)
VALUES (:employee_id, :first_name, :last_name, :email, :salary, :department_id)
"@
    
    $insertParams = @{
        ":employee_id" = 1
        ":first_name" = "John"
        ":last_name" = "Doe"
        ":email" = "john.doe@example.com"
        ":salary" = 50000
        ":department_id" = 10
    }
    
    $affectedRows = $oracleDriver.ExecuteInsert($insertSQL, $insertParams)
    Write-Host "挿入された行数: $affectedRows" -ForegroundColor Green
    
    # 複数行のINSERT
    $employees = @(
        @{id = 2; first = "Jane"; last = "Smith"; email = "jane.smith@example.com"; salary = 60000; dept = 20},
        @{id = 3; first = "Bob"; last = "Johnson"; email = "bob.johnson@example.com"; salary = 55000; dept = 10},
        @{id = 4; first = "Alice"; last = "Brown"; email = "alice.brown@example.com"; salary = 65000; dept = 30}
    )
    
    foreach ($emp in $employees)
    {
        $params = @{
            ":employee_id" = $emp.id
            ":first_name" = $emp.first
            ":last_name" = $emp.last
            ":email" = $emp.email
            ":salary" = $emp.salary
            ":department_id" = $emp.dept
        }
        $oracleDriver.ExecuteInsert($insertSQL, $params)
    }
    
    # SELECT文の実行
    Write-Host "`n9. SELECT文を実行中..." -ForegroundColor Yellow
    $selectSQL = "SELECT * FROM test_employees ORDER BY employee_id"
    $result = $oracleDriver.ExecuteSelect($selectSQL)
    
    Write-Host "取得された行数: $($result.Rows.Count)" -ForegroundColor Green
    foreach ($row in $result.Rows)
    {
        Write-Host "ID: $($row['EMPLOYEE_ID']), 名前: $($row['FIRST_NAME']) $($row['LAST_NAME']), メール: $($row['EMAIL'])" -ForegroundColor White
    }
    
    # UPDATE文の実行
    Write-Host "`n10. UPDATE文を実行中..." -ForegroundColor Yellow
    $updateSQL = "UPDATE test_employees SET salary = :new_salary WHERE employee_id = :employee_id"
    $updateParams = @{
        ":new_salary" = 52000
        ":employee_id" = 1
    }
    
    $affectedRows = $oracleDriver.ExecuteUpdate($updateSQL, $updateParams)
    Write-Host "更新された行数: $affectedRows" -ForegroundColor Green
    
    # MERGE文の実行（UPSERT）
    Write-Host "`n11. MERGE文を実行中..." -ForegroundColor Yellow
    $mergeSQL = @"
MERGE INTO test_employees t
USING (SELECT :employee_id as emp_id, :first_name as fname, :last_name as lname, :email as mail, :salary as sal, :department_id as dept FROM dual) s
ON (t.employee_id = s.emp_id)
WHEN MATCHED THEN
    UPDATE SET first_name = s.fname, last_name = s.lname, email = s.mail, salary = s.sal, department_id = s.dept
WHEN NOT MATCHED THEN
    INSERT (employee_id, first_name, last_name, email, salary, department_id)
    VALUES (s.emp_id, s.fname, s.lname, s.mail, s.sal, s.dept)
"@
    
    $mergeParams = @{
        ":employee_id" = 5
        ":first_name" = "Charlie"
        ":last_name" = "Wilson"
        ":email" = "charlie.wilson@example.com"
        ":salary" = 70000
        ":department_id" = 40
    }
    
    $affectedRows = $oracleDriver.ExecuteMerge($mergeSQL, $mergeParams)
    Write-Host "MERGEで影響を受けた行数: $affectedRows" -ForegroundColor Green
    
    # SAVEPOINTのテスト
    Write-Host "`n12. SAVEPOINTを作成中..." -ForegroundColor Yellow
    $oracleDriver.CreateSavepoint("test_savepoint")
    
    # DELETE文の実行
    Write-Host "`n13. DELETE文を実行中..." -ForegroundColor Yellow
    $deleteSQL = "DELETE FROM test_employees WHERE employee_id = :employee_id"
    $deleteParams = @{":employee_id" = 5}
    
    $affectedRows = $oracleDriver.ExecuteDelete($deleteSQL, $deleteParams)
    Write-Host "削除された行数: $affectedRows" -ForegroundColor Green
    
    # SAVEPOINTにロールバック
    Write-Host "`n14. SAVEPOINTにロールバック中..." -ForegroundColor Yellow
    $oracleDriver.RollbackToSavepoint("test_savepoint")
    
    # 最終的なSELECT
    Write-Host "`n15. 最終的なデータを確認中..." -ForegroundColor Yellow
    $finalResult = $oracleDriver.ExecuteSelect($selectSQL)
    Write-Host "最終的な行数: $($finalResult.Rows.Count)" -ForegroundColor Green
    
    # トランザクションをコミット
    Write-Host "`n16. トランザクションをコミット中..." -ForegroundColor Yellow
    $oracleDriver.CommitTransaction()
    
    # ========================================
    # DCL (Data Control Language) テスト
    # ========================================
    Write-Host "`n=== DCL テスト ===" -ForegroundColor Green
    
    # GRANT文の実行（権限がある場合のみ）
    Write-Host "`n17. GRANT文を実行中..." -ForegroundColor Yellow
    try
    {
        $grantSQL = "GRANT SELECT ON test_employees TO public"
        $oracleDriver.ExecuteGrant($grantSQL)
        Write-Host "GRANT文が正常に実行されました。" -ForegroundColor Green
    }
    catch
    {
        Write-Host "GRANT文の実行に失敗しました（権限不足の可能性）: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # ========================================
    # ユーティリティメソッドのテスト
    # ========================================
    Write-Host "`n=== ユーティリティメソッド テスト ===" -ForegroundColor Green
    
    # テーブル一覧の取得
    Write-Host "`n18. テーブル一覧を取得中..." -ForegroundColor Yellow
    $tableList = $oracleDriver.GetTableList($schema)
    Write-Host "テーブル数: $($tableList.Rows.Count)" -ForegroundColor Green
    
    # テーブル構造の取得
    Write-Host "`n19. テーブル構造を取得中..." -ForegroundColor Yellow
    $tableStructure = $oracleDriver.GetTableStructure("test_employees", $schema)
    Write-Host "カラム数: $($tableStructure.Rows.Count)" -ForegroundColor Green
    
    foreach ($column in $tableStructure.Rows)
    {
        Write-Host "カラム: $($column['COLUMN_NAME']), 型: $($column['DATA_TYPE']), 長さ: $($column['DATA_LENGTH']), NULL許可: $($column['NULLABLE'])" -ForegroundColor White
    }
    
    # 接続状態の確認
    Write-Host "`n20. 接続状態を確認中..." -ForegroundColor Yellow
    $isConnected = $oracleDriver.IsConnected()
    Write-Host "接続状態: $isConnected" -ForegroundColor Green
    
    # トランザクション状態の確認
    $isTransactionActive = $oracleDriver.IsTransactionActive()
    Write-Host "トランザクション状態: $isTransactionActive" -ForegroundColor Green
    
    # ========================================
    # クリーンアップ
    # ========================================
    Write-Host "`n=== クリーンアップ ===" -ForegroundColor Green
    
    # テスト用オブジェクトの削除
    Write-Host "`n21. テスト用オブジェクトを削除中..." -ForegroundColor Yellow
    
    # ビューの削除
    try
    {
        $oracleDriver.ExecuteAlterTable("DROP VIEW v_employees")
        Write-Host "ビューを削除しました。" -ForegroundColor Green
    }
    catch
    {
        Write-Host "ビューの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # シーケンスの削除
    try
    {
        $oracleDriver.ExecuteAlterTable("DROP SEQUENCE seq_employee_id")
        Write-Host "シーケンスを削除しました。" -ForegroundColor Green
    }
    catch
    {
        Write-Host "シーケンスの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # テーブルの削除
    try
    {
        $oracleDriver.ExecuteDropTable("test_employees")
        Write-Host "テーブルを削除しました。" -ForegroundColor Green
    }
    catch
    {
        Write-Host "テーブルの削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host "`n=== テスト完了 ===" -ForegroundColor Cyan
}
catch
{
    Write-Host "`nテスト中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    
    # トランザクションがアクティブな場合はロールバック
    if ($oracleDriver.IsTransactionActive())
    {
        Write-Host "トランザクションをロールバック中..." -ForegroundColor Yellow
        $oracleDriver.RollbackTransaction()
    }
}
finally
{
    # 接続を切断
    if ($oracleDriver.IsConnected())
    {
        Write-Host "`nデータベース接続を切断中..." -ForegroundColor Yellow
        $oracleDriver.Disconnect()
    }
    
    # リソースを破棄
    $oracleDriver.Dispose()
    
    Write-Host "`nOracleDriverテストが終了しました。" -ForegroundColor Cyan
} 