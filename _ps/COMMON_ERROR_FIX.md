# Common.ps1 変数初期化エラー修正ガイド

## 問題の概要

`run.cmd`を実行した際に、`Common.ps1`ファイルで以下のエラーが発生する問題を修正します：

```
At C:\000_common\PowerShell\prototype\test\_ps\_lib\Common.ps1:58 char:22
+ - PowerShellバージョン: $($PSVersionTable.PSVersion)
+                      ~~~~~~~~~~~~~~~
Variable is not assigned in the method.
At C:\000_common\PowerShell\prototype\test\_ps\_lib\Common.ps1:59 char:9
+ - OS: $($PSVersionTable.OS)
+         ~~~~~~~~~~~~~~~
Variable is not assigned in the method.
```

## エラーの原因

PowerShellクラス内でヒアドキュメント（`@"..."@`）を使用する際に、変数がクラスメソッド内で正しくスコープされていないことが原因です。

具体的には：
1. `HandleError`メソッド内でヒアドキュメントを使用している
2. ヒアドキュメント内で`$PSVersionTable`や`$env:USERNAME`などの変数を参照している
3. PowerShellクラス内では、これらの変数がメソッドスコープで認識されない

## 修正手順

### 1. 統合修正スクリプトの実行（推奨）

```powershell
# _psディレクトリに移動
cd _ps

# 統合修正スクリプトを実行（すべてのエラーを修正）
.\fix_all_errors.ps1
```

### 2. Common.ps1専用修正スクリプトの実行

```powershell
# _psディレクトリに移動
cd _ps

# Common.ps1専用修正スクリプトを実行
.\fix_common_error.ps1
```

### 3. 手動修正（自動修正が失敗した場合）

#### Common.ps1 の修正

```powershell
# HandleErrorメソッド内の修正
[void] HandleError([string]$errorCode, [string]$message, [string]$module = "Common")
{
    $timestamp = $(Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
    $errorMessage = "[$timestamp], ERROR_CODE:$errorCode, MODULE:$module, ERROR_MESSAGE:$message"
    
    # ログファイルに書き込み
    $logFile = ".\Common_Error.log"
    $errorMessage | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
    
    # コンソールにエラーを表示
    Write-Error $errorMessage
    
    # 詳細なエラー情報をログに追加（デバッグ用）
    # ヒアドキュメントの代わりに文字列連結を使用
    $psVersion = $PSVersionTable.PSVersion.ToString()
    $osInfo = $PSVersionTable.OS
    $userName = $env:USERNAME
    $currentPath = $PWD.Path
    
    $debugInfo = "詳細情報:`n" +
                "- エラーコード: $errorCode`n" +
                "- エラーメッセージ: $message`n" +
                "- モジュール: $module`n" +
                "- タイムスタンプ: $timestamp`n" +
                "- PowerShellバージョン: $psVersion`n" +
                "- OS: $osInfo`n" +
                "- 実行ユーザー: $userName`n" +
                "- 実行パス: $currentPath"
    
    $debugInfo | Out-File -Append -FilePath $logFile -Encoding UTF8 -ErrorAction SilentlyContinue
}
```

## 修正内容の詳細

### 1. ヒアドキュメントから文字列連結への変更

```powershell
# 修正前（問題のあるコード）
$debugInfo = @"
詳細情報:
- エラーコード: $errorCode
- エラーメッセージ: $message
- モジュール: $module
- タイムスタンプ: $timestamp
- PowerShellバージョン: $($PSVersionTable.PSVersion)
- OS: $($PSVersionTable.OS)
- 実行ユーザー: $env:USERNAME
- 実行パス: $PWD
"@

# 修正後
$psVersion = $PSVersionTable.PSVersion.ToString()
$osInfo = $PSVersionTable.OS
$userName = $env:USERNAME
$currentPath = $PWD.Path

$debugInfo = "詳細情報:`n" +
            "- エラーコード: $errorCode`n" +
            "- エラーメッセージ: $message`n" +
            "- モジュール: $module`n" +
            "- タイムスタンプ: $timestamp`n" +
            "- PowerShellバージョン: $psVersion`n" +
            "- OS: $osInfo`n" +
            "- 実行ユーザー: $userName`n" +
            "- 実行パス: $currentPath"
```

### 2. 変数の明示的宣言

```powershell
# 修正前
# 変数が直接ヒアドキュメント内で参照されている

# 修正後
# 変数を明示的に宣言してから使用
$psVersion = $PSVersionTable.PSVersion.ToString()
$osInfo = $PSVersionTable.OS
$userName = $env:USERNAME
$currentPath = $PWD.Path
```

## 修正後の確認

修正後、以下のコマンドで動作確認を行ってください：

```powershell
# Common.ps1の読み込みテスト
. ".\_lib\Common.ps1"

# Commonクラスのテスト
$testCommon = [Common]::new()
$testCommon.WriteLog("テストメッセージ", "INFO")
$testCommon.HandleError("TEST001", "テストエラー", "TestModule")
```

## バックアップ

修正前に、元のファイルは自動的にバックアップされます。
バックアップディレクトリは `backup_YYYYMMDD_HHMMSS` の形式で作成されます。

## トラブルシューティング

### 修正スクリプトが失敗する場合

1. PowerShellを管理者として実行
2. 実行ポリシーを確認
   ```powershell
   Get-ExecutionPolicy
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

### 手動修正が必要な場合

1. バックアップから元のファイルを復元
2. 上記の手動修正手順に従って修正を適用
3. 修正後のテストを実行

### その他のエラーが発生する場合

1. エラーログを確認（`Common_Error.log`ファイル）
2. PowerShellバージョンを確認
3. 必要なモジュールが読み込まれているか確認

## 注意事項

- 修正後は、Commonライブラリが正常に動作するようになります
- エラーログは `_lib` ディレクトリ内に作成されます
- ヒアドキュメントの代わりに文字列連結を使用することで、変数スコープの問題を解決しています
- この修正により、他のドライバークラスでも同様の問題が解決されます 