# バグレポート

## 発見されたバグと問題点

### 1. WebDriver.ps1 の問題

#### 1.1 FindElementsメソッドのバグ
**場所**: 行 1028-1050
**問題**: `FindElements`メソッドの戻り値処理に論理エラーがあります。

```powershell
if ($response_json.result.result.objectIds.Count -eq 0)
{
    return @{ nodeId = $response_json.result.result.objectIds; selector = $selector }
}
else
{
    throw "CSSセレクタで複数の要素を取得できません。セレクタ：$selector"
}
```

**修正案**:
```powershell
if ($response_json.result.result.objectIds.Count -gt 0)
{
    return @{ nodeId = $response_json.result.result.objectIds; selector = $selector }
}
else
{
    throw "CSSセレクタで複数の要素を取得できません。セレクタ：$selector"
}
```

#### 1.2 FindElementByXPathメソッドの未実装
**場所**: 行 1023
**問題**: `FindElementByXPath`メソッドがコメントアウトされており、実装されていません。

**修正案**: 以下のメソッドを実装する必要があります。
```powershell
# XPathで要素を検索
[hashtable] FindElementByXPath([string]$xpath)
{
    $expression = "document.evaluate('$xpath', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    return $this.FindElementGeneric($expression, 'XPath', $xpath)
}
```

#### 1.3 複数要素検索メソッドの未実装
**場所**: 行 1024-1027
**問題**: 以下のメソッドが実装されていません：
- `FindElementByClassName`
- `FindElementById`
- `FindElementByName`
- `FindElementByTagName`

#### 1.4 Disposeメソッドの呼び出しエラー
**場所**: ChromeDriver.ps1 行 247, EdgeDriver.ps1 行 243
**問題**: 親クラスのDisposeメソッドの呼び出し方が間違っています。

```powershell
[WebDriver]::Dispose($this)  # 間違い
```

**修正案**:
```powershell
[WebDriver]::Dispose.Invoke($this)  # 正しい呼び出し方
```

### 2. エラーハンドリングの問題

#### 2.1 エラーモジュールの依存関係
**場所**: 各ドライバーファイル
**問題**: エラーモジュール（.psm1ファイル）が存在しない可能性があります。

**必要なファイル**:
- `ChromeDriverErrors.psm1`
- `EdgeDriverErrors.psm1`
- `WordDriverErrors.psm1`

#### 2.2 エラーログの出力先
**場所**: WebDriverErrors.psm1 行 348
**問題**: エラーログファイルのパスが相対パスで指定されており、実行場所によっては書き込みに失敗する可能性があります。

**修正案**:
```powershell
$log_file = Join-Path $PSScriptRoot "WebDriver_Error.log"
```

### 3. WordDriver.ps1 の問題

#### 3.1 アセンブリの依存関係
**場所**: 行 3
**問題**: `Microsoft.Office.Interop.Word`アセンブリがシステムにインストールされていない場合、エラーが発生します。

**修正案**: アセンブリの存在確認を追加
```powershell
try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
} catch {
    throw "Microsoft.Office.Interop.Wordアセンブリが見つかりません。Microsoft Officeがインストールされているか確認してください。"
}
```

#### 3.2 テーブルデータの型変換
**場所**: 行 200-220
**問題**: `AddTable`メソッドで、テーブルデータの型変換が適切に行われていない可能性があります。

### 4. パフォーマンスの問題

#### 4.1 WebSocket接続のタイムアウト
**場所**: WebDriver.ps1 行 200-250
**問題**: WebSocket接続のタイムアウト設定が短すぎる可能性があります。

#### 4.2 メモリリークの可能性
**場所**: 各Disposeメソッド
**問題**: リソースの解放が不完全な場合があります。

### 5. セキュリティの問題

#### 5.1 ファイルパスの検証不足
**場所**: 各ファイル操作メソッド
**問題**: ファイルパスの検証が不十分で、パストラバーサル攻撃の可能性があります。

### 6. 互換性の問題

#### 6.1 PowerShellバージョン依存
**場所**: 全ファイル
**問題**: PowerShell 5.1以降の機能を使用しているため、古いバージョンでは動作しません。

#### 6.2 ブラウザバージョン依存
**場所**: ChromeDriver.ps1, EdgeDriver.ps1
**問題**: 特定のブラウザバージョンでのみ動作する可能性があります。

## 推奨される修正手順

1. **緊急修正**:
   - FindElementsメソッドの論理エラー修正
   - Disposeメソッドの呼び出し修正
   - エラーモジュールの作成

2. **重要修正**:
   - 未実装メソッドの実装
   - エラーハンドリングの改善
   - パフォーマンス最適化

3. **推奨修正**:
   - セキュリティ強化
   - 互換性向上
   - ドキュメント整備

## テスト推奨事項

1. 各ブラウザでの動作確認
2. エラーケースのテスト
3. パフォーマンステスト
4. セキュリティテスト
5. 異なるPowerShellバージョンでのテスト 