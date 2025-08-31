# 🚀 PowerShell Driver Classes - Ollama セットアップガイド

## 📋 概要

このドキュメントは、PowerShell Driver ClassesのチャットボットでローカルLLM（Ollama）を使用するためのセットアップガイドです。

## 🎯 特徴

- **完全オフライン動作**: インターネット接続不要
- **ポータブル**: インストール作業不要
- **簡単起動**: ワンクリックでOllamaサービス開始
- **自動設定**: 環境変数の自動設定

## 📁 必要なファイル構成

```
_docs/
├── index.html              # メインドキュメント
├── start_ollama.bat        # Windows起動スクリプト（推奨）
├── start_ollama.ps1        # PowerShell起動スクリプト
├── README_OLLAMA.md        # このファイル
└── ollama/                 # Ollama実行ファイルとモデル
    ├── ollama.exe          # Ollama実行ファイル
    └── models/             # モデルファイル
        ├── llama2:7b/
        └── ...
```

## 🚀 使用方法

### 方法1: バッチファイルで起動（推奨）

1. **`start_ollama.bat` をダブルクリック**
2. コマンドプロンプトが開き、Ollamaが起動します
3. 「✅ Ollamaが正常に起動しました！」と表示されたら成功
4. このウィンドウは開いたままにしてください

### 方法2: PowerShellで起動

1. **`start_ollama.ps1` を右クリック → 「PowerShellで実行」**
2. 実行ポリシーの変更が必要な場合は以下を実行：
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## 🌐 チャットボットの使用

1. **Ollamaが起動したら、`index.html` を開く**
2. 画面右下の 💬 ボタンをクリック
3. チャットボットが開きます
4. 設定パネル（⚙️）で以下を確認：
   - ✅ ローカルLLMを有効にする
   - LLMタイプ: `ollama`
   - エンドポイント: `http://localhost:11434`
   - モデル: `llama2:7b-chat-q4_0` など

## 📚 利用可能なモデル

### 推奨モデル

- **llama2:7b-chat-q4_0**: 軽量で高速、日本語対応
- **codellama:7b-instruct-q4_0**: コード生成に特化
- **llama2:13b-chat-q4_0**: より高精度（容量大）

### モデルの追加

```bash
# コマンドプロンプトで実行
cd ollama
ollama pull llama2:7b-chat-q4_0
ollama pull codellama:7b-instruct-q4_0
```

## ⚠️ 注意事項

### システム要件

- **OS**: Windows 10/11
- **メモリ**: 最低8GB（推奨16GB以上）
- **ストレージ**: モデルサイズ + 5GB程度
- **管理者権限**: 初回起動時のみ必要

### トラブルシューティング

#### Ollamaが起動しない場合

1. **ファイアウォールの確認**
   - Windows Defenderファイアウォールでポート11434を許可
   
2. **ポートの競合確認**
   - 他のアプリがポート11434を使用していないか確認
   
3. **管理者権限での実行**
   - 起動スクリプトを右クリック → 「管理者として実行」

#### モデルが見つからない場合

1. **モデルのダウンロード確認**
   ```bash
   cd ollama
   ollama list
   ```

2. **モデルの再ダウンロード**
   ```bash
   ollama pull llama2:7b-chat-q4_0
   ```

## 🔧 カスタマイズ

### 環境変数の変更

起動スクリプト内の以下を編集：

```batch
set OLLAMA_DIR=%CURRENT_DIR%ollama
set MODELS_DIR=%OLLAMA_DIR%\models
```

### ポート番号の変更

1. 起動スクリプトでポート番号を変更
2. チャットボットの設定でエンドポイントを更新

## 📞 サポート

問題が発生した場合は、以下を確認してください：

1. **ログの確認**: 起動スクリプトの出力メッセージ
2. **システム要件**: メモリとストレージの空き容量
3. **権限**: 管理者権限での実行
4. **ファイアウォール**: ポート11434の通信許可

## 🎉 完了

これで、PowerShell Driver Classesのチャットボットが完全オフラインで動作するようになります！

- インターネット接続不要
- プライバシー保護
- 高速レスポンス
- カスタマイズ可能

チャットボットでPowerShell Driver Classesについて何でも質問してください！
