# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

このプロジェクトは、**埋め込み型Python 3.14.0**ディストリビューション（`python-3.14.0-embed-amd64`）を使用したPythonプロトタイプ環境です。埋め込み型Pythonパッケージは、システム全体へのインストールが不要な最小限の自己完結型Python環境で、Windows上での可搬性を重視した設計になっています。

## コードの実行

### メインスクリプトの実行
```cmd
run.cmd
```
このバッチファイルは作業ディレクトリを設定し、埋め込みPythonインタープリタを使用して`main.py`を実行します。

### Pythonの直接実行
```cmd
python-3.14.0-embed-amd64\python.exe main.py
```

### Pythonを対話モードで実行
```cmd
python-3.14.0-embed-amd64\python.exe
```

## パッケージ管理

### パッケージのインストール
```cmd
python-3.14.0-embed-amd64\python.exe -m pip install <パッケージ名>
```

### インストール済みパッケージの一覧表示
```cmd
python-3.14.0-embed-amd64\python.exe -m pip list
```

### パッケージのアンインストール
```cmd
python-3.14.0-embed-amd64\python.exe -m pip uninstall <パッケージ名>
```

## アーキテクチャに関する注意事項

### 埋め込み型Pythonの設定

埋め込み型Pythonディストリビューションは、`python314._pth`ファイルを使用してPythonパスを設定しています。このファイルでは現在`import site`がコメント解除されており、pip機能とsite-packagesのサポートが有効になっています。

主な特徴：
- **可搬性**: Python環境全体が`python-3.14.0-embed-amd64`ディレクトリ内で自己完結
- **システムPATH非依存**: インタープリタは相対パスまたは絶対パスで呼び出す必要がある
- **site packages有効化**: サードパーティパッケージはpip経由で`python-3.14.0-embed-amd64\Lib\site-packages`にインストール可能

### インストール済みの依存関係

環境には、データサイエンスとWeb開発向けのいくつかのパッケージがプリインストールされています：
- **データサイエンス**: pandas, numpy, matplotlib
- **Webフレームワーク**: Django, Flask
- **ユーティリティ**: colorama, click, Jinja2

### プロジェクト構成

- `main.py` - Flaskベースのタスク管理Webアプリケーション（http://localhost:5000）
- `test.py` - PyScriptを使用したコーヒーメニュー焙煎シミュレーターのHTMLプロトタイプ
- `run.cmd` - 埋め込みPythonインタープリタでmain.pyを実行するバッチファイル
- `python-3.14.0-embed-amd64/` - 埋め込み型Pythonディストリビューションのディレクトリ
  - `python.exe` - Pythonインタープリタ
  - `Lib/site-packages/` - インストールされたサードパーティパッケージ
  - `python314._pth` - Pythonパス設定ファイル

## 開発時の考慮事項

このコードベースで開発する際は：
- 常に埋め込みPythonインタープリタのパスを使用すること: `python-3.14.0-embed-amd64\python.exe`
- これはWindows専用の環境です（埋め込み型PythonディストリビューションはWindows専用）
- この環境はPython 3.14.0（最新/開発版）を使用しています
