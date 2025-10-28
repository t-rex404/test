@echo off
rem このファイルの位置を作業ディレクトリに
cd /d %~dp0
REM 設定ファイル python\_json\config.json または 環境変数 FEEDBACK_CSV_SOURCE を参照
REM 監視モードで1秒間隔監視
.\python-3.14.0-embed-amd64\python.exe main.py --watch --interval 1

pause
