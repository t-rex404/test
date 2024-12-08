@echo off

echo %~dp0

pushd %~dp0

python %~dp0\_main.py

pause