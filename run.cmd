@ECHO OFF

SET START_TIME=%TIME%

PUSHD %~dp0
REM powershell -ExecutionPolicy Bypass %~dp0\_ps\main.ps1
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_EdgeDriver_AllWebDriverMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_WordDriver_AllMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_ExcelDriver_AllMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_PowerPointDriver_AllMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file %~dp0\_ps\hoge.ps1
powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\generate_SampleReport.ps1"
POPD

REM "C:\000_common\PowerShell\prototype\run.cmd"
SET END_TIME=%TIME%

ECHO 開始時間：%START_TIME%
ECHO 終了時間：%END_TIME%

PAUSE

