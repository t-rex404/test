@ECHO OFF

chcp 65001 >nul
SET START_TIME=%TIME%

PUSHD %~dp0

REM powershell -ExecutionPolicy Bypass %~dp0\_ps\main.ps1
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_EdgeDriver_AllWebDriverMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_WordDriver_AllMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_ExcelDriver_AllMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_PowerPointDriver_AllMethods.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file %~dp0\_ps\hoge.ps1
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\generate_SampleReport.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_UIAutomationDriver_Calculator_Notepad.ps1"
REM powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_GUIDriver_Calculator_Notepad.ps1"
powershell.exe -ExecutionPolicy Bypass -file "%~dp0\_ps\test_UIAutomationDriver_AllMethods.ps1"

POPD

SET END_TIME=%TIME%

ECHO Start time: %START_TIME%
ECHO End time: %END_TIME%

ECHO Press any key to exit...
PAUSE >nul

