@echo off
title PowerShell Driver Classes - Ollama Startup Script

echo.
echo ========================================
echo PowerShell Driver Classes - Ollama Startup
echo ========================================
echo.

:: Get current directory
set CURRENT_DIR=%~dp0
set OLLAMA_DIR=%CURRENT_DIR%ollama
set MODELS_DIR=%OLLAMA_DIR%\models

:: Check if Ollama directory exists
if not exist "%OLLAMA_DIR%" (
    echo [ERROR] Ollama directory not found
    echo Directory: %OLLAMA_DIR%
    echo.
    echo Please check if ollama folder is included in this document folder.
    echo.
    pause
    exit /b 1
)

:: Check if Ollama executable exists
if not exist "%OLLAMA_DIR%\ollama.exe" (
    echo [ERROR] Ollama executable not found
    echo File: %OLLAMA_DIR%\ollama.exe
    echo.
    echo Please check if ollama.exe is included in ollama folder.
    echo.
    pause
    exit /b 1
)

:: Set environment variables
set PATH=%OLLAMA_DIR%;%PATH%
set OLLAMA_HOME=%OLLAMA_DIR%
set OLLAMA_MODELS=%MODELS_DIR%

echo [OK] Ollama environment configured
echo Executable: %OLLAMA_DIR%\ollama.exe
echo Models directory: %MODELS_DIR%
echo.

:: Terminate existing Ollama processes
echo [INFO] Checking existing Ollama processes...
taskkill /f /im ollama.exe >nul 2>&1
timeout /t 2 >nul

:: Start Ollama
echo [INFO] Starting Ollama...
start /min cmd /c ""%OLLAMA_DIR%\ollama.exe" serve"

:: Wait for startup
echo [INFO] Waiting for Ollama to start...
timeout /t 5 >nul

:: Connection test
echo [INFO] Testing connection...
curl -s http://localhost:11434/api/tags >nul 2>&1
if %errorlevel% equ 0 (
    echo [SUCCESS] Ollama started successfully!
    echo.
    echo [INFO] Local endpoint: http://localhost:11434
    echo [INFO] Checking available models...
    echo.
    
    :: Show available models
    "%OLLAMA_DIR%\ollama.exe" list
    
    echo.
    echo [SUCCESS] Chatbot is now available!
    echo Open index.html to try the chatbot.
    echo.
    echo [INFO] Keep this window open while using the chatbot.
    echo.
) else (
    echo [ERROR] Failed to start Ollama
    echo.
    echo Troubleshooting:
    echo 1. Check firewall settings
    echo 2. Check if port 11434 is not used by other apps
    echo 3. Try running with administrator privileges
    echo.
)

echo.
echo Keep this window open.
echo Ollama service is running.
echo.
pause
