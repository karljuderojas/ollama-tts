@echo off
cd /d "%~dp0"

echo Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not on PATH.
    echo Download it from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installing / updating dependencies...
python -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo ERROR: Failed to install dependencies from requirements.txt
    pause
    exit /b 1
)

python -m pip install faster-whisper -q
if errorlevel 1 (
    echo WARNING: Could not install faster-whisper. Mic input will be unavailable.
)

echo Checking Ollama...
ollama list >nul 2>&1
if errorlevel 1 (
    echo Ollama is not running. Starting it...
    start "" ollama serve
    timeout /t 3 /nobreak >nul
)

echo Starting Ollama Voice Reader...
python main.py
if errorlevel 1 (
    echo.
    echo Error: App exited with an error. Check the output above for details.
    pause
)
