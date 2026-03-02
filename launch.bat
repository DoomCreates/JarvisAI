@echo off
title J.A.R.V.I.S.
cd /d "%~dp0"
color 0B
echo.
echo  ================================================================
echo    J.A.R.V.I.S.  ^|  Initializing...
echo  ================================================================
echo.
echo  Voice    : ElevenLabs (Adam)
echo  Brain    : Groq / Llama-3.3-70b
echo  Hotkey   : ` (backtick) to activate mic
echo  Silence  : 1.5s pause ends recording
echo  Server   : http://127.0.0.1:5000
echo.
echo  Keep this window open. Press Ctrl+C to shut down.
echo  ================================================================
echo.

findstr /C:"your_groq_api_key_here" .env >nul 2>&1
if not errorlevel 1 ( echo  [!] Groq API key not set in .env! & echo. )
findstr /C:"your_elevenlabs_api_key_here" .env >nul 2>&1
if not errorlevel 1 ( echo  [!] ElevenLabs API key not set in .env! & echo. )

call venv\Scripts\activate.bat
start /b cmd /c "timeout /t 3 /nobreak >nul && start http://127.0.0.1:5000"
python app.py
pause
