@echo off
setlocal
cd /d "%~dp0"
title LocalGPT
color 0A
echo ==========================================
echo        LocalGPT - One Click Start
 echo ==========================================
where python >nul 2>nul
if errorlevel 1 (
  echo Python 3 is not installed or not in PATH.
  pause
  exit /b 1
)
echo Checking Ollama...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -UseBasicParsing http://127.0.0.1:11434/api/tags -TimeoutSec 2 | Out-Null; Write-Host 'Ollama is already running.' } catch { try { Start-Process ollama -ArgumentList 'serve'; Write-Host 'Started Ollama.' } catch { Write-Host 'Ollama was not found. Start Ollama manually.' } }"
timeout /t 2 /nobreak >nul
echo Starting LocalGPT...
python localgpt_server.py
pause
