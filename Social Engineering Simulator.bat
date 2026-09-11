@echo off
setlocal
cd /d "%~dp0Application"

REM Start the Flask application in a hidden PowerShell process and open the browser.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -WindowStyle Hidden -FilePath 'pythonw.exe' -ArgumentList 'app.py' -WorkingDirectory '%CD%'"

REM Give Flask a moment to start.
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:5000"

endlocal
