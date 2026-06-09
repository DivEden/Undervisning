@echo off
cd /d "%~dp0"
start "" "http://127.0.0.1:5000"
.venv\Scripts\python.exe -m flask --app webapp/app.py run --port 5000
pause
