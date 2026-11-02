@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 goto missing
python -m venv .venv
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Installation completed. Run start.bat.
pause
exit /b 0
:missing
echo Install Python 3.11+ from python.org and select Add Python to PATH.
pause
exit /b 1
:failed
echo Installation failed. Check Python version and internet connection.
pause
exit /b 1
