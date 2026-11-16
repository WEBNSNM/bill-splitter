@echo off
cd /d "%~dp0"
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m unittest discover -s tests -v
if errorlevel 1 goto failed
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name wuli --icon assets/juan.ico --add-data "assets/juan.ico;assets" app.py
if errorlevel 1 goto failed
echo EXE: %cd%\dist\wuli.exe
pause
exit /b 0
:failed
echo Build failed.
pause
exit /b 1
