@echo off
rem Uruchamia program bez okna konsoli. Jesli brakuje srodowiska - odpala najpierw install.bat
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
    call "%~dp0install.bat" -cichy
    if errorlevel 1 exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" "xlsx2xml.py"
exit /b 0
