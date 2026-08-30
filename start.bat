@echo off
cd /d "%~dp0"
set "PY=py -3"
where py >nul 2>&1 || set "PY=python"
if not exist ".venv" (
    echo Tworzenie srodowiska .venv ...
    %PY% -m venv .venv
)
call ".venv\Scripts\activate.bat"
python -m pip install -q -r requirements.txt
python xlsx2xml.py
if errorlevel 1 pause
