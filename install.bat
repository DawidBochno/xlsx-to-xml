@echo off
rem Uruchamia program bez okna konsoli (pythonw.exe).
rem Przy pierwszym starcie tworzy .venv i instaluje wymagane paczki.
cd /d "%~dp0"

set "PY=py -3"
where py >nul 2>&1 || set "PY=python"

if not exist ".venv\Scripts\pythonw.exe" (
    echo Pierwsze uruchomienie - przygotowuje srodowisko, chwile to potrwa...
    %PY% -m venv .venv || goto :blad
)

rem Instaluj paczki tylko gdy requirements.txt sie zmienil
fc /b "requirements.txt" ".venv\requirements.lock" >nul 2>&1 || (
    echo Instaluje wymagane paczki...
    ".venv\Scripts\python.exe" -m pip install -q -r requirements.txt || goto :blad
    copy /y "requirements.txt" ".venv\requirements.lock" >nul
)

start "" ".venv\Scripts\pythonw.exe" "xlsx2xml.py"
exit /b 0

:blad
echo.
echo Cos poszlo nie tak. Sprawdz, czy Python 3.9+ jest zainstalowany i dodany do PATH.
pause
exit /b 1
