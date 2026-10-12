@echo off
rem Instalacja: tworzy srodowisko .venv i instaluje wymagane paczki.
rem Uruchom raz. Program uruchamia sie potem przez program_xlsx-to-xml.bat
cd /d "%~dp0"

set "PY=py -3"
where py >nul 2>&1 || set "PY=python"
%PY% --version >nul 2>&1 || goto :blad
%PY% -c "import sys; sys.exit(sys.version_info < (3, 11))"
if errorlevel 1 (
    echo.
    echo UWAGA: ta wersja Pythona jest przestarzala i nie dostaje juz poprawek
    echo bezpieczenstwa. Zainstaluj aktualnego Pythona ^(3.11 lub nowszego^) z
    echo https://www.python.org/downloads/ i uruchom install.bat ponownie.
    echo Nacisnij dowolny klawisz, aby mimo to kontynuowac.
    pause >nul
)

if not exist ".venv\Scripts\pythonw.exe" (
    echo Tworze srodowisko wirtualne .venv ...
    %PY% -m venv .venv || goto :blad
)

rem Instaluj paczki tylko gdy requirements.txt sie zmienil
fc /b "requirements.txt" ".venv\requirements.lock" >nul 2>&1 || (
    echo Instaluje wymagane paczki ...
    ".venv\Scripts\python.exe" -m pip install -q -r requirements.txt || goto :blad
    copy /y "requirements.txt" ".venv\requirements.lock" >nul
)

echo.
echo Gotowe. Uruchom program plikiem: program_xlsx-to-xml.bat
if /i "%~1"=="-cichy" exit /b 0
pause
exit /b 0

:blad
echo.
echo Cos poszlo nie tak. Sprawdz, czy Python 3.11+ jest zainstalowany i dodany do PATH.
pause
exit /b 1
