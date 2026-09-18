@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" -c "import pandas, openpyxl, tkinter" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
        goto uruchom_exe
    )
)

where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import pandas, openpyxl, tkinter" >nul 2>&1
    if not errorlevel 1 goto uruchom_py
)

where python >nul 2>&1
if not errorlevel 1 (
    python -c "import pandas, openpyxl, tkinter" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=python"
        goto uruchom_exe
    )
)

if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
    "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" -c "import pandas, openpyxl, tkinter" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
        goto uruchom_exe
    )
)

echo.
echo Nie znaleziono dzialajacego Pythona z bibliotekami pandas i openpyxl.
echo Zainstaluj Python 3, a nastepnie wykonaj:
echo     py -m pip install pandas openpyxl
echo.
pause
exit /b 1

:uruchom_py
py -3 "%~dp0laczenie_zamowien_formatki.py"
goto sprawdz_wynik

:uruchom_exe
"%PYTHON_EXE%" "%~dp0laczenie_zamowien_formatki.py"

:sprawdz_wynik
set "KOD_WYJSCIA=%ERRORLEVEL%"
if not "%KOD_WYJSCIA%"=="0" (
    echo.
    echo Program zakonczyl sie bledem. Szczegoly sa widoczne powyzej.
    pause
)
exit /b %KOD_WYJSCIA%
