@echo off
setlocal

REM ===== Python launcher (use py; for Anaconda, set absolute python.exe path) =====
set PY=py

REM Batch file folder is the project root
set BASE_DIR=%~dp0
set SCRIPTS=%BASE_DIR%python_scripts

echo [1/2] Checking Python...
%PY% --version
if errorlevel 1 (
    echo.
    echo [ERROR] Python not found. Install Python first, then re-run.
    pause
    exit /b 1
)

echo.
echo [2/2] Installing libraries (pandas, openpyxl)...
%PY% -m pip install --upgrade pip
%PY% -m pip install pandas openpyxl

echo.
echo Install complete. Press any key to close.
pause
endlocal
