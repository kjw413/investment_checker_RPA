@echo off
setlocal

REM ===== Python launcher (use py; for Anaconda, set absolute python.exe path) =====
set PY=py

REM Batch file folder is the project root
set BASE_DIR=%~dp0
set SCRIPTS=%BASE_DIR%python_scripts
set SCRIPT=%SCRIPTS%\investment_validator.py

if not exist "%SCRIPT%" (
    echo [ERROR] Script file not found:
    echo %SCRIPT%
    pause
    exit /b 1
)

echo Running investment validator...
%PY% "%SCRIPT%"

echo.
echo Done. Check the result folder:
echo %BASE_DIR%result
pause
endlocal
