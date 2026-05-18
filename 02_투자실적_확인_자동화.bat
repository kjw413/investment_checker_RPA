@echo off
chcp 65001 >nul
setlocal

REM ===== Conda Python 경로 (본인 환경에 맞게 수정) =====
set PY=C:\anaconda3\python.exe

REM ===== E드라이브 기준 경로 =====
set BASE_DIR=%~dp0..
set SCRIPTS=%BASE_DIR%\python_scripts
set SCRIPT=%SCRIPTS%\InvestmentCheck.py

echo 투자 실적 점검 실행 중...
%PY% "%SCRIPT%"

echo.
echo 완료되었습니다. result 폴더를 확인하세요:
echo %BASE_DIR%\result
pause
endlocal
