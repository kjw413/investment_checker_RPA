@echo off
chcp 65001 >nul
setlocal

REM ===== Conda Python 경로 (본인 환경에 맞게 수정) =====
set PY=C:\anaconda3\python.exe

REM 현재 배치파일 폴더 기준으로 상위 폴더가 BASE_DIR
set BASE_DIR=%~dp0..
set SCRIPTS=%BASE_DIR%\python_scripts
set REQ=%SCRIPTS%\requirements.txt

echo [1/2] Python 확인 중...

echo [2/2] 라이브러리 설치(pandas, openpyxl)...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r "%REQ%"

echo.
echo 설치 완료. 아무 키나 누르면 닫힙니다.
pause
endlocal
