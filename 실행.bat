@echo off
chcp 65001 > nul
echo [Excel Part Number Matcher] 시작 중...
echo.

set VENV_PATH=%~dp0.venv
set PIP_CMD="%VENV_PATH%\Scripts\pip.exe"
set PYTHON_CMD="%VENV_PATH%\Scripts\pythonw.exe"

:: 필요 패키지 설치 확인
%PIP_CMD% show openpyxl >nul 2>&1
if errorlevel 1 (
    echo openpyxl 설치 중...
    %PIP_CMD% install openpyxl
)

%PIP_CMD% show tkinterdnd2 >nul 2>&1
if errorlevel 1 (
    echo tkinterdnd2 설치 중...
    %PIP_CMD% install tkinterdnd2
)

start "" /B %PYTHON_CMD% "%~dp0main.py"
